VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmCLI 
   BackColor       =   &H00FAEFDA&
   Caption         =   "Clientes / Proveedores"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   1080
   ClientWidth     =   9480
   ControlBox      =   0   'False
   Icon            =   "client.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5805
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.Frame F14 
      BackColor       =   &H00FAEFDA&
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
      Left            =   45
      TabIndex        =   4
      Top             =   8265
      Visible         =   0   'False
      Width           =   4125
      Begin MSFlexGridLib.MSFlexGrid ListExiste 
         Height          =   1455
         Left            =   120
         TabIndex        =   5
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton Op 
         Caption         =   "Ignorar la Lista "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   1335
      End
      Begin VB.OptionButton Op 
         Caption         =   "Seleccionar uno de la Lista "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   2535
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   315
      TabIndex        =   1
      Top             =   8325
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
   Begin TabDlg.SSTab SSTab 
      Height          =   3960
      Left            =   360
      TabIndex        =   16
      Top             =   3240
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   6985
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      BackColor       =   16445402
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos de Domicilio"
      TabPicture(0)   =   "client.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Direccion de Trabajo"
      TabPicture(1)   =   "client.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Otros Datos"
      TabPicture(2)   =   "client.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame fra2 
         BackColor       =   &H00FAEFDA&
         Height          =   3600
         Left            =   -74970
         TabIndex        =   53
         Top             =   315
         Width           =   8145
         Begin VB.CommandButton cmdmante 
            Caption         =   "Editar &Placas"
            Height          =   510
            Left            =   7080
            TabIndex        =   72
            Top             =   2625
            Width           =   960
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
            Left            =   5505
            MaxLength       =   4
            TabIndex        =   71
            Top             =   465
            Visible         =   0   'False
            WhatsThisHelpID =   3
            Width           =   615
         End
         Begin VB.CheckBox otrocontrato 
            BackColor       =   &H00FAEFDA&
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
            Height          =   270
            Left            =   6285
            TabIndex        =   70
            Top             =   1440
            Visible         =   0   'False
            WhatsThisHelpID =   7
            Width           =   375
         End
         Begin VB.CheckBox letraotorgado 
            BackColor       =   &H00FAEFDA&
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
            Height          =   240
            Left            =   6285
            TabIndex        =   69
            Top             =   1740
            Visible         =   0   'False
            WhatsThisHelpID =   10
            Width           =   345
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
            TabIndex        =   68
            Top             =   490
            Visible         =   0   'False
            WhatsThisHelpID =   1
            Width           =   3375
         End
         Begin VB.ComboBox TxtZonaTrabajo 
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   1080
            Visible         =   0   'False
            WhatsThisHelpID =   5
            Width           =   3375
         End
         Begin VB.ComboBox TxtSubZonaTrabajo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   3585
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   1065
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
            Left            =   4950
            MaxLength       =   30
            TabIndex        =   65
            Top             =   2865
            Visible         =   0   'False
            WhatsThisHelpID =   17
            Width           =   1980
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
            Left            =   135
            MaxLength       =   30
            TabIndex        =   64
            Top             =   1700
            Visible         =   0   'False
            WhatsThisHelpID =   8
            Width           =   2475
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
            TabIndex        =   63
            Top             =   2290
            Visible         =   0   'False
            WhatsThisHelpID =   11
            Width           =   2475
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
            Left            =   2760
            MaxLength       =   15
            TabIndex        =   62
            Top             =   2298
            Visible         =   0   'False
            WhatsThisHelpID =   12
            Width           =   1980
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
            Left            =   4920
            MaxLength       =   20
            TabIndex        =   61
            Top             =   2250
            Visible         =   0   'False
            WhatsThisHelpID =   16
            Width           =   1995
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
            TabIndex        =   60
            Top             =   2880
            Visible         =   0   'False
            WhatsThisHelpID =   14
            Width           =   2475
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
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   59
            Top             =   2886
            Visible         =   0   'False
            WhatsThisHelpID =   15
            Width           =   1980
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
            Left            =   5565
            MaxLength       =   12
            TabIndex        =   58
            Top             =   1080
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
            Left            =   7035
            MaxLength       =   2
            TabIndex        =   57
            Top             =   480
            Visible         =   0   'False
            WhatsThisHelpID =   13
            Width           =   975
         End
         Begin VB.ComboBox TxtLugarTrab 
            DataSource      =   "2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   3615
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   474
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
            Left            =   6915
            MaxLength       =   1
            TabIndex        =   55
            Top             =   1110
            Visible         =   0   'False
            WhatsThisHelpID =   19
            Width           =   780
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
            Left            =   2760
            MaxLength       =   15
            TabIndex        =   54
            Top             =   1710
            Visible         =   0   'False
            WhatsThisHelpID =   9
            Width           =   1980
         End
         Begin VB.Label lblnom 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº. Dir."
            DataSource      =   "3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   5505
            TabIndex        =   90
            Tag             =   "16"
            Top             =   210
            Visible         =   0   'False
            WhatsThisHelpID =   3
            Width           =   555
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección Trabajo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   89
            Tag             =   "15"
            Top             =   240
            Visible         =   0   'False
            WhatsThisHelpID =   1
            Width           =   1245
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Provincia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   18
            Left            =   3585
            TabIndex        =   88
            Tag             =   "19"
            Top             =   840
            Visible         =   0   'False
            WhatsThisHelpID =   6
            Width           =   645
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Distrito"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   120
            TabIndex        =   87
            Tag             =   "18"
            Top             =   830
            Visible         =   0   'False
            WhatsThisHelpID =   5
            Width           =   510
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prendas :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   26
            Left            =   4950
            TabIndex        =   86
            Tag             =   "27"
            Top             =   2610
            Visible         =   0   'False
            WhatsThisHelpID =   17
            Width           =   690
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prop. (1) :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   85
            Tag             =   "20"
            Top             =   1450
            Visible         =   0   'False
            WhatsThisHelpID =   8
            Width           =   750
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prop. (2) :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   84
            Tag             =   "22"
            Top             =   2040
            Visible         =   0   'False
            WhatsThisHelpID =   11
            Width           =   750
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rg.Pub.(1)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   20
            Left            =   2760
            TabIndex        =   83
            Tag             =   "21"
            Top             =   1455
            Visible         =   0   'False
            WhatsThisHelpID =   9
            Width           =   795
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rg.Pub.(2)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   22
            Left            =   2760
            TabIndex        =   82
            Tag             =   "23"
            Top             =   2055
            Visible         =   0   'False
            WhatsThisHelpID =   12
            Width           =   795
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Autovaluo :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   25
            Left            =   4920
            TabIndex        =   81
            Tag             =   "26"
            Top             =   1995
            Visible         =   0   'False
            WhatsThisHelpID =   16
            Width           =   840
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Autos (1) :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   80
            Tag             =   "24"
            Top             =   2630
            Visible         =   0   'False
            WhatsThisHelpID =   14
            Width           =   780
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Autos (2) :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   24
            Left            =   2760
            TabIndex        =   79
            Tag             =   "25"
            Top             =   2640
            Visible         =   0   'False
            WhatsThisHelpID =   15
            Width           =   780
         End
         Begin VB.Label lblnom 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefono"
            DataSource      =   "4"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   5580
            TabIndex        =   78
            Tag             =   "17"
            Top             =   825
            Visible         =   0   'False
            WhatsThisHelpID =   4
            Width           =   645
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Letra Otorgado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   28
            Left            =   6795
            TabIndex        =   77
            Tag             =   "29"
            Top             =   1755
            Visible         =   0   'False
            WhatsThisHelpID =   10
            Width           =   1110
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Op 1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   29
            Left            =   7080
            TabIndex        =   76
            Tag             =   "30"
            Top             =   270
            Visible         =   0   'False
            WhatsThisHelpID =   13
            Width           =   345
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lugar"
            DataSource      =   "2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   32
            Left            =   3615
            TabIndex        =   75
            Tag             =   "33"
            Top             =   225
            Visible         =   0   'False
            WhatsThisHelpID =   2
            Width           =   405
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Op 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   33
            Left            =   7005
            TabIndex        =   74
            Tag             =   "34"
            Top             =   840
            Visible         =   0   'False
            WhatsThisHelpID =   19
            Width           =   345
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contrato a Plazo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   27
            Left            =   6750
            TabIndex        =   73
            Tag             =   "28"
            Top             =   1470
            Visible         =   0   'False
            WhatsThisHelpID =   7
            Width           =   1200
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FAEFDA&
         Height          =   3615
         Left            =   30
         TabIndex        =   37
         Top             =   315
         Width           =   8145
         Begin VB.Frame fraplaca 
            BackColor       =   &H00FAEFDA&
            Caption         =   "Descto. Especial"
            Height          =   2085
            Left            =   2490
            TabIndex        =   99
            Top             =   1425
            Visible         =   0   'False
            Width           =   3450
            Begin VB.CommandButton cmddescto 
               Caption         =   "&Editar Descuentos"
               Height          =   300
               Left            =   600
               TabIndex        =   100
               Top             =   1725
               Width           =   2535
            End
            Begin MSFlexGridLib.MSFlexGrid grid_des 
               Height          =   1380
               Left            =   90
               TabIndex        =   101
               Top             =   255
               Width           =   3225
               _ExtentX        =   5689
               _ExtentY        =   2434
               _Version        =   393216
               FixedCols       =   0
               BorderStyle     =   0
               Appearance      =   0
            End
         End
         Begin VB.TextBox t_diascred 
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
            Left            =   6285
            MaxLength       =   12
            TabIndex        =   94
            Top             =   1095
            Visible         =   0   'False
            WhatsThisHelpID =   4
            Width           =   1275
         End
         Begin VB.TextBox t_diasfac 
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
            Left            =   4515
            MaxLength       =   12
            TabIndex        =   93
            Top             =   1065
            Visible         =   0   'False
            WhatsThisHelpID =   4
            Width           =   1275
         End
         Begin VB.TextBox t_fechafac 
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
            Left            =   6285
            MaxLength       =   12
            TabIndex        =   92
            Top             =   480
            Visible         =   0   'False
            WhatsThisHelpID =   4
            Width           =   1275
         End
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
            Left            =   4515
            MaxLength       =   1
            TabIndex        =   91
            Top             =   450
            Visible         =   0   'False
            WhatsThisHelpID =   19
            Width           =   510
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
            Height          =   450
            ItemData        =   "client.frx":0496
            Left            =   180
            List            =   "client.frx":0498
            TabIndex        =   46
            Top             =   2925
            Width           =   2145
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
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   465
            Width           =   1215
         End
         Begin VB.CommandButton cmdcontab 
            Caption         =   "Relacionar a Contabilidad"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1770
            TabIndex        =   44
            Top             =   420
            Width           =   2265
         End
         Begin VB.CommandButton cmdcontab2 
            Caption         =   "Relacionar a Contabilidad"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1770
            TabIndex        =   43
            Top             =   930
            Width           =   2265
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
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   1065
            Width           =   1215
         End
         Begin VB.TextBox txtpordes 
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
            Left            =   180
            MaxLength       =   12
            TabIndex        =   41
            Text            =   " "
            Top             =   2295
            Width           =   1000
         End
         Begin VB.ComboBox txtestado 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "client.frx":049A
            Left            =   180
            List            =   "client.frx":04A4
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   1665
            Width           =   2130
         End
         Begin VB.CommandButton copia 
            Caption         =   "Copia a Otra Cia."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6075
            TabIndex        =   39
            Top             =   1590
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Crear clsal "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   6045
            TabIndex        =   38
            Top             =   2205
            Width           =   1860
         End
         Begin ComctlLib.ProgressBar PB2 
            Height          =   135
            Left            =   2220
            TabIndex        =   47
            Top             =   180
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   238
            _Version        =   327682
            Appearance      =   0
         End
         Begin VB.Label g_fechacred 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dias d' Cred."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   6285
            TabIndex        =   98
            Tag             =   "25"
            Top             =   825
            Visible         =   0   'False
            WhatsThisHelpID =   15
            Width           =   915
         End
         Begin VB.Label g_diasfac 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dias p' Factr."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   4515
            TabIndex        =   97
            Tag             =   "25"
            Top             =   810
            Visible         =   0   'False
            WhatsThisHelpID =   15
            Width           =   945
         End
         Begin VB.Label g_fechafac 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha p' Factr."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   6285
            TabIndex        =   96
            Tag             =   "25"
            Top             =   225
            Visible         =   0   'False
            WhatsThisHelpID =   15
            Width           =   1080
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Programado :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   34
            Left            =   4515
            TabIndex        =   95
            Tag             =   "35"
            Top             =   195
            Visible         =   0   'False
            WhatsThisHelpID =   13
            Width           =   975
         End
         Begin VB.Label LblDatos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bloqueos :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   20
            Left            =   180
            TabIndex        =   52
            Top             =   2655
            Width           =   750
         End
         Begin VB.Label lcuenta 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cta. Activo:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   180
            TabIndex        =   51
            Top             =   210
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cta. Naturaleza:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   180
            TabIndex        =   50
            Top             =   810
            Width           =   1200
         End
         Begin VB.Label lblnom 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descto. Facturación:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   35
            Left            =   180
            TabIndex        =   49
            Tag             =   "9"
            Top             =   2040
            Width           =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estado :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   48
            Top             =   1410
            Width           =   600
         End
      End
      Begin VB.Frame fra1 
         BackColor       =   &H00FAEFDA&
         Height          =   3600
         Left            =   -74955
         TabIndex        =   17
         Top             =   330
         Width           =   8130
         Begin VB.TextBox txtlimite 
            DataField       =   "ART_COSPRO"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6435
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   26
            Top             =   2445
            Width           =   1125
         End
         Begin VB.ComboBox TxtLugarCasa 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   3210
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1134
            Width           =   2760
         End
         Begin VB.ComboBox txtZonaNew 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   225
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1155
            Width           =   2760
         End
         Begin VB.ComboBox txtsubgrupo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   3210
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   2430
            Width           =   2760
         End
         Begin VB.ComboBox cmbgrupo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   225
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   2415
            Width           =   2760
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
            Left            =   225
            MaxLength       =   30
            TabIndex        =   21
            Top             =   555
            Width           =   4470
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
            Left            =   5190
            MaxLength       =   4
            TabIndex        =   20
            Top             =   525
            Width           =   735
         End
         Begin VB.ComboBox TxtZona 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   225
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1785
            Width           =   2760
         End
         Begin VB.ComboBox TxtSubZona 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   3210
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1782
            Width           =   2760
         End
         Begin ComctlLib.ProgressBar PB 
            Height          =   135
            Left            =   6285
            TabIndex        =   27
            Top             =   915
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   238
            _Version        =   327682
            Appearance      =   0
         End
         Begin VB.Label lbllimite 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite de Credito"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   6450
            TabIndex        =   36
            Top             =   2115
            Width           =   1200
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lugar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   31
            Left            =   3225
            TabIndex        =   35
            Tag             =   "32"
            Top             =   870
            Width           =   405
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Zona"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   225
            TabIndex        =   34
            Tag             =   "12"
            Top             =   900
            Width           =   360
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clase de Negocio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   3210
            TabIndex        =   33
            Tag             =   "14"
            Top             =   2166
            Width           =   1230
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Negocio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   32
            Tag             =   "13"
            Top             =   2160
            Width           =   1140
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Provincia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   3210
            TabIndex        =   31
            Tag             =   "11"
            Top             =   1518
            Width           =   645
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Distrito"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   225
            TabIndex        =   30
            Tag             =   "10"
            Top             =   1530
            Width           =   510
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   225
            TabIndex        =   29
            Tag             =   "7"
            Top             =   300
            Width           =   645
         End
         Begin VB.Label lblnom 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° Dir."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   5220
            TabIndex        =   28
            Tag             =   "8"
            Top             =   270
            Width           =   495
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FAEFDA&
      Height          =   7035
      Left            =   8910
      TabIndex        =   10
      Top             =   135
      Width           =   1995
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   15
         Top             =   615
         Width           =   1305
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   14
         Top             =   3555
         Width           =   1305
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   13
         Top             =   1595
         Width           =   1305
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Ce&rrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   12
         Top             =   6315
         Width           =   1305
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   2575
         Width           =   1305
      End
   End
   Begin VB.Timer Parpadea 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   240
      Top             =   6000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FAEFDA&
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
      Height          =   3105
      Left            =   345
      TabIndex        =   3
      Top             =   75
      Width           =   8235
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
         Left            =   2490
         MaxLength       =   12
         TabIndex        =   121
         Text            =   " "
         Top             =   2685
         Width           =   2010
      End
      Begin VB.OptionButton OptNombre 
         BackColor       =   &H00FAEFDA&
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
         Left            =   6375
         TabIndex        =   118
         Top             =   2340
         Width           =   240
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
         Left            =   2490
         MaxLength       =   40
         TabIndex        =   117
         Top             =   2295
         Width           =   3855
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
         Left            =   7110
         MaxLength       =   15
         TabIndex        =   116
         Top             =   2250
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
         Left            =   7080
         MaxLength       =   15
         TabIndex        =   114
         Top             =   1470
         Width           =   1000
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
         Left            =   2490
         MaxLength       =   15
         TabIndex        =   112
         Top             =   1890
         Width           =   1965
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
         Left            =   2490
         MaxLength       =   40
         TabIndex        =   110
         Top             =   1500
         Width           =   3855
      End
      Begin VB.OptionButton OptNombre 
         BackColor       =   &H00FAEFDA&
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
         Left            =   6360
         TabIndex        =   109
         Top             =   1515
         Width           =   300
      End
      Begin VB.OptionButton OptNombre 
         BackColor       =   &H00FAEFDA&
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
         Left            =   6360
         TabIndex        =   107
         Top             =   1125
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
         Left            =   2490
         MaxLength       =   40
         TabIndex        =   106
         Top             =   1110
         Width           =   3855
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
         Height          =   330
         Left            =   2490
         MaxLength       =   40
         TabIndex        =   103
         Top             =   675
         Width           =   5670
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
         Left            =   2490
         MaxLength       =   8
         TabIndex        =   102
         Top             =   285
         Width           =   1215
      End
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
         ItemData        =   "client.frx":04BB
         Left            =   5745
         List            =   "client.frx":04C5
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   2415
      End
      Begin VB.Label lblvar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccionar :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   4875
         TabIndex        =   123
         Top             =   345
         Width           =   915
      End
      Begin VB.Label lblnom 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   1500
         TabIndex        =   122
         Tag             =   "9"
         Top             =   2745
         Width           =   765
      End
      Begin VB.Label lblnom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conyuge :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   1515
         TabIndex        =   120
         Tag             =   "5"
         Top             =   2340
         Width           =   750
      End
      Begin VB.Label lblnom 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L.E."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   6675
         TabIndex        =   119
         Tag             =   "6"
         Top             =   2355
         Width           =   315
      End
      Begin VB.Label lblnom 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L.E."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   6690
         TabIndex        =   115
         Tag             =   "4"
         Top             =   1500
         Width           =   300
      End
      Begin VB.Label lblnom 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " -->RUC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   1650
         TabIndex        =   113
         Tag             =   "2"
         Top             =   1935
         Width           =   615
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gerente / Representate Legal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   111
         Tag             =   "3"
         Top             =   1530
         Width           =   2145
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre / Razon Social"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   660
         TabIndex        =   108
         Tag             =   "1"
         Top             =   1125
         Width           =   1605
      End
      Begin VB.Label lblvar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   1665
         TabIndex        =   105
         Top             =   315
         Width           =   600
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de Cliente :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   30
         Left            =   840
         TabIndex        =   104
         Tag             =   "31"
         Top             =   720
         Width           =   1425
      End
   End
   Begin VB.Label LblMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008B4914&
      Height          =   240
      Left            =   8310
      TabIndex        =   2
      Top             =   3240
      Width           =   75
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
Dim pasa As Integer
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
Dim COD_ORIGINAL As Currency
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
Dim PSPLAC_LLAVE As rdoQuery
Dim cliplac_llave As rdoResultset

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
        cont.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub
Public Sub LLENA_LISTAS(cont As ListBox, tip As Integer, WCODCLIE As Currency)
    PUB_TIPREG = tip
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
        WCODCLIE = tab_mayor!tab_nomlargo
        cont.AddItem tab_mayor!tab_nomlargo & String(80, " ") & tab_mayor!TAB_NUMTAB
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
        cont.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!TAB_NUMTAB
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
        cont.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!TAB_NUMTAB
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
        If cliloc_llave!CLI_TIPO_BLOQ1 = Left(tab_mayor!tab_nomcorto, 1) Then
            ListBloqueos.AddItem tab_mayor!tab_nomcorto
        End If
        If cliloc_llave!CLI_TIPO_BLOQ2 = Left(tab_mayor!tab_nomcorto, 1) Then
            ListBloqueos.AddItem tab_mayor!tab_nomcorto
        End If
        If cliloc_llave!CLI_TIPO_BLOQ3 = Left(tab_mayor!tab_nomcorto, 1) Then
            ListBloqueos.AddItem tab_mayor!tab_nomcorto
        End If
        If cliloc_llave!CLI_TIPO_BLOQ4 = Left(tab_mayor!tab_nomcorto, 1) Then
            ListBloqueos.AddItem tab_mayor!tab_nomcorto
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

    txtDTX.Enabled = False
    txtprog.Enabled = False
    cmdcontab.Enabled = False
    cmdcontab2.Enabled = False
    tcuenta2.Enabled = False
    frmCLI.txtpordes.Enabled = False
    g_fechafac.Enabled = False
    g_diasfac.Enabled = False
    t_fechafac.Enabled = False
    t_diasfac.Enabled = False
    t_diascred.Enabled = False
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

    txtDTX.Enabled = True
    txtprog.Enabled = True
    cmdcontab.Enabled = True
    cmdcontab2.Enabled = True
    tcuenta2.Enabled = True
    frmCLI.txtpordes.Enabled = True
    
    g_fechafac.Enabled = True
    g_diasfac.Enabled = True
    t_fechafac.Enabled = True
    t_diasfac.Enabled = True
    t_diascred.Enabled = True
End Sub



Private Sub CmbCGP_Click()
If llave1 <> "X" Then
  txt_key.Enabled = False
  If Trim(txtnombre.Text) <> "" Then
    LIMPIA_CLI
  End If
  CmbCGP_KeyPress 13
End If
End Sub

Private Sub CmbCGP_GotFocus()
If ListView1.Visible Then
 frmCLI.txt_key.Text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub CmbCGP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CmbCGP.Text = "" Then
       CmbCGP.SetFocus
       SendKeys "%{UP}"
       Exit Sub
    End If
'    ALLVISIBLE
    If Left(CmbCGP.Text, 1) = "P" Then
      ''frmCLI.SSTab1.TabCaption(0) = "&Datos Proveedor - Principales"
      ''frmCLI.SSTab1.TabCaption(1) = "&Datos Proveedor - Opcionales"
       LOC_TIPREG = 310 ' PROVEEDORES
       Screen.MousePointer = 11
       ETIQUETA_CLI
       Screen.MousePointer = 0
       lbllimite.Visible = False
       txtlimite.Visible = False
       lcuenta.Caption = "Cta. Pasivo:"
    Else
      lcuenta.Caption = "Cta. Activo:"
      ''frmCLI.SSTab1.TabCaption(0) = "&Datos Clientes - Principales"
      ''frmCLI.SSTab1.TabCaption(1) = "&Datos Clientes - Opcionales"
      LOC_TIPREG = 300 ' CLIENTES
      Screen.MousePointer = 11
      ETIQUETA_CLI
      lbllimite.Visible = True
      txtlimite.Visible = True
      Screen.MousePointer = 0
    End If
      If Left(CmbCGP.Text, 1) = "C" Then
         LLENA_GRUPOS frmCLI.cmbgrupo, 222
      Else
         LLENA_GRUPOS frmCLI.cmbgrupo, 223
      End If
      If Left(CmbCGP.Text, 1) = "C" Then
          LLENA_GRUPOS txtsubgrupo, 333
      Else
          LLENA_GRUPOS txtsubgrupo, 334
      End If

    frmCLI.txt_key.Locked = False
    frmCLI.txt_key.Enabled = True
    frmCLI.txt_key.SetFocus
    
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
If Left(CmbCGP.Text, 1) = "C" Then
  LLENA_GRUPOS frmCLI.cmbgrupo, 222
Else
  LLENA_GRUPOS frmCLI.cmbgrupo, 223
End If
cmbgrupo.SetFocus
SendKeys "%{up}"
fra1.Refresh

End Sub

Private Sub cmdagregar_Click()
On Error GoTo ESCAPA
If Trim(CmbCGP.Text) = "" Then
   MENSAJE_CLI "NO a seleccionado NADA ... !"
   Exit Sub
End If
If Left(cmdAgregar.Caption, 2) = "&A" And cmdAgregar.Enabled = True Then
    cmdAgregar.Caption = "&Grabar"
    cmdcancelar.Enabled = True
    cmdModificar.Enabled = False
    cmdeliminar.Enabled = False
    DESBLOQUEA_TEXT
    If LK_EMP <> "PAR" Then
     txt_key.Locked = True
    End If
    LIMPIA_CLI
    If Left(CmbCGP.Text, 1) = "C" Then
        frmCLI.OptNombre(0).Value = True
        frmCLI.txt_key = GENERA_CODI
    ElseIf Left(CmbCGP.Text, 1) = "P" Then
        frmCLI.OptNombre(0).Value = True
        frmCLI.txt_key = GENERA_PRO
    End If
    frmCLI.txtesposo.SetFocus
    txt_key.ToolTipText = ""
    CmbCGP.Enabled = False
    frmCLI.txtestado.ListIndex = 0
    frmCLI.SSTab.Tab = 0
    frmCLI.t_fechafac.Text = LK_FECHA_DIA
    pasa = 1
    LOC_CTA_CLI = ""
    LOC_CTA_CLI2 = ""
    TxtZona.ListIndex = 0
    TxtSubZona.ListIndex = 0
    txtZonaNew.ListIndex = 0
    TxtLugarCasa.ListIndex = 0
     'AGREGAMOS EN BLANCO
Else
 If Trim(frmCLI.txtesposo.Text) = "" Then
   MsgBox "Ingrese Nombre...", 48, Pub_Titulo
   frmCLI.txtesposo.SetFocus
  Exit Sub
 End If
  If Left(CmbCGP.Text, 1) = "C" Then
      If pasa = 1 Then
         If EXISTE_CLI("C", Left(frmCLI.txtesposo.Text, 15), Trim(txt_key.Text)) Then
            MENSAJE_CLI " Existen algunos clientes con estos NOMBRES .."
            frmCLI.ListExiste.SetFocus
            Exit Sub
         End If
      End If
      pasa = 0
      If par_llave!PAR_CONTABILIDAD = "" Then
        GoTo PASACONTAB
      End If
      If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" And Left(CmbCGP.Text, 1) = "C" Then
            GoTo PASACONTAB
      ElseIf Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" And Left(CmbCGP.Text, 1) = "P" Then
            GoTo PASACONTAB
      End If
      If Trim(LOC_CTA_CLI) = "" And Trim(LOC_DES_CLI) = "" Then
            If Trim(tcuenta.Text) = "" Then
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
     If (LK_EMP = "PAR" Or LK_EMP = "CAM" Or LK_EMP = "PIU") And COD_ORIGINAL <> Val(txt_key.Text) Then
      SQ_OPER = 1
      pu_codclie = Val(txt_key.Text)
      pu_cp = "C"
      pu_codcia = LK_CODCIA
      LEER_CLILOC_LLAVE
      If Not cliloc_llave.EOF Then
         MsgBox "Cliente Existe en Compañia ..", 48, Pub_Titulo
         Azul txt_key, txt_key
         Exit Sub
      End If
     End If
     On Error GoTo VERLO_GRABAR
     CN.Execute "Begin Transaction", rdExecDirect
     pub_cadena = "SELECT * FROM CONTROLL"
     Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
     frmCLI.txt_key = GENERA_CODI
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
  ElseIf Left(CmbCGP.Text, 1) = "P" Then

      If pasa = 1 Then
         If EXISTE_CLI("P", Left(frmCLI.txtesposo.Text, 15), Trim(txt_key.Text)) Then
            MENSAJE_CLI " Existen algunos Proveedor con estos NOMBRES .."
            frmCLI.ListExiste.SetFocus
            Exit Sub
         End If
      End If
       pasa = 0
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
         If Trim(tcuenta.Text) = "" Then
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
       frmCLI.txt_key = GENERA_PRO
       GRABAR_CLI "P"
       con_llave.Close
       CN.Execute "Commit Transaction", rdExecDirect
       On Error GoTo 0
       MENSAJE_CLI "Proveedor , AGREGADO... "
    End If
    cmdAgregar.Caption = "&Agregar"
    cmdeliminar.Enabled = True
    cmdModificar.Enabled = True

    BLOQUEA_TEXT
    txt_key.Locked = False
    CmbCGP.Enabled = True
    Screen.MousePointer = 0
    frmCLI.SSTab.Tab = 0
    txt_key.ToolTipText = ""
    LIMPIA_CLI
    LOC_CTA_CLI = ""
    LOC_CTA_CLI2 = ""
End If
Exit Sub
    
ESCAPA:
   If Err.Number = 40002 Then
      Screen.MousePointer = 0
      MsgBox "El Codigo generado ya existe " & Chr(13) & "Se procede a generar el siguiente codigo y a continuación " & Chr(13) & "Intente Grabar Nuevamente...", 48, Pub_Titulo
      frmCLI.txt_key = GENERA_CODI
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
     MsgBox Err.Description
     CN.Execute "Rollback Transaction", rdExecDirect
'    End If
    cmdcancelar_Click
fin:
End Sub

Private Sub cmdagregar_GotFocus()
If ListView1.Visible Then
 frmCLI.txt_key.Text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub cmdAgregar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   If frmCLI.txt_key.Visible Then
      frmCLI.txt_key.SetFocus
   End If
End If

End Sub

Private Sub cmdcancelar_Click()
If txt_key.Visible = False Then
  Exit Sub
End If
If Left(cmdAgregar.Caption, 2) = "&A" And Left(cmdModificar.Caption, 2) = "&M" Then
    LIMPIA_CLI
    cmdcancelar.Enabled = True
    txt_key.Locked = False
    MENSAJE_CLI "Proceso Cancelado... !!!    "
    txt_key.Enabled = True
    txt_key.SetFocus
    frmCLI.SSTab.Tab = 0
    Screen.MousePointer = 0
    pasa = 0
    cmdcontab.Enabled = False
    LOC_CTA_CLI = ""
    LOC_CTA_CLI2 = ""
    Exit Sub
End If
     Screen.MousePointer = 11
     If Left(cmdModificar.Caption, 2) = "&G" Then
        cmdModificar.Caption = "&Modificar"
        If Left(CmbCGP.Text, 1) = "C" Then
           LLENA_CLI 1, "C"
        Else
           LLENA_CLI 1, "P"
        End If
        txt_key.Locked = True
     Else
        GoSub ELI_TABLAS
        cmdAgregar.Caption = "&Agregar"
        cmdcontab.Enabled = False
        LIMPIA_CLI
        txt_key.Locked = False
        txt_key.SetFocus
     End If
     cmdAgregar.Enabled = True
     cmdeliminar.Enabled = True
     cmdModificar.Enabled = True

     txt_key.ToolTipText = ""
     wGARANTES = 0
     BLOQUEA_TEXT
     MENSAJE_CLI "Proceso Cancelado... !!!    "
     CmbCGP.Enabled = True
     frmCLI.SSTab.Tab = 0
     Screen.MousePointer = 0
     LOC_CTA_CLI = ""
     LOC_CTA_CLI2 = ""
     pasa = 0


Exit Sub
ELI_TABLAS:
If LK_FLAG_GRIFO <> "A" Then Return
pu_codclie = Val(txt_key.Text)
If pu_codclie = 0 Then Return
PSPLAC_LLAVE(0) = LK_CODCIA
PSPLAC_LLAVE(1) = 2101
PSPLAC_LLAVE(2) = pu_codclie
cliplac_llave.Requery
Do Until cliplac_llave.EOF
  cliplac_llave.Delete
  cliplac_llave.MoveNext
Loop

PSPLAC_LLAVE(0) = LK_CODCIA
PSPLAC_LLAVE(1) = 2301
PSPLAC_LLAVE(2) = pu_codclie
cliplac_llave.Requery
Do Until cliplac_llave.EOF
  cliplac_llave.Delete
  cliplac_llave.MoveNext
Loop




Return
End Sub

Private Sub cmdCancelar_GotFocus()
If ListView1.Visible Then
 frmCLI.txt_key.Text = ""
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
 frmCLI.txt_key.Text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub cmdCerrar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmCLI.txt_key.SetFocus
End If
End Sub

Private Sub cmdconfirma_Click()
  If Op(0).Value And Left(frmCLI.CmbCGP, 1) = "C" Then
     frmCLI.txt_key.Text = ListExiste.TextMatrix(ListExiste.Row, 1)
     pasa = 1
     frmCLI.F14.Visible = False
     cmdagregar_Click
     Exit Sub
  End If
  If Op(0).Value And Left(frmCLI.CmbCGP, 1) = "P" Then
    frmCLI.txtnombre.Text = ListExiste.TextMatrix(ListExiste.Row, 2)
    frmCLI.txt_key.Text = ListExiste.TextMatrix(ListExiste.Row, 1)
     pasa = 1
     frmCLI.F14.Visible = False
     If Left(cmdAgregar.Caption, 2) = "&G" And cmdAgregar.Enabled = True Then cmdagregar_Click
     If Left(cmdModificar.Caption, 2) = "&G" And cmdModificar.Enabled = True Then CmdModificar_Click
     Exit Sub
  End If
  If Op(1).Value Then
     pasa = 0
     frmCLI.F14.Visible = False
     If Left(cmdAgregar.Caption, 2) = "&G" And cmdAgregar.Enabled = True Then cmdagregar_Click
     If Left(cmdModificar.Caption, 2) = "&G" And cmdModificar.Enabled = True Then CmdModificar_Click
     Exit Sub
  End If
  MsgBox "Seleccione una de las dos Opciones ..", 48, Pub_Titulo
End Sub

Private Sub cmdcontab_Click()
If par_llave!PAR_CONTABILIDAD <> "A" Then
  Exit Sub
End If
If Left(CmbCGP.Text, 1) = "C" Then
  If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" And Left(cmdcontab.Caption, 2) = "&Q" Then
      tcuenta.Text = ""
      cmdcontab.Caption = "Relacionar a Con&tabilidad"
      Exit Sub
  End If
ElseIf Left(CmbCGP.Text, 1) = "P" Then
  If Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" And Left(cmdcontab.Caption, 2) = "&Q" Then
      tcuenta.Text = ""
      cmdcontab.Caption = "Relacionar a Con&tabilidad"
      Exit Sub
  End If
End If
If Left(cmdcontab.Caption, 2) = "&Q" Then
    pub_mensaje = "Confirmar la eliminación de la Cuenta : " & tcuenta.Text & " , Continuar ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If
    SQ_OPER = 1
    PUB_CUENTA = Trim(tcuenta.Text)
    PUB_CODCIA = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      tcuenta.Text = ""
    Else
      com_llave.Delete
      tcuenta.Text = ""
      CmdModificar_Click
    End If
    cmdcontab.Caption = "Relacionar a Con&tabilidad"
    Exit Sub
End If
LOC_CANCELA = 0
If txtesposo.Text = "" Then
 MsgBox "Ingrese Descripción del cliente..", 48, Pub_Titulo
 Azul txtesposo, txtesposo
 Exit Sub
End If
If Left(CmbCGP.Text, 1) = "C" Then
    LK_TABLA = "CLIENTE"
    archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND ( COM_CUENTA >= '" & 12 & "' AND COM_CUENTA < '" & 13 & "' OR COM_CUENTA >= '" & 16 & "' AND COM_CUENTA < '" & 17 & "' OR COM_CUENTA >= '" & 14 & "' AND COM_CUENTA < '" & 15 & "' ) ORDER BY COM_CUENTA"
Else
   LK_TABLA = "PROVEEDOR"
   archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND (COM_CUENTA >= '" & 42 & "' AND COM_CUENTA < '" & 43 & "' OR COM_CUENTA >= '" & 46 & "' AND COM_CUENTA < '" & 47 & "' )  ORDER BY COM_CUENTA"
End If
LOC_CTA_CLI = ""
LOC_DES_CLI = ""
pb.Visible = True
DoEvents
Load frmBuscacta
frmBuscacta.lbltabla.Caption = LK_TABLA
pb.Visible = False
frmBuscacta.Show 1
LOC_CTA_CLI = Trim(frmBuscacta.tcuenta)
LOC_DES_CLI = Trim(frmBuscacta.tnombre.Text)
LOC_NIVEL = Val(frmBuscacta.txtdatos(0).Text)
LOC_CTA_SUP = Trim(frmBuscacta.txtdatos(1).Text)
LOC_FLAG_AFEC = Trim(frmBuscacta.txtdatos(2).Text)
LOC_ESTADO = Trim(frmBuscacta.txtdatos(3).Text)
LOC_TIPO_CTA = Val(frmBuscacta.txtdatos(4).Text)
LOC_SIGNO_D = Val(frmBuscacta.txtdatos(5).Text)
LOC_SIGNO_H = Val(frmBuscacta.txtdatos(6).Text)
LOC_ACT_PAS = Val(frmBuscacta.txtdatos(7).Text)
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
If Left(CmbCGP.Text, 1) = "C" Then
   If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" Then
     Exit Sub
   End If
ElseIf Left(CmbCGP.Text, 1) = "P" Then
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
If Left(CmbCGP.Text, 1) = "P" Then
  If Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" And Left(cmdcontab2.Caption, 2) = "&Q" Then
      tcuenta2.Text = ""
      cmdcontab2.Caption = "Relacionar a Con&tabilidad"
      Exit Sub
  End If
End If
If Left(CmbCGP.Text, 1) = "C" Then
  If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" And Left(cmdcontab2.Caption, 2) = "&Q" Then
      tcuenta2.Text = ""
      cmdcontab2.Caption = "Relacionar a Con&tabilidad"
      Exit Sub
  End If
End If

If Left(cmdcontab2.Caption, 2) = "&Q" Then
    pub_mensaje = "Confirmar la eliminación de la Cuenta : " & tcuenta2.Text & " , Continuar ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If
     tcuenta2.Text = ""
     CmdModificar_Click
     cmdcontab2.Caption = "Relacionar a Con&tabilidad"
    Exit Sub
End If
LOC_CANCELA = 0
If txtesposo.Text = "" Then
 MsgBox "Ingrese Descripción del cliente..", 48, Pub_Titulo
 Azul txtesposo, txtesposo
 Exit Sub
End If
If Left(CmbCGP.Text, 1) = "C" Then
   LK_TABLA = "CLIENTES2"
   archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND (COM_CUENTA >= '" & 70 & "' AND COM_CUENTA < '" & 71 & "' OR COM_CUENTA >= '" & 75 & "' AND COM_CUENTA < '" & 78 & "') ORDER BY COM_CUENTA"
End If
If Left(CmbCGP.Text, 1) = "P" Then
   LK_TABLA = "PROVEEDOR2"
   archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND (COM_CUENTA >= '" & 60 & "' AND COM_CUENTA < '" & 61 & "' OR COM_CUENTA >= '" & 62 & "' AND COM_CUENTA < '" & 68 & "' OR COM_CUENTA >= '" & 33 & "' AND COM_CUENTA < '" & 39 & "' ) ORDER BY COM_CUENTA"
End If
LOC_CTA_CLI2 = ""
LOC_DES_CLI2 = ""
PB2.Visible = True
DoEvents
Load frmBuscacta
frmBuscacta.lbltabla.Caption = LK_TABLA
pb.Visible = False
frmBuscacta.Show 1
LOC_CTA_CLI2 = Trim(frmBuscacta.tcuenta)
LOC_DES_CLI2 = Trim(frmBuscacta.tnombre.Text)
LOC_NIVEL2 = Val(frmBuscacta.txtdatos(0).Text)
LOC_CTA_SUP2 = Trim(frmBuscacta.txtdatos(1).Text)
LOC_FLAG_AFEC2 = Trim(frmBuscacta.txtdatos(2).Text)
LOC_ESTADO2 = Trim(frmBuscacta.txtdatos(3).Text)
LOC_TIPO_CTA2 = Val(frmBuscacta.txtdatos(4).Text)
LOC_SIGNO_D2 = Val(frmBuscacta.txtdatos(5).Text)
LOC_SIGNO_H2 = Val(frmBuscacta.txtdatos(6).Text)
LOC_ACT_PAS2 = Val(frmBuscacta.txtdatos(7).Text)
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
If Left(CmbCGP.Text, 1) = "P" Then
   If Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" Then
     Exit Sub
   End If
End If

If Left(cmdModificar.Caption, 2) = "&G" Then
   CmdModificar_Click
End If

End Sub

Private Sub cmdeliminar_Click()
Dim wcias As String
On Error GoTo SALE
If Len(txt_key) = 0 Or Len(txtnombre) = 0 Then
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
  PS_REP01.rdoParameters(0) = " "
  PS_REP01.rdoParameters(1) = 0
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
  If Trim(Nulo_Valors(GEN!gen_cli_cias)) <> "" Then
    wcias = Trim(GEN!gen_cli_cias)
    MsgBox "O J O ...  Al Eliminar este Cliente tambien debe hacerlo con las demas Compañias relacionadas : " & wcias, 48, Pub_Titulo
  End If
  If Trim(tcuenta.Text) <> "" Then
    pub_mensaje = " ¿Desea Eliminar el Registro, y su Relacion a Contabilidad .. ?"
  Else
    pub_mensaje = " ¿Desea Eliminar el Registro... ?"
  End If
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbYes Then   ' El usuario eligió
    Screen.MousePointer = 11
    cliloc_llave.Delete
    frmCLI.txt_key.Text = ""
    frmCLI.txt_key.Locked = False
    If Trim(tcuenta.Text) <> "" Then
     SQ_OPER = 1
     PUB_CUENTA = Trim(tcuenta.Text)
     PUB_CODCIA = LK_CODCIA
     LEER_COM_LLAVE
     If com_llave.EOF Then
         tcuenta.Text = ""
     Else
         com_llave.Delete
         tcuenta.Text = ""
     End If
    End If
    If Trim(tcuenta2.Text) <> "" Then
     SQ_OPER = 1
     PUB_CUENTA = Trim(tcuenta2.Text)
     PUB_CODCIA = LK_CODCIA
     LEER_COM_LLAVE
     If com_llave.EOF Then
         tcuenta2.Text = ""
     Else
         com_llave.Delete
         tcuenta2.Text = ""
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
frmCLI.txt_key.Text = ""
frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub cmdEliminar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmCLI.txt_key.SetFocus
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
If Len(txt_key) = 0 Or Len(txtnombre) = 0 Then
   MENSAJE_CLI "NO a seleccionado NADA ... !"
   Exit Sub
End If
If Left(cmdModificar.Caption, 2) = "&M" Then
    cmdModificar.Caption = "&Grabar"
    cmdeliminar.Enabled = False
    cmdAgregar.Enabled = False
    cmdcancelar.Enabled = True
    CmbCGP.Enabled = False
    DESBLOQUEA_TEXT
    txt_key.Locked = True
    frmCLI.txtesposo.SetFocus
    pasa = 1
    LOC_CTA_CLI = ""
    LOC_CTA_CLI2 = ""

 Else
   If Left(CmbCGP.Text, 1) = "C" Then
      If pasa = 1 Then
         If EXISTE_CLI("C", Left(frmCLI.txtesposo.Text, 15), Trim(txt_key.Text)) Then
            MENSAJE_CLI " Existen algunos clientes con estos NOMBRES .."
            frmCLI.ListExiste.SetFocus
            Exit Sub
         End If
      End If
      pasa = 0
   ElseIf Left(CmbCGP.Text, 1) = "P" Then
     If pasa = 1 Then
      If EXISTE_CLI("P", Left(frmCLI.txtesposo.Text, 15), Trim(txt_key.Text)) Then
         MENSAJE_CLI " Existen algunos Proveedor con estos NOMBRES .."
         frmCLI.ListExiste.SetFocus
         Exit Sub
      End If
    End If
    pasa = 0
   End If
   If Not CONSIS_CLI Then
         '  "NO SE PUEDE.."
      Exit Sub
   End If
   If par_llave!PAR_CONTABILIDAD = "" Then
      GoTo PASACONTAB
   End If
   If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" And Left(CmbCGP.Text, 1) = "C" Then
      GoTo PASACONTAB
   ElseIf Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" And Left(CmbCGP.Text, 1) = "P" Then
      GoTo PASACONTAB
   End If
   If Left(cmdcontab.Caption, 2) <> "&Q" Or Left(cmdcontab2.Caption, 2) <> "&Q" Then
      If Trim(tcuenta.Text) <> "" Or Trim(tcuenta2.Text) <> "" Then
         GRABA_CONTAB LK_CODCIA
      End If
      
   End If
PASACONTAB:
    Screen.MousePointer = 11
    GRABAR_CLI "C"
    MENSAJE_CLI "Registro , MODIFICADO... "
    cmdModificar.Caption = "&Modificar"
    frmCLI.SSTab.Tab = 0
    Screen.MousePointer = 0
    cmdcancelar.Enabled = True
    cmdeliminar.Enabled = True
    cmdAgregar.Enabled = True
    BLOQUEA_TEXT
    txt_key.Locked = True
    CmbCGP.Enabled = True
    cmdcancelar.SetFocus
    Screen.MousePointer = 0
    LOC_CTA_CLI = ""
    LOC_CTA_CLI2 = ""
  
End If
End Sub

Private Sub cmdModificar_GotFocus()
If ListView1.Visible Then
 frmCLI.txt_key.Text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub Command1_Click()
Dim wserie
Dim wnumfac
Dim wnumv

SQ_OPER = 2
pu_codclie = 0
pu_cp = Left(CmbCGP.Text, 1)
pu_codcia = LK_CODCIA
LEER_CLI_LLAVE

Do Until cli_mayor.EOF
  SQ_OPER = 5
  pu_codclie = cli_mayor!CLI_CODCLIE
  pu_cp = Left(CmbCGP.Text, 1)
  pu_codcia = LK_CODCIA
  LEER_CLI_LLAVE
  If cls_llave.EOF Then
  cls_llave.AddNew
  cls_llave!CLS_CODCIA = LK_CODCIA
  cls_llave!CLS_CODCLIE = cli_mayor!CLI_CODCLIE
  cls_llave!CLS_CP = cli_mayor!CLI_CP
  cls_llave!CLS_DEB00 = 0
  cls_llave!CLS_HAB00 = 0
  cls_llave!CLS_DEB01 = 0
  cls_llave!CLS_HAB01 = 0
  cls_llave!CLS_DEB02 = 0
  cls_llave!CLS_HAB02 = 0
  cls_llave!CLS_DEB03 = 0
  cls_llave!CLS_HAB03 = 0
  cls_llave!CLS_DEB04 = 0
  cls_llave!CLS_HAB04 = 0
  cls_llave!CLS_DEB05 = 0
  cls_llave!CLS_HAB05 = 0
  cls_llave!CLS_DEB06 = 0
  cls_llave!CLS_HAB06 = 0
  cls_llave!CLS_DEB07 = 0
  cls_llave!CLS_HAB07 = 0
  cls_llave!CLS_DEB08 = 0
  cls_llave!CLS_HAB08 = 0
  cls_llave!CLS_DEB09 = 0
  cls_llave!CLS_HAB09 = 0
  cls_llave!CLS_DEB10 = 0
  cls_llave!CLS_HAB10 = 0
  cls_llave!CLS_DEB11 = 0
  cls_llave!CLS_HAB11 = 0
  cls_llave!CLS_DEB12 = 0
  cls_llave!CLS_HAB12 = 0
  cls_llave.Update
End If

cli_mayor.MoveNext
Loop


MsgBox "TERMINO"

Exit Sub

Stop
pub_cadena = "SELECT * FROM MOVICONT where MOV_CODCIA = '01' AND MOV_PLANTILLA = 14 ORDER BY MOV_NRO_MES , MOV_NRO_VOUCHER , MOV_DH "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
llave_rep01.Requery
If llave_rep01.EOF Then
 MsgBox "NO HAY DATOS"
 Exit Sub
End If
wserie = llave_rep01!MOV_serie
wnumfac = llave_rep01!MOV_numfac
wnumv = llave_rep01!MOV_NRO_VOUCHER
Do Until llave_rep01.EOF
 If llave_rep01!MOV_NRO_VOUCHER = wnumv Then
   If Val(llave_rep01!MOV_numfac) = 0 Then
    llave_rep01.Edit
     llave_rep01!MOV_serie = wserie
     llave_rep01!MOV_numfac = wnumfac
     llave_rep01.Update
     wnumv = llave_rep01!MOV_NRO_VOUCHER
     GoTo sa
    End If
  Else
   wnumv = llave_rep01!MOV_NRO_VOUCHER
  End If
  wserie = llave_rep01!MOV_serie
  wnumfac = llave_rep01!MOV_numfac
sa:
 llave_rep01.MoveNext
Loop


Exit Sub


End Sub

Private Sub copia_Click()
Dim valor


'Load frmPesos
'frmPesos.Show 1
'Exit Sub
   If Val(frmCLI.txt_key.Text) <= 0 Then
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
    cliloc_llave!CLI_CP = Left(CmbCGP.Text, 1)
    cliloc_llave!CLI_CODCLIE = Val(frmCLI.txt_key.Text)
    cliloc_llave!cli_SALDO = 0
    cliloc_llave!CLI_DET_TOT = "D"
    cliloc_llave!CLI_MONEDA = " "
    cliloc_llave!CLI_CODCIA = valor
    cliloc_llave!CLI_NOMBRE_ESPOSO = txtesposo.Text
    cliloc_llave!CLI_NOMBRE_ESPOSA = Txtesposa.Text
    cliloc_llave!CLI_NOMBRE_EMPRESA = TxtEmpresa.Text
    ASIGNA_123
    cliloc_llave!cli_nombre = frmCLI.txtnombre.Text
    cliloc_llave!CLI_CASA_DIREC = Txtdireccion.Text
    cliloc_llave!CLI_CASA_NUM = Val(Txtnumdir.Text)
    cliloc_llave!CLI_CASA_ZONA = Val(Right(TxtZona.Text, 4))
    cliloc_llave!CLI_LUGAR_CASA = Val(Right(TxtLugarCasa.Text, 4))
    cliloc_llave!CLI_LUGAR_TRAB = Val(Right(TxtLugarTrab.Text, 4))
    cliloc_llave!CLI_CASA_SUBZONA = Val(Right(TxtSubZona.Text, 4))
    cliloc_llave!CLI_ZONA_NEW = Val(Right(txtZonaNew.Text, 4))
    cliloc_llave!CLI_TRAB_DIREC = txtDirTrabajo.Text
    cliloc_llave!CLI_TRAB_NUM = Nulo_Valor0(txtnumdirtrabajo.Text)
    cliloc_llave!cli_TRAB_ZONA = Val(Right(frmCLI.TxtZonaTrabajo.Text, 4))
    cliloc_llave!cli_TRAB_SUBZONA = Val(Right(TxtSubZonaTrabajo.Text, 4))
    cliloc_llave!cli_ruc_esposo = txtRUCesposo.Text
    cliloc_llave!cli_ruc_esposA = txtRUCesposa.Text
    cliloc_llave!CLI_RUC_EMPRESA = txtRUCempresa.Text
    cliloc_llave!CLI_CASA1 = frmCLI.txtpropiedad1.Text
    cliloc_llave!CLI_CASA2 = frmCLI.txtpropiedad2.Text
    cliloc_llave!CLI_REGPUB1 = frmCLI.txtregpublico1.Text
    cliloc_llave!CLI_REGPUB2 = frmCLI.txtregpublico2.Text
    cliloc_llave!CLI_AUTOAVALUO = frmCLI.txtautovaluo.Text
    cliloc_llave!CLI_AUTO1 = frmCLI.txtauto1.Text
    cliloc_llave!CLI_AUTO2 = frmCLI.txtauto2.Text
    cliloc_llave!CLI_PRENDA = frmCLI.txtprendas.Text
    cliloc_llave!CLI_TELEF1 = frmCLI.txttelefono1.Text
    cliloc_llave!CLI_TELEF2 = frmCLI.txttelefono2.Text
    cliloc_llave!CLI_OTRO_CONTR = frmCLI.otrocontrato.Value
    cliloc_llave!CLI_LETRA = frmCLI.letraotorgado.Value
    cliloc_llave!CLI_GRUPO = Val(Right(frmCLI.cmbgrupo.Text, 4))
    cliloc_llave!CLI_SUBGRUPO = Val(Right(frmCLI.txtsubgrupo.Text, 4))
    cliloc_llave!CLI_nucleo = frmCLI.txtNucleo.Text
    cliloc_llave!CLI_estado = Left(frmCLI.txtestado.Text, 1)
    cliloc_llave!CLI_programado = Nulo_Valors(txtprog.Text)
    '  <<< Actualiza La Cta. solo de la Cia Actual >>>
    cliloc_llave!CLI_CUENTA_CONTAB = Trim(frmCLI.tcuenta.Text)
    cliloc_llave!CLI_CUENTA_CONTAB2 = Trim(frmCLI.tcuenta2.Text)
    If txtDTX.Text = "" Then
      txtDTX.Text = " "
    End If
    cliloc_llave!CLI_DET_TOT = txtDTX.Text
    cliloc_llave!cli_limcre = Val(txtlimite.Text)
cliloc_llave.Update
MsgBox "Proceso Copiado .... ", 48, Pub_Titulo
Unload frmCLI
End Sub

Private Sub Form_Activate()
'frmCLI.Refresh
End Sub

Private Sub Form_Load()
Dim i As Integer
COD_ORIGINAL = 0
LOC_CTA_CLI = ""
LOC_DES_CLI = ""
LOC_CTA_CLI2 = ""
LOC_DES_CLI2 = ""

If Not cop_llave.EOF Then
For i = 1 To 6
  If cop_llave.rdoColumns(i) <> 0 Then
     wCOM_NIVEL(i) = cop_llave.rdoColumns(i)
     NIVEL_MAX = i
  End If
Next i
End If

pub_cadena = "SELECT * FROM PARGEN WHERE PAR_CODCIA = ?  ORDER BY PAR_CODCIA "
Set PSPAR_CLI = CN.CreateQuery("", pub_cadena)
PSPAR_CLI.rdoParameters(0) = " "
Set par_llave_cli = PSPAR_CLI.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  
pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CP=? AND CLI_CODCLIE  = ? AND CLI_CODCIA = ? ORDER BY CLI_CP ,CLI_CODCLIE"
Set PSCLILOC_LLAVE = CN.CreateQuery("", pub_cadena)
PSCLILOC_LLAVE.rdoParameters(0) = " "
PSCLILOC_LLAVE.rdoParameters(1) = 0
PSCLILOC_LLAVE.rdoParameters(2) = " "
Set cliloc_llave = PSCLILOC_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CP = ? AND CLI_CODCLIE  >= ? AND CLI_CODCIA = ? ORDER BY CLI_CP ,CLI_CODCLIE"
Set PSCLILOC_MAYOR = CN.CreateQuery("", pub_cadena)
PSCLILOC_MAYOR.rdoParameters(0) = " "
PSCLILOC_MAYOR.rdoParameters(1) = 0
PSCLILOC_MAYOR.rdoParameters(2) = " "
Set cliloc_mayor = PSCLILOC_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_CP = ? AND CLI_CODCLIE  >= ? ORDER BY CLI_CP ,CLI_CODCLIE"
Set PSCLI_MAYOR2 = CN.CreateQuery("", pub_cadena)
PSCLI_MAYOR2.rdoParameters(0) = " "
PSCLI_MAYOR2.rdoParameters(1) = " "
PSCLI_MAYOR2.rdoParameters(2) = 0
Set cli_mayor2 = PSCLI_MAYOR2.OpenResultset(rdOpenKeyset, rdConcurValues)


pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ? AND TAB_CODCLIE = ? ORDER BY TAB_NOMLARGO"
Set PSPLAC_LLAVE = CN.CreateQuery("", pub_cadena)
PSPLAC_LLAVE.rdoParameters(0) = 0
PSPLAC_LLAVE.rdoParameters(1) = 0
PSPLAC_LLAVE.rdoParameters(2) = 0
Set cliplac_llave = PSPLAC_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)


If LK_EMP = "HER" Then
 pub_cadena = "SELECT CLI_NOMBRE,CLI_RUC_ESPOSO FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_CP = ? AND CLI_RUC_ESPOSO = ? and CLI_CODCLIE <> ?  ORDER BY CLI_CODCLIE"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 PS_REP01.rdoParameters(0) = " "
 PS_REP01.rdoParameters(1) = " "
 PS_REP01.rdoParameters(2) = " "
 PS_REP01.rdoParameters(3) = 0
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
End If

ACCESO_CLI
loc_ultcod = 0
' para el reporte de bancos...reinicio
frmCLI.F14.Left = 90
frmCLI.F14.Top = 3360
frmCLI.F14.Height = 2415
frmCLI.F14.Width = 9380
llave1 = ""
UNICO = ""
pasa = 0
For pasa = 0 To frmCLI.ListExiste.Col - 1
  frmCLI.ListExiste.Col = pasa
  frmCLI.ListExiste.FixedAlignment(pasa) = 2
Next pasa
pasa = 0
frmCLI.ListExiste.Cols = 5
frmCLI.ListExiste.ColWidth(0) = 350
frmCLI.ListExiste.ColWidth(1) = 800
frmCLI.ListExiste.ColWidth(2) = 3300
frmCLI.ListExiste.ColWidth(3) = 2000
frmCLI.ListExiste.ColWidth(4) = 3000
wGARANTES = 0
'ALLINVISIBLE
BLOQUEA_TEXT
LLENA_ZONA TxtZona, 20
LLENA_ZONA TxtSubZona, 30
LLENA_ZONA txtZonaNew, 35
LLENA_ZONA TxtZonaTrabajo, 20
LLENA_ZONA TxtSubZonaTrabajo, 35
LLENA_GRUPOS cmbgrupo, 222
LLENA_GRUPOS txtsubgrupo, 333
LLENA_ZONA TxtLugarCasa, 25
LLENA_ZONA TxtLugarTrab, 25
''frmCLI.SSTab1.TabCaption(0) = "&Datos Clientes - Principales"
''frmCLI.SSTab1.TabCaption(1) = "&Datos Clientes - Opcionales"
LOC_TIPREG = 300 ' CLIENTES
ETIQUETA_CLI
llave1 = "X"
CmbCGP.ListIndex = 0
llave1 = ""
Screen.MousePointer = 0
txt_key.MaxLength = 15
cmdcontab.Enabled = False
If LK_FLAG_GRIFO = "A" Then
    fraplaca.Visible = True
    cmdmante.Visible = True
    g_fechafac.Visible = True
    g_diasfac.Visible = True
    t_fechafac.Visible = True
    t_diasfac.Visible = True
Else
    cmdmante.Visible = False
    fraplaca.Visible = False
    g_fechafac.Visible = False
    g_diasfac.Visible = False
    t_fechafac.Visible = False
    t_diasfac.Visible = False
End If
frmCLI.txt_key.TabIndex = 0
copia.Visible = False
If LK_CODUSU = "ADMIN" Then
    copia.Visible = True
End If

End Sub

Public Sub ALLINVISIBLE()
    frmCLI.lcuenta.Visible = False
    txt_key.Visible = False
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
    frmCLI.txtpordes.Visible = False
    t_diascred.Visible = False
End Sub
Public Sub ALLVISIBLE()
    frmCLI.lcuenta.Visible = True
    txt_key.Visible = True
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
    frmCLI.txtpordes.Visible = True
    t_diascred.Visible = True
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
     frmCLI.txtnombre.Text = Nulo_Valors(frmCLI.txtesposo.Text)
     cliloc_llave!CLI_123 = 1
  ElseIf frmCLI.OptNombre(1).Value Then
     frmCLI.txtnombre.Text = Nulo_Valors(frmCLI.Txtesposa.Text)
     cliloc_llave!CLI_123 = 2
  ElseIf frmCLI.OptNombre(2).Value Then
     frmCLI.txtnombre.Text = Nulo_Valors(frmCLI.TxtEmpresa.Text)
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
  tab_llave!TAB_CODCIA = LK_CODCIA
  tab_llave!TAB_TIPREG = LOC_TIPREG
  tab_llave!TAB_NUMTAB = Val(lblnom(Index).Tag)
  tab_llave!tab_nomlargo = Left(wnombre, 40)
  tab_llave!tab_nomcorto = Left(wnombre, 10)
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
 txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
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
 txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 ListView1.Visible = False
 txt_key.Text = ""
 txt_key.SetFocus
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
   Parpadea.Enabled = False
   LblMensaje.Visible = False
 End If
End Sub

Public Sub ASIGNA_INT(WCONTROL As ComboBox, txt As Integer)
For fila = 0 To WCONTROL.ListCount - 1
    If Val(Trim(Right(WCONTROL.LIST(fila), 3))) = txt Then
        WCONTROL.ListIndex = fila
        Exit Sub
    End If
Next fila
End Sub
Public Sub ASIGNA_subgrupo(WCONTROL As ComboBox, txt As String)
For fila = 0 To WCONTROL.ListCount - 1
    If Val(Trim(Right(WCONTROL.LIST(fila), 3))) = Val(txt) Then
        WCONTROL.ListIndex = fila
        Exit Sub
    End If
Next fila
End Sub

Public Sub LLENA_CLI(ban As Integer, CG As String)
    If ban = 0 Then
        '**  BAN = 0 BUSCA DATOS NUEVAMENTE
        If loc_key > ListView1.ListItems.Count Or loc_key = 0 Then
         Else
          txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
        End If
        pu_cp = Left(CmbCGP.Text, 2)
        pu_codclie = Val(txt_key.Text)
        SQ_OPER = 1
        pu_codcia = LK_CODCIA
        LEER_CLILOC_LLAVE
    End If
    loc_ultcod = Val(cliloc_llave!CLI_CODCLIE)
    frmCLI.txt_key.Text = cliloc_llave!CLI_CODCLIE
    LLENA_123
    txtnombre.Text = Nulo_Valors(cliloc_llave!cli_nombre)
    txtnombre.MaxLength = cliloc_llave(3).Size
    txtesposo.Text = Trim(Nulo_Valors(cliloc_llave!CLI_NOMBRE_ESPOSO))
    txtesposo.MaxLength = cliloc_llave(4).Size
    Txtesposa.Text = Trim(Nulo_Valors(cliloc_llave!CLI_NOMBRE_ESPOSA))
    TxtEmpresa.Text = Trim(Nulo_Valors(cliloc_llave!CLI_NOMBRE_EMPRESA))
    Txtdireccion.Text = Trim(Nulo_Valors(cliloc_llave!CLI_CASA_DIREC))
    Txtdireccion.MaxLength = cliloc_llave(10).Size
    Txtnumdir.Text = Trim(Nulo_Valor0(cliloc_llave!CLI_CASA_NUM))
    ASIGNA_INT TxtZona, Nulo_Valor0(cliloc_llave!CLI_CASA_ZONA)
    ASIGNA_INT TxtSubZona, Nulo_Valor0(cliloc_llave!CLI_CASA_SUBZONA)
    ASIGNA_INT txtZonaNew, Nulo_Valor0(cliloc_llave!CLI_ZONA_NEW)
    txtDirTrabajo.Text = Trim(Nulo_Valors(cliloc_llave!CLI_TRAB_DIREC))
    txtDirTrabajo.MaxLength = cliloc_llave(14).Size
    txtnumdirtrabajo.Text = Trim(Nulo_Valor0(cliloc_llave!CLI_TRAB_NUM))
    ASIGNA_INT TxtZonaTrabajo, Nulo_Valor0(cliloc_llave!cli_TRAB_ZONA)
    ASIGNA_INT TxtSubZonaTrabajo, Nulo_Valor0(cliloc_llave!cli_TRAB_SUBZONA)
    ASIGNA_INT TxtLugarCasa, Nulo_Valor0(cliloc_llave!CLI_LUGAR_CASA)
    ASIGNA_INT TxtLugarTrab, Nulo_Valor0(cliloc_llave!CLI_LUGAR_TRAB)
       
    txtRUCesposo.Text = Trim(Nulo_Valors(cliloc_llave!cli_ruc_esposo))
    If LK_DIG_RUC <> 0 Then txtRUCesposo.MaxLength = LK_DIG_RUC
    
    txtRUCesposa.Text = Trim(Nulo_Valors(cliloc_llave!cli_ruc_esposA))
    txtRUCempresa.Text = Trim(Nulo_Valors(cliloc_llave!CLI_RUC_EMPRESA))
    frmCLI.txtpropiedad1.Text = Trim(Nulo_Valors(cliloc_llave!CLI_CASA1))
    frmCLI.txtpropiedad2.Text = Trim(Nulo_Valors(cliloc_llave!CLI_CASA2))
    frmCLI.txtregpublico1.Text = Trim(Nulo_Valors(cliloc_llave!CLI_REGPUB1))
    frmCLI.txtregpublico2.Text = Trim(Nulo_Valors(cliloc_llave!CLI_REGPUB2))
    frmCLI.txtautovaluo.Text = Trim(Nulo_Valors(cliloc_llave!CLI_AUTOAVALUO))
    frmCLI.txtauto1.Text = Trim(Nulo_Valors(cliloc_llave!CLI_AUTO1))
    frmCLI.txtauto2.Text = Trim(Nulo_Valors(cliloc_llave!CLI_AUTO2))
    frmCLI.txtprendas.Text = Trim(Nulo_Valors(cliloc_llave!CLI_PRENDA))
    frmCLI.txttelefono1.Text = Trim(Nulo_Valors(cliloc_llave!CLI_TELEF1))
    frmCLI.txttelefono2.Text = Trim(Nulo_Valors(cliloc_llave!CLI_TELEF2))
    frmCLI.otrocontrato.Value = Nulo_Valor0(cliloc_llave!CLI_OTRO_CONTR)
    frmCLI.letraotorgado.Value = Nulo_Valor0(cliloc_llave!CLI_LETRA)
    LLENA_BLOQ
    ASIGNA_INT cmbgrupo, Nulo_Valors(cliloc_llave!CLI_GRUPO)
    ASIGNA_subgrupo txtsubgrupo, Nulo_Valors(cliloc_llave!CLI_SUBGRUPO)
    frmCLI.txtNucleo.Text = Nulo_Valor0(cliloc_llave!CLI_nucleo)
    If Nulo_Valors(cliloc_llave!CLI_estado) = "A" Then
      frmCLI.txtestado.ListIndex = 0
    Else
    frmCLI.txtestado.ListIndex = 1
    End If
    frmCLI.txtprog.Text = Nulo_Valors(cliloc_llave!CLI_programado)
    frmCLI.tcuenta.Text = Nulo_Valors(cliloc_llave!CLI_CUENTA_CONTAB)
    If Trim(Nulo_Valors(cliloc_llave!CLI_CUENTA_CONTAB)) <> "" Then
        cmdcontab.Caption = "&Quitar Relacion Contable"
    Else
        cmdcontab.Caption = "Relacionar a Con&tabilidad"
    End If
    frmCLI.tcuenta2.Text = Nulo_Valors(cliloc_llave!CLI_CUENTA_CONTAB2)
    If Trim(Nulo_Valors(cliloc_llave!CLI_CUENTA_CONTAB2)) <> "" Then
        cmdcontab2.Caption = "&Quitar Relacion Contable"
    Else
        cmdcontab2.Caption = "Relacionar a Con&tabilidad"
    End If
    frmCLI.txtlimite.Text = Nulo_Valor0(cliloc_llave!cli_limcre)
    txtDTX.Text = Nulo_Valors(cliloc_llave!CLI_DET_TOT)
    frmCLI.txtpordes.Text = Nulo_Valor0(cliloc_llave!CLI_PORDESCTO)
    t_fechafac.Text = Format(cliloc_llave!cli_fecha_fac, "dd/mm/yyyy")
    t_diasfac.Text = Nulo_Valor0(cliloc_llave!cli_DIAS_FAC)
    frmCLI.t_diascred.Text = Nulo_Valor0(cliloc_llave!cli_DIAS_CRED)
    pu_codclie = Val(txt_key.Text)
    If LK_FLAG_GRIFO = "A" Then
      LLENA_DESCTO
    End If
    
End Sub

Public Sub LIMPIA_CLI()
    txt_key.Text = ""
    txtnombre.Text = ""
    txtesposo.Text = ""
    Txtesposa.Text = ""
    TxtEmpresa.Text = ""
    Txtdireccion.Text = ""
    Txtnumdir.Text = ""
    TxtZona.ListIndex = -1
    TxtSubZona.ListIndex = -1
    txtZonaNew.ListIndex = -1
    TxtLugarCasa.ListIndex = -1
    TxtLugarTrab.ListIndex = -1
    txtDirTrabajo.Text = ""
    txtnumdirtrabajo.Text = ""
    frmCLI.TxtZonaTrabajo.ListIndex = -1
    TxtSubZonaTrabajo.ListIndex = -1
    txtRUCesposo.Text = ""
    txtRUCesposa.Text = ""
    txtRUCempresa.Text = ""
    frmCLI.txtpropiedad2.Text = ""
    frmCLI.txtpropiedad1.Text = ""
    frmCLI.txtregpublico1.Text = ""
    frmCLI.txtregpublico2.Text = ""
    frmCLI.txtautovaluo.Text = ""
    frmCLI.txtauto1.Text = ""
    frmCLI.txtauto2.Text = ""
    frmCLI.txtprendas.Text = ""
    frmCLI.txttelefono1.Text = ""
    frmCLI.txttelefono2.Text = ""
    frmCLI.otrocontrato.Value = 0
    frmCLI.letraotorgado.Value = 0
    frmCLI.ListBloqueos.Clear
    frmCLI.cmbgrupo.ListIndex = -1
    frmCLI.txtsubgrupo.ListIndex = -1
    frmCLI.txtNucleo.Text = ""
    frmCLI.txtestado.ListIndex = -1
    frmCLI.tcuenta.Text = ""
    frmCLI.OptNombre(0).Value = False
    frmCLI.OptNombre(1).Value = False
    frmCLI.OptNombre(2).Value = False
    frmCLI.txtlimite.Text = ""
    frmCLI.txtDTX = ""
    txtprog.Text = ""
    LOC_CTA_CLI = ""
    LOC_DES_CLI = ""
    tcuenta2.Text = ""
    LOC_CTA_CLI2 = ""
    LOC_DES_CLI2 = ""
    cmdcontab.Caption = "Relacionar a Con&tabilidad"
    cmdcontab2.Caption = "Relacionar a Con&tabilidad"
    frmCLI.txtpordes.Text = ""
    COD_ORIGINAL = 0
    t_fechafac.Text = ""
    t_diasfac.Text = ""
    t_diascred.Text = ""
    frmCLI.grid_des.Clear
    
    frmCLI.grid_des.Rows = 1
    
End Sub

'Private Sub SSTab1_Click(PreviousTab As Integer)
'If SSTab1.Tab = 0 Then
' If txtesposo.Enabled And txtesposo.Visible Then
'   txtesposo.SetFocus
' End If
'Else
' If txtDirTrabajo.Enabled And txtDirTrabajo.Visible Then
'   txtDirTrabajo.SetFocus
' End If
'End If
'End Sub

Private Sub SSTab1_GotFocus()
If ListView1.Visible Then
 frmCLI.txt_key.Text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub t_diascred_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
    If frmCLI.cmdModificar.Enabled Then
          frmCLI.cmdModificar.SetFocus
    Else
           frmCLI.cmdAgregar.SetFocus
    End If
 End If
End Sub

Private Sub t_diasfac_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  frmCLI.t_fechafac.SetFocus
End If

End Sub

Private Sub t_fechacred_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii

End Sub

Private Sub t_fechafac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  frmCLI.t_diascred.SetFocus
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
           frmCLI.cmdAgregar.SetFocus
    End If
 End If
End Sub

Private Sub txt_key_GotFocus()
 Azul txt_key, txt_key
End Sub

Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo SALE
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If KeyCode = 113 Then
 If CmbCGP.ListIndex = 1 Then
  CmbCGP.ListIndex = 0
 Else
  CmbCGP.ListIndex = 1
 End If
End If
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txt_key.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
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
GoTo fin
POSICION:
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  txt_key.SelStart = Len(txt_key.Text)
fin:
Exit Sub
SALE:

End Sub
Private Sub txt_key_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem
On Error GoTo SALCODI
If KeyAscii = 13 Then
If LK_EMP = "PAR" And Left(cmdAgregar.Caption, 2) = "&G" Then
  Azul txtesposo, txtesposo
  Exit Sub
End If
End If

If KeyAscii = 27 And Trim(txtnombre.Text) = "" Then
 txt_key.Text = ""
End If
If KeyAscii <> 13 Or Left(cmdAgregar.Caption, 2) = "&G" Or Left(cmdModificar.Caption, 2) = "&G" Then
   GoTo fin
End If
On Error GoTo CODI_ERR
pu_codclie = Val(txt_key.Text)
On Error GoTo 0
If Len(txt_key.Text) = 0 Then
   Exit Sub
End If
fra2.Refresh
If pu_codclie <> 0 And IsNumeric(txt_key.Text) = True Then
   If Len(Trim(txt_key.Text)) = LK_DIG_RUC Then ' LONG DEL RUC
        pu_cp = Left(CmbCGP.Text, 1)
        PUB_RUC = Trim(txt_key.Text)
        SQ_OPER = 4
        pu_codcia = LK_CODCIA
        LEER_CLI_LLAVE
        If cli_ruc.EOF Then
           MsgBox "R.U.C. No Existe ", 48, Pub_Titulo
           Exit Sub
        End If
        txt_key.Text = cli_ruc!CLI_CODCLIE
   End If
    SQ_OPER = 1
   On Error GoTo mucho
   pu_codcia = LK_CODCIA
   pu_cp = Left(CmbCGP.Text, 1)
   pu_codclie = Val(txt_key.Text)
   LEER_CLILOC_LLAVE
   On Error GoTo 0
   If cliloc_llave.EOF Then
     Azul txt_key, txt_key
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     txt_key.SetFocus
     GoTo fin
   End If
   ListView1.Visible = False
   cmdcancelar.Enabled = True
   If Left(CmbCGP.Text, 1) = "C" Then
         LLENA_CLI 1, "C"
   End If
   If Left(CmbCGP.Text, 1) = "P" Then
         LLENA_CLI 1, "P"
   End If
   frmCLI.txt_key.Locked = True
   frmCLI.cmdModificar.SetFocus
   Screen.MousePointer = 0
Else
   If loc_key > ListView1.ListItems.Count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(ListView1.ListItems.Item(loc_key).Text)
   If Trim(UCase(txt_key.Text)) = Left(valor, Len(Trim(txt_key.Text))) Then
   Else
      Exit Sub
   End If
   ListView1.Visible = False
   cmdcancelar.Enabled = True
   If Left(CmbCGP.Text, 1) = "C" Then
         LLENA_CLI 0, "C"
   End If
   If Left(CmbCGP.Text, 1) = "P" Then
         LLENA_CLI 0, "P"
   End If
   frmCLI.txt_key.Locked = True
   cmdcancelar.Enabled = True
   frmCLI.cmdModificar.SetFocus
End If
dale:
ListView1.Visible = False
fin:
mucho:
CODI_ERR:
Exit Sub
SALCODI:
MsgBox Err.Description & " Intente Nuevamente ", 48, Pub_Titulo
Unload frmCLI
End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim NADA
Dim var
If Len(txt_key.Text) = 0 Or IsNumeric(txt_key.Text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(txt_key.Text) = 1 Then
    If txt_key.Text = "" Then txt_key.Text = " "
    var = Asc(txt_key.Text)
    var = var + 1
    NADA = var
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 1
    archi = "SELECT CLI_CODCLIE, CLI_CODCIA, CLI_CP, CLI_NOMBRE,CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM FROM CLIENTES WHERE  CLI_CP = '" & Left(CmbCGP.Text, 1) & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txt_key.Text & "' AND  '" & var & "' ORDER BY CLI_NOMBRE"
    PROC_LISVIEW ListView1
    loc_key = 1
    If NADA = 33 Or NADA = 91 Then
      If ListView1.Visible = False Then
        loc_key = 0
        MsgBox "No existe Datos ...", 48, Pub_Titulo
        txt_key.Text = ""
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
  Set itmFound = ListView1.FindItem(LTrim(txt_key.Text), lvwText, , lvwPartial)
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
If Left(CmbCGP.Text, 1) = "C" Then
  If Trim(txtDirTrabajo.Text) = "" Then
    txtDirTrabajo.Text = Trim(Txtdireccion.Text)
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
GEN.Requery
TOTCIAS = Nulo_Valors(GEN!gen_cli_cias)
xcuenta = 1
For fila = 1 To 30
    If Trim(Mid(TOTCIAS, xcuenta, 2)) = LK_CODCIA Then
         GoTo SIGUE_PASA
    End If
    xcuenta = xcuenta + 2
Next fila
GoTo CIA_ACTUAL

SIGUE_PASA:
If Trim(TOTCIAS) <> "" And Left(CmbCGP.Text, 1) = "C" Then
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
             If VAR_CIAS = LK_CODCIA Then GoTo pasa
             SQ_OPER = 1
             pu_cp = wCGH
             pu_codclie = Val(frmCLI.txt_key.Text)
             pu_codcia = VAR_CIAS
             LEER_CLILOC_LLAVE
             If cliloc_llave.EOF Then
'                MsgBox "No Grabo en la Compañia : " + VAR_CIAS + " No Existe cliente ", 48, Pub_Titulo
             Else
               cliloc_llave.Edit
               Modo = "E"
               GoSub grabar
               cliloc_llave.Update
             End If
pasa:
           Else
             cliloc_llave.AddNew
             Modo = "A"
             GoSub grabar
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
    pu_codclie = Val(frmCLI.txt_key.Text)
    pu_codcia = VAR_CIAS
    LEER_CLILOC_LLAVE
    If cliloc_llave.EOF Then
      MsgBox "No Grabo en la Compañia : " + VAR_CIAS + " No Existe cliente ", 48, Pub_Titulo
    Else
      cliloc_llave.Edit
      cliloc_llave!cli_limcre = Val(frmCLI.txtlimite.Text)
      Modo = "E"
      GoSub grabar
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
  GoSub grabar
  cliloc_llave.Update

  Exit Sub
End If
 
Exit Sub
   
grabar:
    If Modo = "A" Then
       cliloc_llave!CLI_CP = wCGH
       cliloc_llave!CLI_CODCLIE = Val(frmCLI.txt_key.Text)
       cliloc_llave!cli_SALDO = 0
       cliloc_llave!CLI_DET_TOT = "D"
       cliloc_llave!CLI_MONEDA = "S"
       If Left(CmbCGP.Text, 1) = "C" Then
        loc_ultcod = Val(frmCLI.txt_key.Text)
       End If
          SQ_OPER = 5
          pu_codclie = Val(frmCLI.txt_key.Text)
          pu_cp = wCGH
          pu_codcia = LK_CODCIA
          LEER_CLI_LLAVE
          If cls_llave.EOF Then
          cls_llave.AddNew
          cls_llave!CLS_CODCIA = VAR_CIAS
          cls_llave!CLS_CODCLIE = Val(frmCLI.txt_key.Text)
          cls_llave!CLS_CP = wCGH
          cls_llave!CLS_DEB00 = 0
          cls_llave!CLS_HAB00 = 0
          cls_llave!CLS_DEB01 = 0
          cls_llave!CLS_HAB01 = 0
          cls_llave!CLS_DEB02 = 0
          cls_llave!CLS_HAB02 = 0
          cls_llave!CLS_DEB03 = 0
          cls_llave!CLS_HAB03 = 0
          cls_llave!CLS_DEB04 = 0
          cls_llave!CLS_HAB04 = 0
          cls_llave!CLS_DEB05 = 0
          cls_llave!CLS_HAB05 = 0
          cls_llave!CLS_DEB06 = 0
          cls_llave!CLS_HAB06 = 0
          cls_llave!CLS_DEB07 = 0
          cls_llave!CLS_HAB07 = 0
          cls_llave!CLS_DEB08 = 0
          cls_llave!CLS_HAB08 = 0
          cls_llave!CLS_DEB09 = 0
          cls_llave!CLS_HAB09 = 0
          cls_llave!CLS_DEB10 = 0
          cls_llave!CLS_HAB10 = 0
          cls_llave!CLS_DEB11 = 0
          cls_llave!CLS_HAB11 = 0
          cls_llave!CLS_DEB12 = 0
          cls_llave!CLS_HAB12 = 0
          cls_llave.Update
        End If
    End If
    cliloc_llave!CLI_CODCIA = VAR_CIAS
    cliloc_llave!CLI_NOMBRE_ESPOSO = txtesposo.Text
    cliloc_llave!CLI_NOMBRE_ESPOSA = Txtesposa.Text
    cliloc_llave!CLI_NOMBRE_EMPRESA = TxtEmpresa.Text
    ASIGNA_123
    cliloc_llave!cli_nombre = frmCLI.txtnombre.Text
    cliloc_llave!CLI_CASA_DIREC = Txtdireccion.Text
    cliloc_llave!CLI_CASA_NUM = Val(Txtnumdir.Text)
    cliloc_llave!CLI_CASA_ZONA = Val(Right(TxtZona.Text, 4))
    cliloc_llave!CLI_LUGAR_CASA = Val(Right(TxtLugarCasa.Text, 4))
    cliloc_llave!CLI_LUGAR_TRAB = Val(Right(TxtLugarTrab.Text, 4))
    cliloc_llave!CLI_CASA_SUBZONA = Val(Right(TxtSubZona.Text, 4))
    cliloc_llave!CLI_ZONA_NEW = Val(Right(txtZonaNew.Text, 4))
    cliloc_llave!CLI_TRAB_DIREC = txtDirTrabajo.Text
    cliloc_llave!CLI_TRAB_NUM = Nulo_Valor0(txtnumdirtrabajo.Text)
    cliloc_llave!cli_TRAB_ZONA = Val(Right(frmCLI.TxtZonaTrabajo.Text, 4))
    cliloc_llave!cli_TRAB_SUBZONA = Val(Right(TxtSubZonaTrabajo.Text, 4))
    cliloc_llave!cli_ruc_esposo = txtRUCesposo.Text
    cliloc_llave!cli_ruc_esposA = txtRUCesposa.Text
    cliloc_llave!CLI_RUC_EMPRESA = txtRUCempresa.Text
    cliloc_llave!CLI_CASA1 = frmCLI.txtpropiedad1.Text
    cliloc_llave!CLI_CASA2 = frmCLI.txtpropiedad2.Text
    cliloc_llave!CLI_REGPUB1 = frmCLI.txtregpublico1.Text
    cliloc_llave!CLI_REGPUB2 = frmCLI.txtregpublico2.Text
    cliloc_llave!CLI_AUTOAVALUO = frmCLI.txtautovaluo.Text
    cliloc_llave!CLI_AUTO1 = frmCLI.txtauto1.Text
    cliloc_llave!CLI_AUTO2 = frmCLI.txtauto2.Text
    cliloc_llave!CLI_PRENDA = frmCLI.txtprendas.Text
    cliloc_llave!CLI_TELEF1 = frmCLI.txttelefono1.Text
    cliloc_llave!CLI_TELEF2 = frmCLI.txttelefono2.Text
    cliloc_llave!CLI_OTRO_CONTR = frmCLI.otrocontrato.Value
    cliloc_llave!CLI_LETRA = frmCLI.letraotorgado.Value
    cliloc_llave!CLI_GRUPO = Val(Right(frmCLI.cmbgrupo.Text, 4))
    cliloc_llave!CLI_SUBGRUPO = Val(Right(frmCLI.txtsubgrupo.Text, 4))
    cliloc_llave!CLI_nucleo = frmCLI.txtNucleo.Text
    cliloc_llave!CLI_estado = Left(frmCLI.txtestado.Text, 1)
    cliloc_llave!CLI_programado = Nulo_Valors(txtprog.Text)
    cliloc_llave!CLI_PORDESCTO = Nulo_Valor0(txtpordes.Text)
    If LK_FLAG_GRIFO = "A" Then
     cliloc_llave!cli_fecha_fac = Format(t_fechafac.Text, "dd/mm/yyyy")
    Else
     cliloc_llave!cli_fecha_fac = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    End If
     cliloc_llave!cli_DIAS_FAC = Val(t_diasfac.Text)
     cliloc_llave!cli_DIAS_CRED = Val(frmCLI.t_diascred.Text)

    '  <<< Actualiza La Cta. solo de la Cia Actual >>>
    If VAR_CIAS = LK_CODCIA Then
      cliloc_llave!CLI_CUENTA_CONTAB = Trim(frmCLI.tcuenta.Text)
      cliloc_llave!CLI_CUENTA_CONTAB2 = Trim(frmCLI.tcuenta2.Text)
    End If
    If txtDTX.Text = "" Then
     txtDTX.Text = " "
    End If
    cliloc_llave!CLI_DET_TOT = txtDTX.Text
    If Trim(TOTCIAS) = "" Then
      cliloc_llave!cli_limcre = Val(txtlimite.Text)
    End If
   
Return
End Sub

Public Sub MENSAJE_CLI(TEXTO As String)
  LblMensaje.Caption = TEXTO
  Parpadea.Enabled = True
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
    If LK_EMP = "PAR" Or LK_EMP = "CAM" Or LK_EMP = "PIU" Then
      INTpub_cadena = Val(NUMCAD)
      If COD_ORIGINAL <> 0 And INTpub_cadena <> Val(txt_key.Text) Then
        INTpub_cadena = Val(txt_key.Text)
        GoTo GEN
      End If
      COD_ORIGINAL = INTpub_cadena
      GoTo GEN
    End If
Else
    cliloc_mayor.MoveLast
    NUMCAD = cliloc_mayor!CLI_CODCLIE
    If LK_EMP = "PAR" Or LK_EMP = "CAM" Or LK_EMP = "PIU" Then
      INTpub_cadena = Val(NUMCAD) + 1
      If COD_ORIGINAL <> 0 And INTpub_cadena <> Val(txt_key.Text) Then
        INTpub_cadena = Val(txt_key.Text)
        GoTo GEN
      End If
      COD_ORIGINAL = INTpub_cadena
      GoTo GEN
    End If
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
GEN:
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
    If LK_EMP = "PAR" Or LK_EMP = "CAM" Or LK_EMP = "PIU" Then
      INTpub_cadena = Val(NUMCAD)
      If COD_ORIGINAL <> 0 And INTpub_cadena <> Val(txt_key.Text) Then
        INTpub_cadena = Val(txt_key.Text)
        GoTo GEN
      End If
      COD_ORIGINAL = INTpub_cadena
      GoTo GEN
    End If
Else
    cliloc_mayor.MoveLast
    NUMCAD = cliloc_mayor!CLI_CODCLIE
    If LK_EMP = "PAR" Or LK_EMP = "CAM" Or LK_EMP = "PIU" Then
      INTpub_cadena = Val(NUMCAD) + 1
      If COD_ORIGINAL <> 0 And INTpub_cadena <> Val(txt_key.Text) Then
        INTpub_cadena = Val(txt_key.Text)
        GoTo GEN
      End If
      COD_ORIGINAL = INTpub_cadena
      GoTo GEN
    End If
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
GEN:
GENERA_CODI = INTpub_cadena

End Function
Public Function CONSIS_CLI() As Boolean
Dim wruc As Integer
If frmCLI.OptNombre(0).Value Then
    If Trim(frmCLI.txtesposo.Text) = "" Then
        CONSIS_CLI = False
        MENSAJE_CLI "Ingrese Datos Principal ..."
        txtesposo.SetFocus
        GoTo ESCAPA
    End If
ElseIf frmCLI.OptNombre(1).Value Then
    If Trim(frmCLI.Txtesposa.Text) = "" Then
        CONSIS_CLI = False
        MENSAJE_CLI "Ingrese Datos Principal ..."
        Txtesposa.SetFocus
        GoTo ESCAPA
    End If
ElseIf frmCLI.OptNombre(2).Value Then
    If Trim(frmCLI.TxtEmpresa.Text) = "" Then
        CONSIS_CLI = False
        MENSAJE_CLI "Ingrese Datos Principal ..."
        TxtEmpresa.SetFocus
        GoTo ESCAPA
    End If
End If

If Len(frmCLI.txtesposo.Text) = 0 And Len(frmCLI.Txtesposa.Text) = 0 And Len(frmCLI.TxtEmpresa.Text) = 0 Then
        CONSIS_CLI = False
        MENSAJE_CLI "Ingrese Algun Nombre  ..."
        txtesposo.SetFocus
        GoTo ESCAPA
ElseIf frmCLI.OptNombre(0).Value And Len(frmCLI.txtesposo.Text) = 0 Then
        CONSIS_CLI = False
        MENSAJE_CLI "Nombre  NO Puede estar en Blanco ..."
        txtesposo.SetFocus
        GoTo ESCAPA
ElseIf frmCLI.OptNombre(1).Value And Len(frmCLI.Txtesposa.Text) = 0 Then
        CONSIS_CLI = False
        MENSAJE_CLI "Nombre  NO Puede estar en Blanco ..."
        Txtesposa.SetFocus
        GoTo ESCAPA
ElseIf frmCLI.OptNombre(2).Value And Len(frmCLI.TxtEmpresa.Text) = 0 Then
        CONSIS_CLI = False
        MENSAJE_CLI "Nombre  NO Puede estar en Blanco ..."
        TxtEmpresa.SetFocus
        GoTo ESCAPA
ElseIf Len(frmCLI.txtesposo.Text) = 0 And Len(frmCLI.txtRUCesposo.Text) > 0 Then
        CONSIS_CLI = False
        MENSAJE_CLI "RUC  debe estar en Blanco ..."
        txtRUCesposo.SetFocus
        GoTo ESCAPA
'ElseIf Len(frmCLI.Txtesposa.Text) = 0 And Len(frmCLI.txtRUCesposa.Text) > 0 Then
'        CONSIS_CLI = False
'        MENSAJE_CLI "L.E.  debe estar en Blanco ..."
'        txtRUCesposa.SetFocus
'        GoTo ESCAPA
ElseIf Len(frmCLI.TxtEmpresa.Text) = 0 And Len(frmCLI.txtRUCempresa.Text) > 0 Then
     If LK_EMP <> "PLA" Then
        CONSIS_CLI = False
        MENSAJE_CLI "RUC  debe estar en Blanco ..."
        txtRUCempresa.SetFocus
        GoTo ESCAPA
     End If
End If
wruc = 8
If LK_DIG_RUC <> 0 Then wruc = LK_DIG_RUC


If frmCLI.txtRUCesposo.Text <> "" Then
    If Len(Trim(frmCLI.txtRUCesposo.Text)) = wruc Then
    
    Else
       CONSIS_CLI = False
       MENSAJE_CLI "R.U.C. de No es Validad ..."
       frmCLI.txtRUCesposo.SetFocus
       GoTo ESCAPA
    End If
End If
If frmCLI.txtRUCesposa.Text <> "" Then
    If Len(Trim(frmCLI.txtRUCesposa.Text)) = 8 Or Len(Trim(frmCLI.txtRUCesposa.Text)) = 12 Then
    Else
       CONSIS_CLI = False
       MENSAJE_CLI "L.E. de No es Validad ..."
       frmCLI.txtRUCesposa.SetFocus
       GoTo ESCAPA
    End If
End If
If LK_EMP <> "PLA" Then
 If Left(CmbCGP.Text, 1) = "C" Then
  If frmCLI.txtRUCempresa.Text <> "" Then
    If Len(Trim(frmCLI.txtRUCempresa.Text)) <> 8 Then
       CONSIS_CLI = False
       MENSAJE_CLI "L.E. de No es Validad ..."
       txtRUCempresa.SetFocus
       GoTo ESCAPA
    End If
  End If
 End If
End If
If LK_EMP = "HER" And frmCLI.txtRUCesposo.Text <> "" Then
'CLI_CODCIA = ? AND CLI_CP = 'C' AND CLI_RUC_ESPOSO = ? and CLI_CODCLIE <> ?

 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = Left(frmCLI.CmbCGP, 1)
 PS_REP01(2) = frmCLI.txtRUCesposo.Text
 PS_REP01(3) = frmCLI.txt_key.Text
 llave_rep01.Requery
 If Not llave_rep01.EOF Then
   MsgBox "RUC Existe en otro Cliente : " + Trim(llave_rep01!cli_nombre), 48, Pub_Titulo
    CONSIS_CLI = False
    Azul frmCLI.txtRUCesposo, frmCLI.txtRUCesposo
    GoTo ESCAPA
 End If
End If
If Trim(tcuenta.Text) <> "" Then
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
   frmCLI.txtNucleo.Text = " "
End If

If Left(CmbCGP.Text, 1) = "C" Then
 If Trim(TxtLugarCasa.Text) = "" Then
    MsgBox "Dato no es opcional ,Lugar.", 48, Pub_Titulo
    CONSIS_CLI = False
    TxtLugarCasa.SetFocus
    GoTo ESCAPA
 End If
 If Trim(TxtZona.Text) = "" Then
    MsgBox "Dato no es opcional ,Definir.", 48, Pub_Titulo
    CONSIS_CLI = False
    TxtZona.SetFocus
    GoTo ESCAPA
 End If
 If Trim(TxtSubZona.Text) = "" Then
    MsgBox "Dato no es opcional ,Definir.", 48, Pub_Titulo
    CONSIS_CLI = False
    TxtSubZona.SetFocus
    GoTo ESCAPA
 End If
 If Trim(txtZonaNew.Text) = "" Then
    MsgBox "Dato no es opcional ,Definir.", 48, Pub_Titulo
    CONSIS_CLI = False
    txtZonaNew.SetFocus
    GoTo ESCAPA
 End If
End If
If LK_FLAG_GRIFO = "A" And Left(CmbCGP.Text, 1) = "C" Then
 If Not IsDate(frmCLI.t_fechafac.Text) Then
    MsgBox "Fecha para la Facturacion no procede.", 48, Pub_Titulo
    CONSIS_CLI = False
    Azul frmCLI.t_fechafac, frmCLI.t_fechafac
    GoTo ESCAPA
 End If
 If Trim(frmCLI.t_fechafac.Text) = "" Then
    frmCLI.t_fechafac.Text = LK_FECHA_DIA
 End If
Else
   frmCLI.t_fechafac.Text = LK_FECHA_DIA
End If
 


CONSIS_CLI = True
ESCAPA:
End Function

Private Sub txtesposo_LostFocus()
fra1.Refresh
End Sub

Private Sub txtestado_GotFocus()
'Azul txtestado, txtestado
End Sub

Private Sub txtestado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   'SIGUE_CAMPO frmCLI.txtestado.TabIndex
   fra2.Refresh
End If
End Sub

Private Sub txtlimite_GotFocus()
Azul txtlimite, txtlimite
End Sub

Private Sub txtlimite_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Left(cmdAgregar.Caption, 2) = "&G" Then
   cmdAgregar.SetFocus
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
On Error GoTo SIGUE
If Left(CmbCGP.Text, 1) = "C" Then
  If Trim(TxtLugarTrab.Text) = "" Then
    TxtLugarTrab.ListIndex = TxtLugarCasa.ListIndex
  End If
End If
fra1.Refresh
Exit Sub
SIGUE:
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
On Error GoTo SIGUE
If Left(CmbCGP.Text, 1) = "C" Then
  If Val(txtnumdirtrabajo.Text) = 0 Then
    txtnumdirtrabajo.Text = Txtnumdir.Text
  End If
End If
Exit Sub
SIGUE:
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

Private Sub txtpordes_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL txtpordes, KeyAscii
If KeyAscii = 13 Then
  frmCLI.TxtZona.SetFocus
  SendKeys "%{up}"
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
If Left(CmbCGP.Text, 1) = "C" Then
  LLENA_GRUPOS txtsubgrupo, 333
Else
  LLENA_GRUPOS txtsubgrupo, 334
End If

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
  Azul txtpordes, txtpordes
End If
End Sub

Private Sub txttelefono1_LostFocus()
On Error GoTo SIGUE
If Left(CmbCGP.Text, 1) = "C" Then
  If Trim(txttelefono2.Text) = "" Then
    txttelefono2.Text = txttelefono1.Text
  End If
End If
fra1.Refresh
Exit Sub
SIGUE:
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
On Error GoTo SIGUE
If Left(CmbCGP.Text, 1) = "C" Then
  If Trim(TxtZonaTrabajo.Text) = "" Then
      TxtZonaTrabajo.ListIndex = TxtZona.ListIndex
  End If
End If
fra1.Refresh
Exit Sub
SIGUE:


fra1.Refresh
End Sub

Private Sub txtZonaNew_LostFocus()
On Error GoTo SIGUE
If Left(CmbCGP.Text, 1) = "C" Then
  If Trim(TxtSubZonaTrabajo.Text) = "" Then
      TxtSubZonaTrabajo.ListIndex = txtZonaNew.ListIndex
  End If
End If
fra1.Refresh
Exit Sub
SIGUE:
End Sub

Private Sub TxtZonaTrabajo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.TxtZonaTrabajo.TabIndex
End If
End Sub

Public Function EXISTE_CLI(WCP As String, VALOR1 As String, WCODI As String) As Boolean
Dim var
Dim tempo
tempo = Left(Trim(VALOR1), Len(VALOR1) - 1)
var = Asc(Right(Trim(VALOR1), 1))
var = var + 1
If var = 91 Then
  var = "ZZZZZZZZ"
Else
  var = Chr(var)
End If
tempo = tempo + var
archi = "SELECT * FROM CLIENTES WHERE  CLI_CODCLIE <> " & WCODI & " AND CLI_CP = '" & WCP & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & VALOR1 & "' AND  '" & tempo & "' ORDER BY CLI_NOMBRE"
ListExiste.Clear
Set PSX = CN.CreateQuery("", archi)
Set X = PSX.OpenResultset(rdOpenKeyset)
X.Requery
EXISTE_CLI = False
If X.EOF Then
 frmCLI.ListExiste.Clear
 GoTo fin
End If

If WCP = "P" Then
    F14.Caption = "Lista de Proveedores parecidos ... "
End If
frmCLI.ListExiste.TextMatrix(0, 0) = "Cia"
frmCLI.ListExiste.TextMatrix(0, 1) = "Codigo "
If WCP = "C" Then
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
GoTo fin
Exit Function

CHECKERROR:
MsgBox Err.Description
fin:

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
  lblnom(fila).Caption = Trim(tab_llave!tab_nomlargo)
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
Dim sal As Boolean
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
   GoTo otrito
 End If
 If TypeOf frmCLI.Controls(fila) Is MSFlexGrid Then
   GoTo otrito
 End If
 If TypeOf frmCLI.Controls(fila) Is Line Then
   GoTo otrito
 End If
 If frmCLI.Controls(fila).WhatsThisHelpID = WNUM Then
   frmCLI.Controls(fila).Visible = True
 End If
 
otrito:
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
On Error GoTo SIGUE
Dim wmax As Integer
Dim cuenta As Integer
wmax = 42
fila = WTAG
Do Until fila >= wmax
 fila = fila + 1
 cuenta = 0
 Do Until cuenta >= frmCLI.Controls.Count - 1
  If TypeOf frmCLI.Controls(cuenta) Is Timer Then
    GoTo otrito
  End If
  If TypeOf frmCLI.Controls(cuenta) Is MSFlexGrid Then
    GoTo otrito
  End If
'  MsgBox frmCLI.Controls(fila).Name
  If TypeOf frmCLI.Controls(cuenta) Is OptionButton Then
    GoTo otrito
  End If
  If frmCLI.Controls(cuenta).TabIndex = fila Then
    If frmCLI.Controls(cuenta).Visible Then
         frmCLI.Controls(cuenta).SetFocus
         Exit Sub
    End If
  End If
otrito:
  cuenta = cuenta + 1
 Loop
Loop
If frmCLI.cmdModificar.Enabled Then
   frmCLI.cmdModificar.SetFocus
Else
   frmCLI.cmdAgregar.SetFocus
End If
Exit Sub
SIGUE:
Resume Next
End Sub

Public Sub GRABA_CONTAB(wcia As String)
Dim flagpase As String * 1
 If Left(CmbCGP.Text, 1) = "C" Then
   If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" Then
     Exit Sub
   End If
 ElseIf Left(CmbCGP.Text, 1) = "P" Then
   If Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" Then
     Exit Sub
   End If
 End If
If Trim(LOC_CTA_CLI) <> "" Then
 If Trim(tcuenta.Text) <> "" Then
   flagpase = ""
LETRAS:
   On Error GoTo cuenta1
  
    com_llave.AddNew
    com_llave!COM_CODCIA = wcia
    com_llave!com_cuenta = LOC_CTA_CLI
    com_llave!com_DESCRIPCION = LOC_DES_CLI
    com_llave!com_nivel = LOC_NIVEL
    com_llave!com_cuenta_sup = LOC_CTA_SUP
    com_llave!com_FLAG_AFECTACION = LOC_FLAG_AFEC
    com_llave!com_ESTADO = LOC_ESTADO
    com_llave!com_tipo_cta = LOC_TIPO_CTA
    com_llave!com_signo_d = LOC_SIGNO_D
    com_llave!com_signo_h = LOC_SIGNO_H
    com_llave!com_ACT_PAS = LOC_ACT_PAS
    com_llave!com_signo_h = LOC_SIGNO_H
    com_llave!com_ACT_PAS = LOC_ACT_PAS
    com_llave!COM_DEB_MES = 0
    com_llave!COM_HAB_MES = 0
    com_llave!COM_DEB_ANO = 0
    com_llave!COM_HAB_ANO = 0
    com_llave!com_cuenta_AUTOM_D = " "
    com_llave!com_cuenta_AUTO_H = " "
    com_llave!COM_CUENTA_AUTOM_D2 = " "
    com_llave!COM_CUENTA_AUTOM_D3 = " "
    com_llave!COM_CUENTA_AUTOM_D4 = " "
    com_llave!COM_CUENTA_AUTOM_D5 = " "
    com_llave!COM_POR_AUTOM_D = 0
    com_llave!COM_POR_AUTOM_D2 = 0
    com_llave!COM_POR_AUTOM_D3 = 0
    com_llave!COM_POR_AUTOM_D4 = 0
    com_llave!COM_POR_AUTOM_D5 = 0
    com_llave!COM_CENTRO_COSTOS = " "
  com_llave.Update
  If Left(CmbCGP.Text, 1) = "P" And Left(Trim(LOC_CTA_CLI), 2) = "42" And flagpase <> "A" Then
    flagpase = "A"
    'LOC_CTA_CLI = "42101"
    LOC_CTA_CLI = Mid(Trim(LOC_CTA_CLI), 1, 2) + "3" + Mid(Trim(LOC_CTA_CLI), 4, Len(Trim(LOC_CTA_CLI)))
    GoTo LETRAS
  End If
  On Error GoTo 0
  cmdcontab.Caption = "&Quitar Relación Contable"
 End If
End If

Exit Sub
If Trim(LOC_CTA_CLI2) <> "" Then
 If Trim(tcuenta2.Text) <> "" Then
    On Error GoTo CUENTA2
    com_llave.AddNew
    com_llave!COM_CODCIA = wcia
    com_llave!com_cuenta = LOC_CTA_CLI2
    com_llave!com_DESCRIPCION = LOC_DES_CLI2
    com_llave!com_nivel = LOC_NIVEL2
    com_llave!com_cuenta_sup = LOC_CTA_SUP2
    com_llave!com_FLAG_AFECTACION = LOC_FLAG_AFEC2
    com_llave!com_ESTADO = LOC_ESTADO2
    com_llave!com_tipo_cta = LOC_TIPO_CTA2
    com_llave!com_signo_d = LOC_SIGNO_D2
    com_llave!com_signo_h = LOC_SIGNO_H2
    com_llave!com_ACT_PAS = LOC_ACT_PAS2
    com_llave!com_signo_h = LOC_SIGNO_H2
    com_llave!com_ACT_PAS = LOC_ACT_PAS2
    com_llave!COM_DEB_MES = 0
    com_llave!COM_HAB_MES = 0
    com_llave!COM_HAB_ANO = 0
    com_llave!COM_DEB_ANO = 0
    com_llave!com_cuenta_AUTOM_D = ""
    com_llave!com_cuenta_AUTO_H = ""
    com_llave!COM_CUENTA_AUTOM_D2 = " "
    com_llave!COM_CUENTA_AUTOM_D3 = " "
    com_llave!COM_CUENTA_AUTOM_D4 = " "
    com_llave!COM_CUENTA_AUTOM_D5 = " "
    com_llave!COM_POR_AUTOM_D = 0
    com_llave!COM_POR_AUTOM_D2 = 0
    com_llave!COM_POR_AUTOM_D3 = 0
    com_llave!COM_POR_AUTOM_D4 = 0
    com_llave!COM_POR_AUTOM_D5 = 0
    com_llave!COM_CENTRO_COSTOS = " "
    com_llave.Update
    On Error GoTo 0
    cmdcontab2.Caption = "&Quitar Relación Contable"
 End If
End If
Exit Sub

cuenta1:
If Err.Number = 40002 Then
  MsgBox "Cuenta Existe, NO Procede. Cta.: " & LOC_CTA_CLI, 48, Pub_Titulo
   tcuenta.Text = ""
  com_llave.CancelUpdate
End If
Exit Sub

CUENTA2:
If Err.Number = 40002 Then
  MsgBox "Cuenta Existe, NO Procede. Cta.: " & LOC_CTA_CLI2, 48, Pub_Titulo
  tcuenta2.Text = ""
  com_llave.CancelUpdate
End If
Exit Sub

End Sub

Public Sub LLENA_DESCTO()

PSPLAC_LLAVE(0) = LK_CODCIA
PSPLAC_LLAVE(1) = 2301
PSPLAC_LLAVE(2) = pu_codclie
cliplac_llave.Requery
frmCLI.grid_des.Cols = 4
frmCLI.grid_des.Clear
frmCLI.grid_des.ColWidth(0) = 0
frmCLI.grid_des.ColWidth(1) = 1600
frmCLI.grid_des.ColWidth(2) = 600
frmCLI.grid_des.ColWidth(3) = 600
frmCLI.grid_des.Rows = 1
frmCLI.grid_des.TextMatrix(0, 0) = "Cod."
frmCLI.grid_des.TextMatrix(0, 1) = "Descrip."
frmCLI.grid_des.TextMatrix(0, 2) = "P.Contado"
frmCLI.grid_des.TextMatrix(0, 3) = "P.Credito"
SQ_OPER = 1
pu_codcia = LK_CODCIA
Do Until cliplac_llave.EOF
  frmCLI.grid_des.Rows = frmCLI.grid_des.Rows + 1
  frmCLI.grid_des.TextMatrix(frmCLI.grid_des.Rows - 1, 0) = cliplac_llave!TAB_CODART
  PUB_KEY = cliplac_llave!TAB_CODART
  LEER_ART_LLAVE
  If Not art_LLAVE.EOF Then frmCLI.grid_des.TextMatrix(frmCLI.grid_des.Rows - 1, 1) = art_LLAVE!ART_NOMBRE
  frmCLI.grid_des.TextMatrix(frmCLI.grid_des.Rows - 1, 2) = Format(cliplac_llave!tab_nomlargo, "0.00")
  frmCLI.grid_des.TextMatrix(frmCLI.grid_des.Rows - 1, 3) = Format(cliplac_llave!tab_nomcorto, "0.00")
  cliplac_llave.MoveNext
Loop
grid_des.SetFocus
End Sub
