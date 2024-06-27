VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form FORM_GRIFO 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Transacciones en Tiempo Real"
   ClientHeight    =   6495
   ClientLeft      =   75
   ClientTop       =   1380
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6495
   ScaleWidth      =   9480
   Tag             =   "55"
   WindowState     =   2  'Maximized
   Begin ComctlLib.ListView LV_ART2 
      Height          =   375
      Left            =   3360
      TabIndex        =   133
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox i_unidades 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   80
      Tag             =   "0"
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid gridC 
      Height          =   495
      Left            =   6600
      TabIndex        =   175
      Tag             =   "195"
      Top             =   720
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   327680
      Rows            =   50
      Cols            =   15
      BackColor       =   16777215
      BackColorFixed  =   12632256
      ForeColorFixed  =   64
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
   End
   Begin VB.CommandButton compras2 
      BackColor       =   &H00808000&
      Caption         =   "Compras &Parte II"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   174
      Tag             =   "196"
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox i_destino 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   173
      Tag             =   "138"
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox i_origen 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   172
      Tag             =   "137"
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox textovar3 
      BackColor       =   &H00FFFF00&
      Height          =   405
      Left            =   360
      TabIndex        =   171
      Tag             =   "0"
      Text            =   "textovar3"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox textovar2 
      BackColor       =   &H00FFFF00&
      Height          =   405
      Left            =   120
      TabIndex        =   132
      Tag             =   "0"
      Text            =   "textovar2"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox i_fecha_lote 
      Height          =   285
      Left            =   720
      TabIndex        =   140
      Tag             =   "71"
      Text            =   "i_fecha_lote"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton lotes 
      BackColor       =   &H00808000&
      Caption         =   "Ordenes de &Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   139
      Tag             =   "175"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox i_ser_lote 
      Height          =   285
      Left            =   4560
      TabIndex        =   138
      Tag             =   "89"
      Text            =   "I_num_lote"
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame opciones 
      Height          =   615
      Left            =   600
      TabIndex        =   136
      Tag             =   "183"
      Top             =   2160
      Visible         =   0   'False
      Width           =   3735
      Begin VB.OptionButton Option8 
         Caption         =   "&P.T."
         Height          =   375
         Left            =   1200
         TabIndex        =   170
         Tag             =   "9999"
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton Option7 
         Caption         =   "&Caract."
         Height          =   375
         Left            =   120
         TabIndex        =   137
         Tag             =   "9999"
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Caption         =   "&Det."
         Height          =   375
         Left            =   2040
         TabIndex        =   135
         Tag             =   "9999"
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton Option6 
         Caption         =   "&Res."
         Height          =   375
         Left            =   2880
         TabIndex        =   134
         Tag             =   "9999"
         Top             =   120
         Width           =   735
      End
   End
   Begin ComctlLib.ProgressBar Barra 
      Height          =   320
      Left            =   120
      TabIndex        =   54
      Tag             =   "0"
      Top             =   4560
      Visible         =   0   'False
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   0
      Min             =   77
      Max             =   91
   End
   Begin VB.ComboBox i_responsable 
      Height          =   315
      ItemData        =   "FORM_GRIFO.frx":0000
      Left            =   240
      List            =   "FORM_GRIFO.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   128
      Tag             =   "34"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.CommandButton Boton_Cons 
      BackColor       =   &H00808000&
      Caption         =   "&Ver Consignacion "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   127
      Tag             =   "189"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   960
      MouseIcon       =   "FORM_GRIFO.frx":0004
      Picture         =   "FORM_GRIFO.frx":014E
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   126
      Tag             =   "0"
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin ComctlLib.ListView LV_VEN2 
      Height          =   375
      Left            =   4440
      TabIndex        =   110
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ListView LV_VEN 
      Height          =   495
      Left            =   5640
      TabIndex        =   81
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ListView LV_CCM 
      Height          =   375
      Left            =   6480
      TabIndex        =   78
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ListView LV_CLI 
      Height          =   375
      Left            =   7320
      TabIndex        =   76
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ListView LV_ART 
      Height          =   375
      Left            =   8160
      TabIndex        =   77
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   6120
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox textovar 
      BackColor       =   &H00FFFF00&
      Height          =   405
      Left            =   600
      TabIndex        =   125
      Tag             =   "0"
      Text            =   "i_vendedor"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Items"
      ForeColor       =   &H00000000&
      Height          =   2325
      Left            =   2040
      TabIndex        =   18
      Tag             =   "100"
      Top             =   3600
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox i_cant 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3840
         TabIndex        =   89
         Tag             =   "9999"
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox i_autorizacion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   42
         TabStop         =   0   'False
         Tag             =   "9999"
         Text            =   "I_AUTORIZACION"
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   0
         TabIndex        =   67
         Tag             =   "119"
         Top             =   1560
         Width           =   6495
         Begin VB.TextBox i_moneda 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   525
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   119
            TabStop         =   0   'False
            Tag             =   "9999"
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox i_neto 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   65
            TabStop         =   0   'False
            Tag             =   "9999"
            Text            =   "i_neto"
            Top             =   300
            Width           =   1335
         End
         Begin VB.TextBox i_flete 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   405
            Left            =   4080
            TabIndex        =   64
            TabStop         =   0   'False
            Tag             =   "9999"
            Top             =   300
            Width           =   630
         End
         Begin VB.TextBox i_impto 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3000
            TabIndex        =   63
            TabStop         =   0   'False
            Tag             =   "9999"
            Text            =   "i_impto"
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox i_descto 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   62
            TabStop         =   0   'False
            Tag             =   "9999"
            Text            =   "i_descto"
            Top             =   300
            Width           =   750
         End
         Begin VB.TextBox i_gastos 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1320
            MaxLength       =   25
            TabIndex        =   61
            TabStop         =   0   'False
            Tag             =   "9999"
            Text            =   "i_gastos"
            Top             =   300
            Width           =   750
         End
         Begin VB.TextBox i_subtotal 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   60
            TabStop         =   0   'False
            Tag             =   "9999"
            Text            =   "i_subtotal"
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label LCODART 
            Caption         =   "Subtotal:"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   73
            Tag             =   "9999"
            Top             =   120
            Width           =   765
         End
         Begin VB.Label LCODART 
            Caption         =   "Gastos"
            Height          =   255
            Index           =   6
            Left            =   1440
            TabIndex        =   72
            Tag             =   "9999"
            Top             =   120
            Width           =   765
         End
         Begin VB.Label LCODART 
            Caption         =   "Descto."
            Height          =   255
            Index           =   5
            Left            =   2280
            TabIndex        =   71
            Tag             =   "9999"
            Top             =   120
            Width           =   765
         End
         Begin VB.Label LCODART 
            Caption         =   "Impto."
            Height          =   255
            Index           =   4
            Left            =   3240
            TabIndex        =   70
            Tag             =   "9999"
            Top             =   120
            Width           =   525
         End
         Begin VB.Label LCODART 
            Caption         =   "N e t o"
            Height          =   255
            Index           =   2
            Left            =   5400
            TabIndex        =   69
            Tag             =   "9999"
            Top             =   120
            Width           =   765
         End
         Begin VB.Label LCODART 
            Caption         =   "Flete"
            Height          =   255
            Index           =   1
            Left            =   3960
            TabIndex        =   68
            Tag             =   "9999"
            Top             =   120
            Width           =   765
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grid_fac 
         Height          =   1095
         Left            =   0
         TabIndex        =   44
         Tag             =   "9999"
         Top             =   120
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1931
         _Version        =   327680
         Rows            =   3
         Cols            =   9
         FixedRows       =   2
         GridLines       =   3
         AllowUserResizing=   3
      End
      Begin VB.Label label_precio 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3120
         TabIndex        =   90
         Tag             =   "9999"
         Top             =   960
         Width           =   885
      End
      Begin VB.Label label_nomart 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3120
         TabIndex        =   88
         Tag             =   "9999"
         Top             =   600
         Width           =   765
      End
      Begin VB.Image i_image_llave 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   6000
         Tag             =   "9999"
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.TextBox i_serguia 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3360
      TabIndex        =   124
      Tag             =   "64"
      Text            =   "i_serguia"
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox i_cambio 
      Caption         =   "Manual"
      Height          =   375
      Left            =   2640
      TabIndex        =   123
      Tag             =   "0"
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox i_fecha_compra 
      Height          =   285
      Left            =   3720
      TabIndex        =   122
      Tag             =   "36"
      Text            =   "i_fecha_compra"
      Top             =   120
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton Tipo_Cambio 
      BackColor       =   &H00808000&
      Caption         =   "&T/ Cambio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   121
      Tag             =   "167"
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox i_numser_r 
      Height          =   285
      Left            =   2400
      TabIndex        =   120
      TabStop         =   0   'False
      Tag             =   "88"
      Text            =   "i_numser_r"
      Top             =   120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox i_numfac 
      Height          =   400
      Left            =   3360
      TabIndex        =   117
      Tag             =   "26"
      Text            =   "i_numfac"
      Top             =   3120
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Frame Frame3 
      Caption         =   "Opción : "
      Height          =   1815
      Left            =   1080
      TabIndex        =   112
      Tag             =   "50"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
      Begin VB.OptionButton Option4 
         Caption         =   "Ninguno"
         Height          =   375
         Left            =   240
         TabIndex        =   118
         Tag             =   "9999"
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Fecha "
         Height          =   255
         Left            =   240
         TabIndex        =   115
         Tag             =   "9999"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Octanaje"
         Height          =   495
         Left            =   240
         TabIndex        =   114
         Tag             =   "9999"
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Placas"
         Height          =   375
         Left            =   240
         TabIndex        =   113
         Tag             =   "9999"
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.ComboBox i_situacion 
      Height          =   315
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   106
      Tag             =   "133"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Guia Ant."
      Height          =   375
      Left            =   5160
      TabIndex        =   105
      Tag             =   "125"
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grid_trans 
      Height          =   495
      Left            =   5400
      TabIndex        =   104
      Tag             =   "0"
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   327680
      Rows            =   50
      Cols            =   15
      BackColor       =   16777215
      BackColorFixed  =   12632256
      ForeColorFixed  =   64
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
   End
   Begin VB.ComboBox i_cias 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   3840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   103
      Tag             =   "13"
      Top             =   3480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox i_def 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "9999"
      Top             =   120
      Width           =   3600
   End
   Begin VB.ComboBox i_placas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   96
      Tag             =   "66"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1365
   End
   Begin RichTextLib.RichTextBox textovarl 
      Height          =   375
      Left            =   6600
      TabIndex        =   95
      Tag             =   "0"
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   327680
      BackColor       =   16776960
      MultiLine       =   0   'False
      TextRTF         =   $"FORM_GRIFO.frx":0298
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox textovar_canje 
      Height          =   375
      Left            =   2280
      TabIndex        =   94
      Tag             =   "0"
      Top             =   6000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   327680
      BackColor       =   16776960
      MultiLine       =   0   'False
      MousePointer    =   1
      TextRTF         =   $"FORM_GRIFO.frx":0363
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox i_tipdoc 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Tag             =   "16"
      Top             =   600
      Visible         =   0   'False
      Width           =   1005
   End
   Begin MSFlexGridLib.MSFlexGrid grid_canje 
      Height          =   495
      Left            =   8040
      TabIndex        =   53
      Tag             =   "105"
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   327680
      Rows            =   3
      Cols            =   7
      FixedRows       =   2
      BackColor       =   16777215
      BackColorFixed  =   12632256
      ForeColorFixed  =   64
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   3
   End
   Begin VB.ComboBox i_is 
      Height          =   315
      Left            =   4560
      TabIndex        =   93
      Tag             =   "0"
      Text            =   "Combo1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Boton_Recepcion 
      BackColor       =   &H00808000&
      Caption         =   "&Recepcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   92
      Tag             =   "158"
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid_all 
      Height          =   495
      Left            =   5880
      TabIndex        =   51
      Tag             =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   327680
      Rows            =   50
      Cols            =   15
      BackColor       =   16777215
      BackColorFixed  =   12632256
      ForeColorFixed  =   64
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
   End
   Begin VB.ListBox LisTransa 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5040
      TabIndex        =   41
      Tag             =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox i_dias 
      Height          =   285
      Left            =   1920
      TabIndex        =   87
      Tag             =   "43"
      Text            =   "i_dias"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox i_precios 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   79
      Tag             =   "0"
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox i_limcre 
      Height          =   285
      Left            =   0
      TabIndex        =   75
      Tag             =   "54"
      Text            =   "i_limcre"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.ComboBox i_mortal 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   74
      Tag             =   "0"
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox i_camal 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      MaxLength       =   255
      TabIndex        =   66
      Tag             =   "97"
      Text            =   "i_camal"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox i_num_lote 
      Height          =   285
      Left            =   4560
      TabIndex        =   59
      Tag             =   "91"
      Text            =   "I_num_lote"
      Top             =   2520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox i_TEXTONCRE 
      Height          =   285
      Left            =   5280
      TabIndex        =   58
      Tag             =   "90"
      Text            =   "i_textoncre"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox i_serdoc 
      Height          =   285
      Left            =   3840
      TabIndex        =   57
      Tag             =   "86"
      Text            =   "i_serdoc"
      Top             =   2760
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox i_numser_c 
      Height          =   285
      Left            =   3240
      TabIndex        =   56
      Tag             =   "84"
      Text            =   "i_numser_c"
      Top             =   2760
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox i_numfac_c 
      Height          =   285
      Left            =   1080
      TabIndex        =   55
      Tag             =   "85"
      Text            =   "i_numfac_c"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.ComboBox i_fbg 
      Height          =   315
      ItemData        =   "FORM_GRIFO.frx":042E
      Left            =   1320
      List            =   "FORM_GRIFO.frx":043B
      Style           =   2  'Dropdown List
      TabIndex        =   48
      Tag             =   "14"
      Top             =   3000
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton Diario 
      BackColor       =   &H00808000&
      Caption         =   "Diari&o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   84
      Tag             =   "153"
      Top             =   5280
      Width           =   855
   End
   Begin VB.ComboBox i_ds 
      Height          =   315
      ItemData        =   "FORM_GRIFO.frx":0448
      Left            =   4320
      List            =   "FORM_GRIFO.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Tag             =   "76"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1000
   End
   Begin MSFlexGridLib.MSFlexGrid grid_autorizacion 
      Height          =   495
      Left            =   7320
      TabIndex        =   46
      Tag             =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   327680
      Rows            =   50
      Cols            =   15
      BackColor       =   16777215
      BackColorFixed  =   12632256
      ForeColorFixed  =   64
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
      FormatString    =   " |#  |Usuario|Articulo|Precio|Factura|Cantidad|Tasa%|Dias|Cash|Est.|Codtra|   Fecha     |  hora        |"
   End
   Begin VB.CommandButton boton_autorizacion 
      Caption         =   "&Autoriz."
      Height          =   375
      Left            =   4920
      TabIndex        =   49
      Tag             =   "150"
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox i_num_oper 
      Height          =   285
      Left            =   7680
      TabIndex        =   45
      Tag             =   "73"
      Text            =   "I_num_oper"
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Boton_Letras 
      BackColor       =   &H00808000&
      Caption         =   "&Detalle Facturacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   43
      Tag             =   "151"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Cheques 
      BackColor       =   &H00FF00FF&
      Caption         =   "C&heques"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   85
      Tag             =   "152"
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox i_turno 
      Height          =   285
      Left            =   3600
      TabIndex        =   40
      Tag             =   "51"
      Text            =   "i_turno"
      Top             =   1200
      Visible         =   0   'False
      Width           =   700
   End
   Begin VB.TextBox i_placa 
      Height          =   405
      Left            =   3240
      TabIndex        =   39
      Tag             =   "35"
      Text            =   "i_placa"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox i_numser 
      Height          =   400
      Left            =   2280
      TabIndex        =   37
      Tag             =   "25"
      Text            =   "i_numser"
      Top             =   3120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton SALIR 
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   83
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cancelar 
      BackColor       =   &H00000040&
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   82
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton grabar 
      BackColor       =   &H00FF00FF&
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox i_situacion_ant 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   35
      Tag             =   "134"
      Top             =   4440
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.TextBox i_numplan 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   0
      TabIndex        =   28
      Tag             =   "37"
      Text            =   "i_diasV"
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox i_gastos_fijos 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   27
      Tag             =   "10"
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox i_importe_amort 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   0
      TabIndex        =   26
      Tag             =   "6"
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox i_numguia 
      Height          =   285
      Left            =   3120
      TabIndex        =   25
      Tag             =   "18"
      Text            =   "i_numguia"
      Top             =   480
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox i_pordescto1 
      Height          =   285
      Left            =   0
      TabIndex        =   19
      Tag             =   "33"
      Text            =   "i_pordescto1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox i_numdoc 
      Height          =   285
      Left            =   480
      TabIndex        =   16
      Tag             =   "42"
      Text            =   "i_numdoc"
      Top             =   360
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TRANS 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   600
      MaxLength       =   7
      TabIndex        =   1
      Tag             =   "9999"
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox i_importe 
      Height          =   285
      HideSelection   =   0   'False
      Left            =   0
      MaxLength       =   14
      TabIndex        =   6
      Tag             =   "7"
      Top             =   1200
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox i_codcli 
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Tag             =   "4"
      Text            =   "i_codcli"
      Top             =   840
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox i_codven 
      Height          =   285
      Left            =   0
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "41"
      Text            =   "i_codven"
      Top             =   480
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSFlexGridLib.MSFlexGrid gridl 
      Height          =   495
      Left            =   8160
      TabIndex        =   47
      Tag             =   "166"
      Top             =   480
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      _Version        =   327680
      Rows            =   50
      Cols            =   16
      FixedRows       =   2
      BackColor       =   16777215
      BackColorFixed  =   12632256
      ForeColorFixed  =   64
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
   End
   Begin VB.TextBox i_fecha_vcto 
      Height          =   285
      Left            =   0
      TabIndex        =   34
      Tag             =   "24"
      Text            =   "i_fecha_vcto"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1245
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   8760
      Tag             =   "9999"
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin VB.TextBox i_concepto 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   6360
      MaxLength       =   255
      TabIndex        =   52
      Tag             =   "20"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estados del Cliente"
      Height          =   2175
      Left            =   6840
      TabIndex        =   97
      Tag             =   "44"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
      Begin VB.CheckBox BLOQ 
         Caption         =   "Check4"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   101
         Tag             =   "0"
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox BLOQ 
         Caption         =   "Check3"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   100
         Tag             =   "0"
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox BLOQ 
         Caption         =   "Check2"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   99
         Tag             =   "0"
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox BLOQ 
         Caption         =   "Check1"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   98
         Tag             =   "0"
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid_liq 
      Height          =   495
      Left            =   8160
      TabIndex        =   109
      Tag             =   "177"
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   327680
      Rows            =   50
      Cols            =   15
      BackColor       =   16777215
      BackColorFixed  =   12632256
      ForeColorFixed  =   64
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
      FormatString    =   " |Amortizacion  |    Int. Vencido| Int. Adelantado|  Gastos F. |  Total | Dias Venc.| Dias Adelant.| "
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   3495
      Left            =   360
      TabIndex        =   129
      Tag             =   "111"
      Top             =   1680
      Visible         =   0   'False
      Width           =   9135
      Begin VB.Frame Frame6 
         Height          =   2175
         Left            =   6120
         TabIndex        =   141
         Tag             =   "9999"
         Top             =   360
         Width           =   2775
         Begin VB.TextBox Text1 
            Height          =   325
            Index           =   14
            Left            =   6480
            MaxLength       =   15
            TabIndex        =   169
            Tag             =   "0"
            Text            =   "Text1"
            Top             =   3000
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.TextBox Text1 
            Height          =   325
            Index           =   13
            Left            =   6480
            MaxLength       =   15
            TabIndex        =   168
            Tag             =   "0"
            Text            =   "Text1"
            Top             =   2520
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.TextBox Text1 
            Height          =   325
            Index           =   12
            Left            =   6480
            MaxLength       =   15
            TabIndex        =   167
            Tag             =   "0"
            Text            =   "Text1"
            Top             =   2040
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.TextBox Text1 
            Height          =   325
            Index           =   11
            Left            =   6480
            MaxLength       =   15
            TabIndex        =   166
            Tag             =   "0"
            Text            =   "Text1"
            Top             =   1560
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.TextBox Text1 
            Height          =   325
            Index           =   10
            Left            =   6480
            MaxLength       =   15
            TabIndex        =   161
            Tag             =   "0"
            Text            =   "Text1"
            Top             =   1080
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.TextBox Text1 
            Height          =   325
            Index           =   9
            Left            =   6480
            MaxLength       =   15
            TabIndex        =   159
            Tag             =   "0"
            Text            =   "Text1"
            Top             =   600
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.TextBox Text1 
            Height          =   325
            Index           =   8
            Left            =   6480
            MaxLength       =   15
            TabIndex        =   158
            Tag             =   "0"
            Text            =   "Text1"
            Top             =   120
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.TextBox Text1 
            Height          =   325
            Index           =   7
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   157
            Tag             =   "0"
            Text            =   "Text1"
            Top             =   3000
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Text1 
            Height          =   325
            Index           =   6
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   156
            Tag             =   "0"
            Text            =   "Text1"
            Top             =   2520
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Text1 
            Height          =   325
            Index           =   5
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   155
            Tag             =   "0"
            Text            =   "Text1"
            Top             =   2040
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Text1 
            Height          =   325
            Index           =   4
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   154
            Tag             =   "0"
            Text            =   "Text1"
            Top             =   1560
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Text1 
            Height          =   325
            Index           =   3
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   153
            Tag             =   "0"
            Text            =   "Text1"
            Top             =   1080
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Text1 
            Height          =   325
            Index           =   2
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   152
            Tag             =   "0"
            Text            =   "Text1"
            Top             =   600
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Text1 
            Height          =   325
            Index           =   1
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   151
            Tag             =   "0"
            Text            =   "Text1"
            Top             =   120
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   14
            Left            =   4440
            TabIndex        =   165
            Tag             =   "0"
            Top             =   3000
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   13
            Left            =   4440
            TabIndex        =   164
            Tag             =   "0"
            Top             =   2520
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   12
            Left            =   4440
            TabIndex        =   163
            Tag             =   "0"
            Top             =   1080
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   11
            Left            =   4440
            TabIndex        =   162
            Tag             =   "0"
            Top             =   1560
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   10
            Left            =   3840
            TabIndex        =   160
            Tag             =   "0"
            Top             =   2040
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   9
            Left            =   4440
            TabIndex        =   150
            Tag             =   "0"
            Top             =   600
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   8
            Left            =   4440
            TabIndex        =   149
            Tag             =   "0"
            Top             =   120
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   148
            Tag             =   "0"
            Top             =   3000
            Visible         =   0   'False
            Width           =   1900
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   147
            Tag             =   "0"
            Top             =   2520
            Visible         =   0   'False
            Width           =   1900
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   146
            Tag             =   "0"
            Top             =   1920
            Visible         =   0   'False
            Width           =   1900
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   145
            Tag             =   "0"
            Top             =   1440
            Visible         =   0   'False
            Width           =   1900
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   144
            Tag             =   "0"
            Top             =   1080
            Visible         =   0   'False
            Width           =   1900
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   143
            Tag             =   "0"
            Top             =   600
            Visible         =   0   'False
            Width           =   1900
         End
         Begin VB.Label Label1 
            Caption         =   "123456789012345"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   142
            Tag             =   "0"
            Top             =   120
            Visible         =   0   'False
            Width           =   1905
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridD 
         Height          =   975
         Left            =   4320
         TabIndex        =   131
         Tag             =   "9999"
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1720
         _Version        =   327680
         Rows            =   3
         FixedRows       =   2
      End
      Begin MSFlexGridLib.MSFlexGrid GridT 
         Height          =   2415
         Left            =   120
         TabIndex        =   130
         Tag             =   "9999"
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   4260
         _Version        =   327680
         Rows            =   3
         FixedRows       =   2
      End
   End
   Begin VB.Label label_grid 
      BackColor       =   &H00000000&
      Caption         =   "Ordenes de Pedido Pendientes  <Seleccionar con ENTER>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   7200
      TabIndex        =   102
      Tag             =   "0"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label i_nomturno 
      Caption         =   "i_nomven"
      Height          =   255
      Left            =   3240
      TabIndex        =   116
      Tag             =   "0"
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label i_nomven2 
      Caption         =   "i_nomven"
      Height          =   255
      Left            =   6240
      TabIndex        =   111
      Tag             =   "0"
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line RAYA 
      BorderWidth     =   2
      Tag             =   "9999"
      Visible         =   0   'False
      X1              =   120
      X2              =   8880
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label LMONEDAB 
      Caption         =   "Lmoneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   108
      Tag             =   "0"
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label i_nomban2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "i_nomban2"
      Height          =   255
      Left            =   4440
      TabIndex        =   107
      Tag             =   "0"
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Lmoneda 
      Caption         =   "Lmoneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   91
      Tag             =   "182"
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label BolFac 
      BackColor       =   &H00FFFF80&
      Caption         =   "bol/fact."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   400
      Left            =   2880
      TabIndex        =   86
      Tag             =   "9999"
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label i_nomart 
      BackColor       =   &H00C0C0C0&
      Caption         =   "i_nomart"
      Height          =   255
      Left            =   4680
      TabIndex        =   36
      Tag             =   "0"
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   3
      Index           =   1
      Tag             =   "9999"
      X1              =   -480
      X2              =   9480
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   3
      Index           =   0
      Tag             =   "9999"
      X1              =   0
      X2              =   9480
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   17
      Left            =   6360
      TabIndex        =   33
      Tag             =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   16
      Left            =   6360
      TabIndex        =   32
      Tag             =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   15
      Left            =   6480
      TabIndex        =   31
      Tag             =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   14
      Left            =   6480
      TabIndex        =   30
      Tag             =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   13
      Left            =   6480
      TabIndex        =   29
      Tag             =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   12
      Left            =   6480
      TabIndex        =   24
      Tag             =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   11
      Left            =   6480
      TabIndex        =   23
      Tag             =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   10
      Left            =   6480
      TabIndex        =   22
      Tag             =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   9
      Left            =   6480
      TabIndex        =   21
      Tag             =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   8
      Left            =   6480
      TabIndex        =   20
      Tag             =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   7
      Left            =   6480
      TabIndex        =   15
      Tag             =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   6
      Left            =   6500
      TabIndex        =   14
      Tag             =   "0"
      Top             =   100
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   5
      Left            =   6500
      TabIndex        =   13
      Tag             =   "0"
      Top             =   100
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   4
      Left            =   6500
      TabIndex        =   12
      Tag             =   "0"
      Top             =   100
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   3
      Left            =   6500
      TabIndex        =   11
      Tag             =   "0"
      Top             =   100
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   2
      Left            =   6500
      TabIndex        =   10
      Tag             =   "0"
      Top             =   100
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   1
      Left            =   6500
      TabIndex        =   9
      Tag             =   "0"
      Top             =   100
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      Caption         =   "LABELGEN"
      Height          =   495
      Index           =   0
      Left            =   6500
      TabIndex        =   8
      Tag             =   "0"
      Top             =   100
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label i_nomven 
      Caption         =   "i_nomven"
      Height          =   255
      Left            =   4320
      TabIndex        =   17
      Tag             =   "0"
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label i_nomcli 
      Caption         =   "i_nomcli"
      DataField       =   "1234"
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Tag             =   "0"
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label nomtra 
      BackColor       =   &H00C0C0C0&
      Caption         =   "nomtra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Tag             =   "9999"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "FORM_GRIFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CAMTEX
Dim pub_mensaje_err As String
Dim WGUIA  As String
Dim WW_FECHA_ANT As Date
Dim CONTADOR As Integer
Dim ww_codtra_ext As Integer
Dim tab_derL(10) As Integer
Dim tab_izqL(10) As Integer
Dim tab_derecha(38) As Integer
Dim tab_izquierda(38) As Integer
Dim tab_turnos(4) As String * 20
Dim tab_items(20, 20) As Integer
Dim pub_numfac_ult As Long
Dim PSFFF_ART As rdoQuery
Dim FFF_ART As rdoResultset

Dim FAR_TRANS3 As rdoResultset
Dim PSFAR_TRANS3 As rdoQuery

Dim pdat_llave As rdoResultset
Dim PSDAT_LLAVE As rdoQuery

Dim PF_BUSCA As rdoQuery
Dim pff_buscar As rdoResultset

Dim PSFAR_TEXT As rdoQuery
Dim FAR_TEXT As rdoResultset

Dim PSFAR_TRANS As rdoQuery
Dim FAR_TRANS As rdoResultset

Dim PSFAR_TRANS2 As rdoQuery
Dim FAR_TRANS2 As rdoResultset

Dim PSPED_MAYOR As rdoQuery
Dim CONS_llave As rdoResultset
Dim PSCONS_LLAVE As rdoQuery
Dim PS_FOR As rdoQuery
Dim llave_FOR As rdoResultset
Dim ped_mayor As rdoResultset

Dim ped_menor As rdoResultset
Dim PSPED_MENOR As rdoQuery

Dim PS_PLACAS As rdoQuery
Dim TAB_PLACAS As rdoResultset
Dim PS_VE2 As rdoQuery
Dim VE2_LLAVE As rdoResultset
Dim PSCAR_GRIFO As rdoQuery
Dim CAR_GRIFO As rdoResultset
Dim PS_PRECIOS As rdoQuery
Dim TAB_PRECIOS As rdoResultset
Dim pub_responsable As Integer
Dim pub_montodescto As Currency
Dim WS_NUMSEC As Integer
Public textovar_bak As String
Dim flag_salto As Integer
Dim loc_key As Integer
Dim WS_TIPO_CAMBIO  As Currency
Dim wtemporal As String * 1
Dim xl As Object
Dim WS_MONEDA_CLI As String * 1
Dim WS_MONEDA_CCM As String * 1
Dim VAR_ACTIVAR As Integer
Dim PLAZA_FLAG_MANUAL As String * 1
Dim ext_conta As Integer
Dim cliplac_llave   As rdoResultset
Dim PSPLAC_LLAVE  As rdoQuery
Dim cliplac_secu As rdoResultset
Dim PSPLAC_SECU As rdoQuery
Public flag_textil As Integer
Option Explicit
Public Sub RECALCULA()
Dim WW_TOT As Currency
Dim xx  As Integer

   PUB_CONCEPTO = Trim(GridD.TextMatrix(GridD.Row, 0))
   WW_TOT = 0
   xx = 2
   Do Until xx > GridD.Rows - 1
      If PUB_CONCEPTO = Trim(GridD.TextMatrix(xx, 0)) Then
         WW_TOT = WW_TOT + Val(GridD.TextMatrix(xx, 6))
      End If
      xx = xx + 1
   Loop
   xx = 2
   Do Until xx > GridD.Rows - 1
      If PUB_CONCEPTO = Trim(GridD.TextMatrix(xx, 0)) Then
         GridD.TextMatrix(xx, 2) = WW_TOT
      End If
      xx = xx + 1
   Loop
calcula_totales24

End Sub
Public Sub RECALCULA_TODO()
Dim xx As Integer

   xx = 2
   PUB_CONCEPTO = ""
   Do Until xx > GridD.Rows - 1
      
      If PUB_CONCEPTO <> Trim(GridD.TextMatrix(xx, 0)) Then
         GridD.Row = xx
         RECALCULA
      End If
      PUB_CONCEPTO = Trim(GridD.TextMatrix(xx, 0))

      xx = xx + 1
      
      
   Loop
   
calcula_totales24

End Sub



Private Sub LLENA_CAMPOS()
Dim ws_codcia As String * 2
lmoneda.Visible = False
lmoneda.Visible = False
Dim indice
Dim wxtabindex
Dim enlace
Dim pos1
Dim CARAC
Dim cias As String
nn = 2
m_ind = 0
wxtabindex = 2
Do Until Val(tra_llave(nn)) = 0 Or nn = 62

m_ind = m_ind + 1
If LABELGEN(m_ind) <> "0" Then
   LABELGEN(m_ind).Visible = True
   If tra_llave(nn + 2) > 1200 And tra_llave(nn + 3) > 7000 Then
      LABELGEN(m_ind).Top = tra_llave(nn + 2) - 200
      LABELGEN(m_ind).Left = tra_llave(nn + 3)
   Else
      LABELGEN(m_ind).Top = tra_llave(nn + 2)
      LABELGEN(m_ind).Left = tra_llave(nn + 3) - 1000
   End If
   LABELGEN(m_ind).Caption = tra_llave(nn + 1)
End If

indice = TABLA_TAG(tra_llave(nn))

   If TypeOf Controls(indice) Is MSFlexGrid Then
      LABELGEN(m_ind).Caption = ""
      pos1 = InStr(1, tra_llave(nn + 1), ",", 1)
      pos1 = pos1 - 1
      If pos1 = -1 Then
         pos1 = 0
      End If
      CARAC = Left(tra_llave(nn + 1), pos1)
      Controls(indice).Width = Val(CARAC)
      pos1 = pos1 + 2
      CARAC = Mid(tra_llave(nn + 1), pos1, 10)
      DoEvents
      Controls(indice).Height = Val(CARAC)
    End If

    
If TypeOf Controls(indice) Is Label Then
    enlace = 1
Else
    enlace = 0
End If
 

    Controls(indice).Top = tra_llave(nn + 2)
    Controls(indice).Left = tra_llave(nn + 3) + 1300 * enlace
    Controls(indice).Visible = True
    If TypeOf Controls(indice) Is MSFlexGrid Then
       Controls(indice).Clear
    End If
   
    
    Controls(indice).Enabled = True
    If tra_llave(106 + m_ind) = -1 Then
       Controls(indice).Enabled = False
    End If
    If tra_llave(106 + m_ind) = -2 Then
       Controls(indice).Locked = True
    End If
    
    
    
    
    If Val(Controls(indice).Tag) = 4 Then
       i_nomCLI.Top = tra_llave(nn + 2)
       i_nomCLI.Left = tra_llave(nn + 3) + 1300 * 1
       i_nomCLI.Visible = True
       i_nomCLI.Caption = ""
       i_nomCLI.Width = 3000
    End If
    If Val(Controls(indice).Tag) = 41 Then
       i_nomven.Top = tra_llave(nn + 2)
       i_nomven.Left = tra_llave(nn + 3) + 1300 * 1
       i_nomven.Visible = True
       i_nomven.Caption = ""
       i_nomven.Width = 3000
    End If
    If Val(Controls(indice).Tag) = 34 Then
       i_nomven2.Top = tra_llave(nn + 2)
       i_nomven2.Left = tra_llave(nn + 3) + 900 * 1
       i_nomven2.Visible = True
       i_nomven2.Caption = ""
       i_nomven2.Width = 2000
    End If
    
    If Val(Controls(indice).Tag) = 51 Then
       i_nomturno.Top = tra_llave(nn + 2)
       i_nomturno.Left = tra_llave(nn + 3) + 900 * 1
       i_nomturno.Visible = True
       i_nomturno.Caption = ""
       i_nomturno.Width = 2000
    End If

    
    If Val(Controls(indice).Tag) = 7 Then
       LMONEDAB.Top = tra_llave(nn + 2)
       LMONEDAB.Left = tra_llave(nn + 3) - 300
       LMONEDAB.Visible = True
       LMONEDAB.Caption = ""
    End If
    If Val(Controls(indice).Tag) = 6 Then
       lmoneda.Top = tra_llave(nn + 2)
       lmoneda.Left = tra_llave(nn + 3) - 300
       lmoneda.Visible = True
       lmoneda.Caption = ""
    End If
    
    If Val(Controls(indice).Tag) = 70 Then
       i_nomban2.Top = tra_llave(nn + 2)
       i_nomban2.Left = tra_llave(nn + 3) + 1300 * 1
       i_nomban2.Visible = True
       i_nomban2.Caption = ""
       i_nomban2.Width = 4500
    End If
    If Val(Controls(indice).Tag) = 5 Then
       i_nomart.Top = tra_llave(nn + 2)
       i_nomart.Left = tra_llave(nn + 3) + 1300 * 1
       i_nomart.Visible = True
       i_nomart.Caption = ""
       i_nomart.Width = 3000
    End If
    
    
    If enlace = 0 Then
       Controls(indice).WhatsThisHelpID = m_ind
       wxtabindex = wxtabindex + 1
'       MsgBox Controls(indice).Name
       Controls(indice).TabIndex = wxtabindex
       If Trim(Nulo_Valors(tra_llave(106 + m_ind))) = Null Then
          tab_avanza(m_ind) = 0
       Else
       tab_avanza(m_ind) = tra_llave(106 + m_ind)
       End If
    Else
       Controls(indice).Width = 3200
    End If
    If TypeOf Controls(indice) Is TextBox Then
        Controls(indice).Text = ""
    End If

nn = nn + 4
Loop
Picture1.Visible = False

i_cambio.Top = grabar.Top
i_cambio.Left = grabar.Left - 3500
i_cambio.Visible = True

If LK_CODTRA = 2105 Then
   Picture1.Top = grabar.Top
   Picture1.Left = grabar.Left + 4600
   Picture1.Visible = True
End If
i_fecha_compra.Text = LK_FECHA_DIA
grid_canje.Visible = False

If Frame4.Visible = True Then
   grid_fac.Clear
   i_moneda.Visible = True
   fila = 0
End If
If FORM_GRIFO.i_cias.Visible = True Then
   If Trim(par_llave!PAR_CIAS) <> "" Then
      cias = par_llave!PAR_CIAS
      nn = 1
    For m_ind = 1 To 15
        ws_codcia = Mid(cias, nn, 2)
        If Trim(ws_codcia) = "" Then Exit For
           SQ_OPER = 1
           PUB_CODCIA = ws_codcia
           LEER_PAR_LLAVE
           FORM_GRIFO.i_cias.AddItem par_llave!PAR_NOMBRE & "               " & par_llave!PAR_CODCIA
        nn = nn + 2
    Next m_ind
   End If
   SQ_OPER = 1
   PUB_CODCIA = LK_CODCIA
   LEER_PAR_LLAVE
End If



   i_subtotal.Text = ""
   i_gastos.Text = ""
   i_descto.Text = ""
   i_impto.Text = ""
   i_neto.Text = ""


If i_tipdoc.Visible = True Then
i_tipdoc.Clear
PUB_TIPREG = 8
PUB_CODCIA = "00"
SQ_OPER = 2
LEER_TAB_LLAVE
If tab_mayor.EOF Then
   MsgBox "LLenar tabla de tipo de documentos...LE, CH,ETC"
End If
Do Until tab_mayor.EOF
  DoEvents
  i_tipdoc.AddItem Nulo_Valors(tab_mayor!tab_nomcorto)
  tab_mayor.MoveNext
Loop
If i_tipdoc.ListCount > 0 Then i_tipdoc.ListIndex = 0


End If

If LK_CODTRA = 1111 Then i_fecha_vcto.Text = LK_FECHA_DIA

i_fecha_vcto.Text = LK_FECHA_DIA

WS_BRUTO = 0
SUB_CANT = 0
SUB_JABAS = 0
SUB_UNIDAD = 0
fila = 0

If LK_CODTRA = 2103 Then
   Frame6.Width = Frame5.Width
   Frame6.Top = 0
   Frame6.Left = Frame5.Left
   Frame6.Height = 3500
   Frame6.Visible = True
   GridT.Visible = False
   GridD.Visible = False
   GridT.Top = 200
   GridD.Top = 300
   GridT.Left = 0
   GridD.Left = 0
End If

fin:


End Sub
Private Sub BUSCA_RESP()
inicio:
PS_VE2.rdoParameters(0) = LK_CODCIA
PS_VE2.rdoParameters(1) = Val(i_codven.Text)
PS_VE2.rdoParameters(2) = Val(i_turno.Text)
VE2_LLAVE.Requery
If VE2_LLAVE.EOF Then
   VE2_LLAVE.AddNew
   VE2_LLAVE!VE2_CODCIA = LK_CODCIA
   VE2_LLAVE!VE2_ISLA = Val(i_codven.Text)
   VE2_LLAVE!VE2_TURNO = Val(i_turno.Text)
   VE2_LLAVE!VE2_CODRES = 0
   VE2_LLAVE.Update
   GoTo inicio
End If

FILAX = 0
Do Until FILAX = i_responsable.ListCount
   i_responsable.ListIndex = FILAX
   If Val(Right(i_responsable.Text, 5)) = VE2_LLAVE!VE2_CODRES Then Exit Do
   FILAX = FILAX + 1
Loop
If FILAX = i_responsable.ListCount Then
   i_responsable.ListIndex = -1
   i_responsable.SetFocus
End If


End Sub

Public Sub VERIFICA_REPET()
Dim i As Integer
Dim J As Integer
J = grid_fac.Row
For i = 2 To grid_fac.Rows - 1
If i <> J Then
   If grid_fac.TextMatrix(i, 16) = art_LLAVE!ART_KEY Then MsgBox "Ojo ...Articulo ya existe en lista", 48, Pub_Titulo
End If

Next i

End Sub
Public Sub VERIFICA_REPET2()
Dim i As Integer
Dim J As Integer
J = GridT.Row
For i = 2 To GridT.Row - 1
If GridT.TextMatrix(i, 5) = art_LLAVE!ART_KEY Then
   MsgBox "Ojo ...Articulo ya existe en lista...No Procede ", 48, Pub_Titulo
   If PUB_TIPMOV = 177 Then
    GridT.TextMatrix(GridT.Row, 0) = ""
    GridT.TextMatrix(GridT.Row, 1) = ""
   End If
   textovar2.Text = ""
End If
Next i

End Sub



Public Sub LLENADOS(cont As ListBox, tip As Integer)

    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
       cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
       tab_mayor.MoveNext
    Loop
End Sub

Public Sub cancela_todito()
Dim indice
nn = 2
m_ind = 0
Do Until Val(tra_llave(nn)) = 0 Or nn = 62 '

m_ind = m_ind + 1

indice = TABLA_TAG(tra_llave(nn))

If TypeOf Controls(indice) Is MSFlexGrid Then
   If Controls(indice).Visible = True Then
      Controls(indice).Clear
      Controls(indice).Rows = 3
   End If
End If

nn = nn + 4
Loop

i_importe.Text = ""
If Frame4.Visible = True Then
   i_subtotal.Text = ""
   i_gastos.Text = ""
   i_descto.Text = ""
   i_impto.Text = ""
   i_neto.Text = ""
   i_cant.Text = ""
   i_flete.Text = ""
   i_dias.Text = ""
   grid_fac.Clear
   fila = 0
   i_cant.Locked = False
   textovar.Visible = False
End If
textovarl.Visible = False
i_nomCLI.Caption = ""
i_limcre.Text = ""
i_codcli.Text = ""
i_numguia.Text = ""
i_placa.Text = ""
'If i_codcli.Visible = True Then gridl.Visible = False
textovar_canje.Visible = False
'grid_che.Visible = False
'i_fbg.Enabled = True
i_dias.Enabled = True
i_fecha_vcto.Enabled = True
i_dias.Enabled = True
i_cant.Locked = False
textovarl.Visible = False
textovar_canje.Visible = False
gridl.Visible = False
grid_canje.Visible = False

i_fecha_vcto.Text = LK_FECHA_DIA
WS_BRUTO = 0
fila = 0

If SUT_LLAVE.EOF = False Then
   pasa_def
End If
FILAX = 0
fin:

End Sub

Public Sub VERIFICA_VENTA()

End Sub
Public Sub verifica_precios()
Dim WS_PRECIO As Currency
Dim ws_int As Currency
Dim ww_num As Integer
SQ_OPER = 1

pu_codcia = LK_CODCIA
PUB_CODART = art_LLAVE!ART_KEY
LEER_ARM_LLAVE

pu_codcia = LK_CODCIA
PUB_CODART = art_LLAVE!ART_KEY
LEER_PRE_LLAVE


PS_PRECIOS(0) = Val(i_codcli.Text)
PS_PRECIOS(1) = LK_CODCIA
PS_PRECIOS(2) = art_LLAVE!ART_KEY
TAB_PRECIOS.Requery





If pre_llave.EOF Then Exit Sub

If TAB_PRECIOS.EOF Then
   WS_PRECIO = pre_llave!PRE_PRE1
Else
   If PUB_SECUENCIA = 20 Then
      WS_PRECIO = Val(TAB_PRECIOS!tab_nomcorto)
   ElseIf PUB_SECUENCIA = 25 Then
     WS_PRECIO = pre_llave!PRE_PRE1
   Else
     WS_PRECIO = Val(TAB_PRECIOS!TAB_NOMLARGO)
   End If
End If

grid_fac.TextMatrix(grid_fac.Row, 6) = WS_PRECIO
grid_fac.TextMatrix(grid_fac.Row, 36) = WS_PRECIO
grid_fac.TextMatrix(grid_fac.Row, 35) = pre_llave!PRE_PRE1
grid_fac.TextMatrix(grid_fac.Row, 33) = PUB_CODART
  

End Sub
Public Sub verifica_unidades()
Dim cont As Integer
i_unidades.Clear
Dim longit As String * 15
Dim ww As String * 15
ww = "Unidades"
cont = -1
SQ_OPER = 2
PUB_CODART = art_LLAVE!ART_KEY
LEER_PRE_LLAVE
Do Until pre_mayor.EOF
   cont = cont + 1
   longit = pre_mayor!pre_UNIDAD
   i_unidades.AddItem longit & "       " & pre_mayor!pre_secuencia & "       " & pre_mayor!pre_equiv
   If pre_mayor!PRE_FLAG_UNIDAD = "A" Then i_unidades.ListIndex = cont
   pre_mayor.MoveNext
Loop
If PUB_TIPMOV = 20 And i_unidades.ListCount <> 0 Then i_unidades.ListIndex = i_unidades.ListCount - 1

End Sub
Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii As Integer, Optional ConsisVal, Optional ConsisCol)
  Static valor
  Dim car As String
  If ConsisVal = 2 Then ' NUMEROS CON DECIMALES
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
'      Flag_Consis = "A"
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.Text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." Then
      
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If
  ElseIf ConsisVal = 1 Then ' NUMEROS ENTEROS
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And wsKeyAscii <> 32 Then
          wsKeyAscii = 0
          Beep
        End If
      End If
  End If

End Sub
Public Sub llena_numfac_grifo()
Dim FLAG As Integer
Dim CONTADOR As Integer

PUB_FBG = i_fbg.Text

If i_fbg.Visible = False Then
   PUB_FBG = ""
   GoTo SALTA
End If



If PUB_TIPMOV = 10 Then
If LK_FLAG_FACTURACION = "V" Then
   If ven_llave.EOF Then
      MsgBox "Falta Vendedor ..."
      Exit Sub
   End If
   Select Case PUB_FBG
   Case "G"
       i_numser.Text = ven_llave!VEM_SERIE_G
   Case "B"
       i_numser.Text = ven_llave!VEM_SERIE_B
   Case "F"
       i_numser.Text = ven_llave!VEM_SERIE_F
   Case "P"
       i_numser.Text = ven_llave!VEM_SERIE_P
   End Select
Else
   Select Case PUB_FBG
   Case "G"
       i_numser.Text = par_llave!PAR_G_SERIE
   Case "B"
       i_numser.Text = par_llave!PAR_B_SERIE
   Case "F"
       i_numser.Text = par_llave!PAR_F_SERIE
   End Select

End If
End If


SALTA:
If PUB_TIPMOV = 20 Then PUB_FBG = ""
If PUB_TIPMOV = 99 Then PUB_FBG = "K"

If PUB_TIPMOV = 97 Then
   PUB_FBG = "N"
   i_numser.Text = par_llave!PAR_SERIE_NCRE
End If
If PUB_TIPMOV = 97 And PUB_CP = "P" Then
   PUB_FBG = "C"
   i_numser.Text = 0
End If

If PUB_TIPMOV = 98 Then
   PUB_FBG = "D"
   i_numser.Text = Nulo_Valor0(par_llave!PAR_SERIE_NDEB)
End If
If PUB_TIPMOV = 98 And PUB_CP = "P" Then
   PUB_FBG = "A"
   i_numser.Text = 0
End If
If PUB_TIPMOV = 177 And PUB_CP = "C" Then
   PUB_FBG = "O"
   i_numser.Text = Nulo_Valor0(par_llave!PAR_SERIE_P)
End If

If LK_CODTRA = 1122 Or LK_CODTRA = 1111 Then
   PUB_FBG = "X"
  i_numser.Text = ""
End If

SQ_OPER = 3
PU_TIPMOV = PUB_TIPMOV
pu_codcia = LK_CODCIA
PU_NUMSER = Val(i_numser.Text)

PU_FBG = PUB_FBG
LEER_FAR_LLAVE
PU_NUMFAC = 1
If Not far_menor.EOF Then
   PU_NUMFAC = Val(far_menor!FAR_NUMFAC) + 1
End If
If PUB_TIPMOV = 99 Then GoTo fin

 If i_fbg.Visible = False Then
   If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
   Else
      GoTo fin
   End If
End If

If LK_FLAG_FACTURACION = "V" Then
   If Not ven_llave.EOF Then
   If (ven_llave!VEM_FLAG_G = "A" Or ven_llave!VEM_FLAG_B = "A" Or ven_llave!VEM_FLAG_F = "A" Or ven_llave!VEM_FLAG_P = "A") Then
   Select Case PUB_FBG
   Case "G"
        If ven_llave!VEM_FLAG_G = "A" Then
           PU_NUMFAC = ven_llave!VEM_NUMFAC_G_INI
        End If
   Case "B"
        If ven_llave!VEM_FLAG_B = "A" Then
           PU_NUMFAC = ven_llave!VEM_NUMFAC_B_INI
        End If

   Case "F"
        If ven_llave!VEM_FLAG_F = "A" Then
           PU_NUMFAC = ven_llave!VEM_NUMFAC_F_INI
        End If
   Case "P"
        If ven_llave!VEM_FLAG_P = "A" Then
           PU_NUMFAC = ven_llave!VEM_NUMFAC_P_INI
        End If
   End Select
 End If
 End If
 End If
 
If LK_FLAG_FACTURACION = "A" And (par_llave!PAR_FLAG_G = "A" Or par_llave!PAR_FLAG_B = "A" Or par_llave!PAR_FLAG_F = "A") Then
   Select Case PUB_FBG
   Case "G"
       If par_llave!PAR_FLAG_G = "A" Then PU_NUMFAC = par_llave!PAR_G_INICIAL
   Case "B"
       If par_llave!PAR_FLAG_B = "A" Then PU_NUMFAC = par_llave!PAR_B_INICIAL
   Case "F"
       If par_llave!PAR_FLAG_F = "A" Then PU_NUMFAC = par_llave!PAR_F_INICIAL
   End Select
End If

NOTACRE:
If PUB_TIPMOV = 97 Then
 If par_llave!PAR_FLAG_NCRE = "A" Then
   i_numser.Text = par_llave!PAR_SERIE_NCRE
   PU_NUMFAC = par_llave!PAR_numfac_ncre
   PUB_FBG = "N"
 End If
End If
If PUB_TIPMOV = 98 Then
 If par_llave!PAR_FLAG_NDEB = "A" Then
   i_numser.Text = par_llave!PAR_SERIE_NDEB
   PU_NUMFAC = par_llave!PAR_numfac_nDEB
   PUB_FBG = "D"
 End If
End If
      
      
fin:
If PUB_FBG = "R" Then
   i_numguia.Text = PU_NUMFAC
   i_serguia.Text = i_numser.Text
   PUB_FBG = i_fbg.Text
   Exit Sub
End If

If PUB_TIPMOV = 20 And LK_CODTRA = 1401 Then
 If Val(par_llave!PAR_SER_KARDEX) <> Val(i_numser.Text) Then
  par_llave.Edit
  par_llave!PAR_SER_KARDEX = Val(i_numser.Text)
  par_llave.Update
 End If
End If


PUB_NUMSER = Val(i_numser.Text)
i_numfac.Text = PU_NUMFAC
PUB_NUMFAC = Val(PU_NUMFAC)
pub_numfac_ult = PUB_NUMFAC
If Trim(PUB_FBG) = "" Then
   BolFac.Caption = "Guia: "
Else
   BolFac.Caption = PUB_FBG & "/. "
End If

End Sub
Public Sub llena_pedido()
Dim FLAG As Integer
Dim CONTADOR As Integer

PSPED_MENOR.rdoParameters(0) = LK_CODCIA
PSPED_MENOR.rdoParameters(1) = PUB_TIPMOV
PSPED_MENOR.rdoParameters(2) = Val(i_ser_lote.Text) ' Nulo_Valor0(par_llave!PAR_SERIE_P)
PU_NUMSER = Val(i_ser_lote.Text)  'Nulo_Valor0(par_llave!PAR_SERIE_P)

ped_menor.Requery
PU_NUMFAC = 1
If Not ped_menor.EOF Then
   PU_NUMFAC = Val(ped_menor!PED_NUMFAC) + 1
End If

i_num_lote.Text = PU_NUMFAC
BolFac.Caption = "O/P " & PU_NUMFAC
PUB_PEDSER = PU_NUMSER
PUB_PEDFAC = PU_NUMFAC
i_num_lote.Locked = True
i_ser_lote.Locked = True
i_ser_lote.BackColor = QBColor(7)
i_num_lote.BackColor = QBColor(7)
i_cambio.Value = 0
If i_num_lote.Visible Then i_num_lote.SetFocus
End Sub

Public Sub llena_numGUIA()
Dim FLAG As Integer
Dim CONTADOR As Integer




If LK_FLAG_FACTURACION = "V" Then
   i_serguia.Text = ven_llave!VEM_SERIE_R
Else
   i_serguia.Text = par_llave!PAR_R_SERIE
End If

If LK_CODTRA = 1122 Or LK_CODTRA = 1111 Then
  i_serguia.Text = ""
End If

PUB_SERGUIA = Val(i_serguia.Text)
SQ_OPER = 7
pu_codcia = LK_CODCIA
LEER_FAR_LLAVE
PU_NUMFAC = 1
If Not far_guiam.EOF Then
   PU_NUMFAC = Val(far_guiam!far_NUMGUIA) + 1
End If


If LK_FLAG_FACTURACION = "V" Then
   If Not ven_llave.EOF Then
   If ven_llave!VEM_FLAG_R = "A" Then
        If ven_llave!VEM_FLAG_R = "A" Then PU_NUMFAC = ven_llave!VEM_NUMFAC_R_INI
   End If
   End If
End If
 
If LK_FLAG_FACTURACION = "A" And par_llave!PAR_FLAG_R = "A" Then
   PU_NUMFAC = par_llave!PAR_R_INICIAL
End If

      
fin:
i_numguia.Text = PU_NUMFAC


End Sub

Private Sub Consistencias_rich(wsGrid As MSFlexGrid, wsTexto As RichTextBox, wsKeyAscii As Integer, Optional ConsisVal, Optional ConsisCol)
  Static valor
  Dim car As String
  If ConsisVal = 2 Then ' NUMEROS CON DECIMALES
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
'      Flag_Consis = "A"
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.Text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." Then
      
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If
  ElseIf ConsisVal = 1 Then ' NUMEROS ENTEROS
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And wsKeyAscii <> 32 Then
          wsKeyAscii = 0
          Beep
        End If
      End If
  End If

End Sub
Public Sub PROCESA_GRIDL()
Dim WS_SIGNO As Integer
Dim WS_TIPDOC As String * 2
Dim NUMCAMPO As Integer
Dim TOTAL_DEUDA, TOTAL_GEN  As Variant
TOTAL_DEUDA = 0
TOTAL_GEN = 0
Dim Tit As String
gridl.Visible = True
gridl.Clear
pasa_cabeza_L
grid_canje.Clear
pasa_cabeza_canje
If Val(i_codcli.Text) = 0 Then Exit Sub

If cli_llave.EOF Then Exit Sub
pub_cadena = "SELECT * FROM CARTERA WHERE  CAR_CODCIA = ? AND CAR_CP = ? AND CAR_CODCLIE = ?  AND CAR_NUMFAC_C = 0 ORDER BY CAR_FECHA_INGR"
Frame4.Visible = False
If Option1.Value = True Then
  pub_cadena = "SELECT * FROM CARTERA WHERE  CAR_CODCIA = ? AND CAR_CP = ? AND CAR_CODCLIE = ?  AND CAR_NUMFAC_C = 0 ORDER BY CAR_PLACA, CAR_FECHA_INGR"
End If
If Option2.Value = True Then
  pub_cadena = "SELECT * FROM CARTERA WHERE  CAR_CODCIA = ? AND CAR_CP = ? AND CAR_CODCLIE = ?  AND CAR_NUMFAC_C = 0 ORDER BY CAR_CODART, CAR_FECHA_INGR"
End If
Set PSCAR_GRIFO = CN.CreateQuery("", pub_cadena)
PSCAR_GRIFO(0) = LK_CODCIA
PSCAR_GRIFO(1) = " "
PSCAR_GRIFO(2) = 0
Set CAR_GRIFO = PSCAR_GRIFO.OpenResultset(rdOpenKeyset, rdConcurValues)
PSCAR_GRIFO.rdoParameters(0) = LK_CODCIA
PSCAR_GRIFO.rdoParameters(1) = "C"
PSCAR_GRIFO.rdoParameters(2) = pu_codclie
CAR_GRIFO.Requery

fila = 1

If CAR_GRIFO.EOF = True Then
   gridl.Row = 1
   gridl.COL = 1
   gridl.Text = "No hay registros"
End If
Do Until CAR_GRIFO.EOF
    Pub_Respuesta = vbYes
    If CAR_GRIFO!CAR_TIPDOC <> SUT_LLAVE!SUT_tipdoc And CAR_GRIFO!CAR_TIPDOC <> "DC" Then Pub_Respuesta = vbNo
    If Pub_Respuesta = vbYes Then
       fila = fila + 1
       gridl.Rows = fila + 1
       grid_canje.Rows = fila + 1
       
       gridl.RowHeight(fila) = 350
       gridl.TextMatrix(fila, 15) = cli_llave!cli_codclie
       
       If Option1.Value = True Then
          gridl.TextMatrix(fila, 0) = Nulo_Valors(CAR_GRIFO!CAR_PLACA)
       ElseIf Option2.Value = True Then
               PUB_KEY = CAR_GRIFO!CAR_CODART
               pu_codcia = LK_CODCIA
               SQ_OPER = 1
               LEER_ART_LLAVE
               If art_LLAVE.EOF Then
                  gridl.TextMatrix(fila, 0) = " ...."
               Else
                  gridl.TextMatrix(fila, 0) = art_LLAVE!ART_NOMBRE
               End If
       Else
          gridl.TextMatrix(fila, 0) = CAR_GRIFO!CAR_FECHA_INGR
          'gridl.TextMatrix(fila, 0) = CAR_GRIFO!CAR_FECHA_sunat
       End If
       
       
       gridl.TextMatrix(fila, 13) = CAR_GRIFO!CAR_TIPDOC
       
       gridl.TextMatrix(fila, 1) = Nulo_Valors(CAR_GRIFO!CAR_FBG)
       If Trim(gridl.TextMatrix(fila, 1)) = "" Then
          If Nulo_Valors(CAR_GRIFO!CAR_TIPDOC) = "LE" Then
             gridl.TextMatrix(fila, 1) = Nulo_Valors(CAR_GRIFO!CAR_TIPDOC) & "-" & Nulo_Valors(CAR_GRIFO!CAR_SITUACION)
          Else
             gridl.TextMatrix(fila, 1) = Nulo_Valors(CAR_GRIFO!CAR_TIPDOC)
          End If
       End If
       
       If Nulo_Valors(CAR_GRIFO!CAR_FBG) = "N" Then gridl.TextMatrix(fila, 1) = "N.Cred." & CAR_GRIFO!CAR_numfac
       If Nulo_Valors(CAR_GRIFO!CAR_FBG) = "D" Then gridl.TextMatrix(fila, 1) = "N.Deb ." & CAR_GRIFO!CAR_numfac
       If Nulo_Valors(CAR_GRIFO!CAR_TIPDOC) = "CH" Then gridl.TextMatrix(fila, 1) = "Chq." & CAR_GRIFO!CAR_NUM_CHEQUE
             
       If cli_llave!CLI_CP = "P" Then
          gridl.TextMatrix(fila, 2) = Nulo_Valor0(CAR_GRIFO!CAR_NUMSER_C)
          gridl.TextMatrix(fila, 3) = Nulo_Valor0(CAR_GRIFO!CAR_NUMFAC_C)
       Else
          gridl.TextMatrix(fila, 2) = Nulo_Valor0(CAR_GRIFO!CAR_numSER)
          gridl.TextMatrix(fila, 3) = Nulo_Valor0(CAR_GRIFO!CAR_numfac)
          If CAR_GRIFO!CAR_TIPMOV = 0 Then
             gridl.TextMatrix(fila, 3) = Nulo_Valor0(CAR_GRIFO!CAR_NUM_CHEQUE)
          End If
       End If
       
       If CAR_GRIFO!CAR_MONEDA = "S" Then
          gridl.TextMatrix(fila, 4) = "S/." & CAR_GRIFO!CAR_IMPORTE
       Else
          gridl.TextMatrix(fila, 4) = "  $" & CAR_GRIFO!CAR_IMPORTE
       End If
       gridl.TextMatrix(fila, 5) = ""
       gridl.TextMatrix(fila, 5) = CAR_GRIFO!CAR_IMP_INI
       
       If CAR_GRIFO!CAR_IMPORTE = 0 Then
          gridl.TextMatrix(fila, 4) = CAR_GRIFO!CAR_IMP_INI
          gridl.TextMatrix(fila, 5) = CAR_GRIFO!CAR_IMPORTE
       End If
         
       gridl.TextMatrix(fila, 6) = CAR_GRIFO!car_fecha_vcto
       gridl.TextMatrix(fila, 7) = Nulo_Valor0(CAR_GRIFO!CAR_COBRADOR) ' Nulo_Valors(CAR_GRIFO!car_concepto)
       gridl.TextMatrix(fila, 9) = CAR_GRIFO!CAR_IMP_INI
       gridl.TextMatrix(fila, 8) = CAR_GRIFO!CAR_FECHA_INGR
       gridl.TextMatrix(fila, 10) = CAR_GRIFO!CAR_cp
       gridl.TextMatrix(fila, 11) = CAR_GRIFO!car_SERDOC
       gridl.TextMatrix(fila, 12) = CAR_GRIFO!car_NUMDOC
       gridl.TextMatrix(fila, 16) = CAR_GRIFO!car_numguia
       gridl.TextMatrix(fila, 17) = Nulo_Valor0(CAR_GRIFO!CAR_SITUACION)
       gridl.TextMatrix(fila, 24) = Nulo_Valors(CAR_GRIFO!CAR_MONEDA)
       gridl.TextMatrix(fila, 25) = CAR_GRIFO!CAR_IMPORTE
       gridl.TextMatrix(fila, 26) = CAR_GRIFO!car_fecha_vcto
       
       TOTAL_DEUDA = CAR_GRIFO!CAR_IMPORTE + TOTAL_DEUDA
    End If

    WS_TIPDOC = CAR_GRIFO!CAR_TIPDOC
    CAR_GRIFO.MoveNext
Loop
     If LK_CODTRA <> 1455 Then
        If gridl.Enabled And gridl.Visible Then
         gridl.SetFocus
         textovarl.Width = 30
         textovarl.Height = 30
       End If
     End If
     
   If gridl.Visible = False Then gridl.Visible = True
   If textovarl.Visible = False Then textovarl.Visible = True

   If LK_CODTRA = 2770 Then
      grid_canje.ColWidth(0) = 1500
   Else
      grid_canje.ColWidth(0) = 0
   End If
grid_canje.Rows = gridl.Rows
fila = 2
Do Until fila = gridl.Rows
   grid_canje.RowHeight(fila) = 350
   fila = fila + 1
Loop
If LK_CODTRA = 2410 Or LK_CODTRA = 2412 Or LK_CODTRA = 2418 Then grid_canje.Visible = False
calcula_totales2
SendKeys "{UP}"

End Sub
Public Sub procesa_celdas()
textovar.MaxLength = 0
If grid_fac.Text = "" Then flag_salto = 0
'If flag_salto = 1 Then Exit Sub


If grid_fac.COL = 10 And (PUB_TIPMOV = 93) Then Exit Sub
If grid_fac.COL = 8 And PUB_TIPMOV = 10 Then Exit Sub


If Not SUT_LLAVE.EOF Then
   If grid_fac.COL = 6 And Val(Nulo_Valor0(SUT_LLAVE!SUT_PRECIO)) = 1 Then Exit Sub
End If

calcula_totales
If grid_fac.COL = 9 Then
    textovar.Visible = False
    i_mortal.Left = Frame4.Left + grid_fac.CellLeft
    i_mortal.Width = grid_fac.CellWidth
    i_mortal.Top = Frame4.Top + grid_fac.Top + grid_fac.CellTop
    i_mortal.Visible = True
    textovar.Visible = False
    i_unidades.Visible = False
    i_mortal.SetFocus
ElseIf grid_fac.COL = 12 Then
    If LK_CODTRA = 2414 Then Exit Sub
    textovar.Visible = False
    i_is.Left = Frame4.Left + grid_fac.CellLeft
    i_is.Width = grid_fac.CellWidth
    i_is.Top = Frame4.Top + grid_fac.Top + grid_fac.CellTop
    i_is.Visible = True
    textovar.Visible = False
    i_mortal.Visible = False
    i_unidades.Visible = False
    i_is.SetFocus
ElseIf grid_fac.COL = 5 And grid_fac.TextMatrix(grid_fac.Row, 16) <> "" Then    ' AND Val(Nulo_Valor0(SUT_LLAVE!SUT_precio)) = 3 Then
    i_mortal.Visible = False
    i_precios.Visible = False
    i_is.Visible = False
    textovar.Visible = False
    i_unidades.Left = Frame4.Left + grid_fac.CellLeft
    i_unidades.Width = grid_fac.CellWidth + 200
    i_unidades.Top = Frame4.Top + grid_fac.Top + grid_fac.CellTop
    i_unidades.Visible = True
    i_unidades.SetFocus
ElseIf grid_fac.COL = 6 And Val(Nulo_Valor0(SUT_LLAVE!SUT_PRECIO)) = 3 And grid_fac.TextMatrix(grid_fac.Row, 16) <> "" Then
    i_mortal.Visible = False
    i_is.Visible = False
    i_unidades.Visible = False
    textovar.Visible = False
    i_precios.Left = Frame4.Left + grid_fac.CellLeft
    i_precios.Width = grid_fac.CellWidth + 800
    i_precios.Top = Frame4.Top + grid_fac.Top + grid_fac.CellTop
    i_precios.Visible = True
    i_precios.SetFocus
ElseIf grid_fac.COL = 1 Or grid_fac.COL = 2 Or grid_fac.COL = 3 Or grid_fac.COL = 4 Or grid_fac.COL = 8 Or grid_fac.COL = 10 Or grid_fac.COL = 15 Or (grid_fac.COL = 38 And LK_CODTRA = 2103) Or (grid_fac.COL = 6 And LK_CODTRA = 1402) Or (grid_fac.COL = 6 And LK_CODTRA = 2211) Then
    textovar.Left = Frame4.Left + grid_fac.CellLeft
    textovar.Width = grid_fac.CellWidth
    textovar.Height = grid_fac.CellHeight
    textovar.Top = Frame4.Top + grid_fac.Top + grid_fac.CellTop
    textovar.Text = grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL)
    textovar_bak = textovar.Text
    textovar.DataField = grid_fac.COL
    textovar.Tag = grid_fac.Row
    textovar.Visible = True
    i_mortal.Visible = False
    i_is.Visible = False
    i_precios.Visible = False
    i_unidades.Visible = False
    Azul textovar, textovar
ElseIf grid_fac.COL = 7 Or (SUT_LLAVE!SUT_PRECIO = "2" And PUB_TIPMOV = 10) Or (SUT_LLAVE!SUT_PRECIO = "2" And PUB_TIPMOV = 101) Then
     If Val(i_codcli.Text) <> 0 And Not cli_llave.EOF And grid_fac.COL = 7 And PUB_TIPMOV = 10 And Val(grid_fac.TextMatrix(grid_fac.Row, 1)) <> 0 And grid_fac.Row = grid_fac.Rows - 1 Then
     If cli_llave!cli_codclie = Val(i_codcli.Text) Then
        If Left(SUT_LLAVE!sut_descto, 1) = "P" And Nulo_Valor0(cli_llave!CLI_PORDESCTO) <> 0 And LK_CODTRA = 2101 Then grid_fac.TextMatrix(grid_fac.Row, 10) = cli_llave!CLI_PORDESCTO
     End If
     End If
    If grid_fac.CellWidth <= 15 Then Exit Sub
    
    textovar.Left = Frame4.Left + grid_fac.CellLeft
    textovar.Width = grid_fac.CellWidth
    textovar.Height = grid_fac.CellHeight
    textovar.Top = Frame4.Top + grid_fac.Top + grid_fac.CellTop
    textovar.Text = grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL)
    textovar_bak = textovar.Text
    textovar.DataField = grid_fac.COL
    textovar.Tag = grid_fac.Row
    textovar.Visible = True
    i_mortal.Visible = False
    i_is.Visible = False
    i_precios.Visible = False
    i_unidades.Visible = False
    Azul textovar, textovar
End If

End Sub
Public Sub PROCESA_CELDAS2()
If GridT.Row < 2 Then Exit Sub
textovar2.MaxLength = 0
'If GridT.Text <> "" Then flag_salto = 0
'If GridT.TextMatrix(GridT.Row, 0) <> "" And GridT.COL <> 1 Then flag_salto = 0
'If flag_salto = 1 Then Exit Sub

If Not SUT_LLAVE.EOF Then
   If GridT.COL = 6 And Val(Nulo_Valor0(SUT_LLAVE!SUT_PRECIO)) = 1 Then Exit Sub
End If

calcula_totales22
If GridT.COL = 1 Or GridT.COL = 2 Or GridT.COL = 3 Or GridT.COL = 7 Then
    textovar2.Left = Frame5.Left + GridT.CellLeft
    textovar2.Width = GridT.CellWidth
    textovar2.Height = GridT.CellHeight
    textovar2.Top = Frame5.Top + GridT.Top + GridT.CellTop
    textovar2.Text = GridT.TextMatrix(GridT.Row, GridT.COL)
    textovar2.DataField = GridT.COL
    textovar2.Tag = GridT.Row
    textovar2.Visible = True
    Azul textovar2, textovar2
End If
End Sub
Public Sub PROCESA_CELDAS3()
If GridD.Row < 2 Then Exit Sub
textovar3.MaxLength = 0
If GridD.Text = "" Then flag_salto = 0
'If flag_salto = 1 Then Exit Sub

'calcula_totales22
If GridD.COL = 3 Or GridD.COL = 5 Or GridD.COL = 6 Then
    textovar3.Left = Frame5.Left + GridD.CellLeft
    textovar3.Width = GridD.CellWidth
    textovar3.Height = GridD.CellHeight
    textovar3.Top = Frame5.Top + GridD.Top + GridD.CellTop
    textovar3.Text = GridD.TextMatrix(GridD.Row, GridD.COL)
    textovar3.DataField = GridD.COL
    textovar3.Tag = GridD.Row
    textovar3.Visible = True
    Azul textovar3, textovar3
End If
End Sub


Private Sub procesa_celdas_l()

End Sub

Private Sub pasa_cabeza_canje()

   grid_canje.TextMatrix(0, 1) = "Orden"
   grid_canje.TextMatrix(0, 2) = "TipoDoc"
   grid_canje.TextMatrix(0, 3) = "Importe"
   grid_canje.TextMatrix(0, 4) = "Fecha "
   grid_canje.TextMatrix(0, 5) = "Banco"
   grid_canje.TextMatrix(0, 6) = "N. Documento"
   grid_canje.TextMatrix(0, 7) = "Observaciones"

End Sub

Private Sub pasa_cabeza()
   grid_fac.TextMatrix(0, 0) = "Descripción"
   grid_fac.TextMatrix(0, 1) = "Codigo"
   grid_fac.TextMatrix(0, 2) = "Bultos"
   grid_fac.TextMatrix(0, 4) = "Cantidad"
   grid_fac.TextMatrix(0, 5) = "Unidad"
   grid_fac.TextMatrix(0, 6) = "Precio"
   grid_fac.TextMatrix(0, 7) = "Subtotal"
   grid_fac.TextMatrix(0, 8) = "Flete"
   grid_fac.TextMatrix(0, 9) = "Causas"
   grid_fac.TextMatrix(0, 10) = "Descto"
   grid_fac.TextMatrix(0, 11) = "Costo"
   grid_fac.TextMatrix(0, 12) = "Ing/Salida"
   grid_fac.TextMatrix(0, 13) = "S t o c k"
   grid_fac.TextMatrix(0, 15) = "Placas"
   grid_fac.TextMatrix(0, 38) = "Pedido"


End Sub
Private Sub pasa_cabeza_grid_trans()
   grid_trans.TextMatrix(0, 0) = "Fecha"
'   grid_trans.TextMatrix(0, 1) = "N.Serie "
   grid_trans.TextMatrix(0, 1) = "N.Pedido"
   grid_trans.TextMatrix(0, 2) = "Articulo"
   grid_trans.TextMatrix(0, 3) = "Bultos"
   grid_trans.TextMatrix(0, 4) = "Kilos "
   grid_trans.TextMatrix(0, 11) = "Guia"
   grid_trans.TextMatrix(0, 12) = "Nro/Ped"
   
End Sub

Private Sub pasa_cabeza_L()
   gridl.TextMatrix(0, 0) = "      Nombre "
   gridl.TextMatrix(0, 1) = "FBG"
   gridl.TextMatrix(0, 2) = "Serie"
   gridl.TextMatrix(0, 3) = "N.Factura"
   gridl.TextMatrix(0, 4) = "Saldo"
   gridl.TextMatrix(0, 5) = "Efectivo"
   gridl.TextMatrix(0, 6) = "Nueva Fecha"
   gridl.TextMatrix(0, 7) = "Cobrador"
   gridl.TextMatrix(0, 8) = "F.Ingreso"
   gridl.TextMatrix(0, 9) = "Saldo Inicial"
   gridl.TextMatrix(0, 10) = "C/P"
   gridl.TextMatrix(0, 11) = "Serdoc"
   gridl.TextMatrix(0, 12) = "NumDoc"
   gridl.TextMatrix(0, 13) = "Tipdoc"
   gridl.TextMatrix(0, 14) = "Codcia"
   gridl.TextMatrix(0, 15) = "codclie"
   gridl.TextMatrix(0, 16) = "N. Guia"
   gridl.TextMatrix(0, 18) = "Vendedor"
   gridl.TextMatrix(0, 19) = "Imp.Cheque"
   gridl.TextMatrix(0, 20) = "Num.Cheque"
   gridl.TextMatrix(0, 21) = "Banco"
   gridl.TextMatrix(0, 22) = "Fecha de cobro"
   gridl.TextMatrix(0, 23) = "Tipo de Documento"
   gridl.TextMatrix(0, 24) = "Moneda"


End Sub
Private Sub pasa_cabeza_L2()
   gridl.TextMatrix(0, 0) = "      Nombre "
   gridl.TextMatrix(0, 1) = "FBG"
   gridl.TextMatrix(0, 2) = "Serie"
   gridl.TextMatrix(0, 3) = "N.Factura"
   gridl.TextMatrix(0, 4) = "N. Guia "
   gridl.TextMatrix(0, 5) = "  Saldo"
   gridl.TextMatrix(0, 6) = "Fecha Vcto."
   gridl.TextMatrix(0, 7) = "Concepto    "
   gridl.TextMatrix(0, 8) = "F.Ingreso"
   gridl.TextMatrix(0, 9) = "Saldo Inicial"
   gridl.TextMatrix(0, 10) = "C/P"
   gridl.TextMatrix(0, 11) = "Serdoc"
   gridl.TextMatrix(0, 12) = "NumDoc"
   gridl.TextMatrix(0, 13) = "Tipdoc"
   gridl.TextMatrix(0, 14) = "Codcia"
   gridl.TextMatrix(0, 15) = "codclie"

End Sub

Private Sub BolFac_Click()

If LK_FLAG_FACTURACION = "V" And PUB_TIPMOV = 10 And i_codven.Text = "" Then Exit Sub


If (PUB_TIPMOV <> 0 And pub_signo_arm <> 0) Or LK_CODTRA = 1402 Or LK_CODTRA = 2212 Or LK_CODTRA = 2406 Then
   If LK_FLAG_FACTURACION = "V" And PUB_TIPMOV = 10 Then
      If ven_llave.EOF Then
          MsgBox "Falta indicar la Isla..."
          Exit Sub
      Else
          If ven_llave!vem_codven <> Val(i_codven.Text) Then
             MsgBox "Falta indicar la Isla..."
          Exit Sub
          End If
      End If
   End If
   llena_numfac_grifo
End If
' AGREGE

i_cambio.Visible = True
i_cambio.Value = 0
i_numfac.Locked = True
i_numser.Locked = True
i_numfac.BackColor = QBColor(7)
i_numser.BackColor = QBColor(7)
i_ser_lote.BackColor = QBColor(7)
i_num_lote.BackColor = QBColor(7)
If LK_CODTRA = 2401 Then i_fecha_compra.Visible = False


End Sub


Private Sub CAMBIA_MONEDA()
Exit Sub
Dim WRES
Dim fila As Integer
Dim ws_curr As Currency
fila = 2
  If Val(WS_TIPO_CAMBIO) = 0 Then WS_TIPO_CAMBIO = LK_TIPO_CAMBIO

  WRES = InputBox("Ingrese Tipo de cambio :", , WS_TIPO_CAMBIO)
  If WRES = "" Then
  Exit Sub
  End If
  If Val(WRES) <= 0 Then
    Exit Sub
  End If
  WS_TIPO_CAMBIO = WRES
  If (Abs(WS_TIPO_CAMBIO - LK_TIPO_CAMBIO) * 100 / LK_TIPO_CAMBIO) > 5 Then
     WS_TIPO_CAMBIO = 0
     MsgBox "Cambiar Tipo De Cambio... Mucha diferencia"
  End If


If pub_signo_car = 0 Then
   If Mid(Tipo_Cambio.Caption, 2, 1) = "S" Then
      Tipo_Cambio.Caption = "&Dollares " & WS_TIPO_CAMBIO
   Else
      Tipo_Cambio.Caption = "&Soles " & WS_TIPO_CAMBIO
   End If
Exit Sub
End If

Do Until fila = 999
   If Mid(Tipo_Cambio.Caption, 2, 1) = "S" Then
      gridl.TextMatrix(fila, 6) = Val(grid_fac.TextMatrix(fila, 6)) / WS_TIPO_CAMBIO
      grid_fac.TextMatrix(fila, 6) = Format(Val(grid_fac.TextMatrix(fila, 6)), "#####.0000")
   Else
   ws_curr = Val(grid_fac.TextMatrix(fila, 6)) * WS_TIPO_CAMBIO
   grid_fac.TextMatrix(fila, 6) = ws_curr
   grid_fac.TextMatrix(fila, 6) = Format(Val(grid_fac.TextMatrix(fila, 6)), "#######.0000")
   End If
   fila = fila + 1
   If fila = grid_fac.Rows Then
      fila = 999
   End If
Loop
If Mid(Tipo_Cambio.Caption, 2, 1) = "S" Then
   i_descto.Text = redondea(Val(i_descto.Text) / WS_TIPO_CAMBIO)
   i_gastos.Text = redondea(Val(i_gastos.Text) / WS_TIPO_CAMBIO)
   Tipo_Cambio.Caption = "&Dollares " & WS_TIPO_CAMBIO
Else
   i_descto.Text = redondea(Val(i_descto.Text) * WS_TIPO_CAMBIO)
   i_gastos.Text = redondea(Val(i_gastos.Text) * WS_TIPO_CAMBIO)
   Tipo_Cambio.Caption = "&Soles " & WS_TIPO_CAMBIO
End If

calcula_totales
grid_fac.SetFocus

End Sub
Private Sub calcula_totales22()
Dim f As Integer
If SUT_LLAVE.EOF Then Exit Sub
Dim subtotal As Currency
Dim ROLLOS As Currency

GridT.TextMatrix(0, 0) = "Producto"
GridT.TextMatrix(0, 1) = "Codigo"
GridT.TextMatrix(0, 2) = "Bultos"
GridT.TextMatrix(0, 3) = "Klgr."
 If PUB_TIPMOV = 177 Then
   GridT.ColWidth(6) = 1
   GridT.ColWidth(7) = 1
 Else
   GridT.ColWidth(6) = 500
   GridT.ColWidth(7) = 800
 End If

f = 2
Do Until f = GridT.Rows - 1
   ROLLOS = Val(GridT.TextMatrix(f, 2)) + ROLLOS
   subtotal = Val(GridT.TextMatrix(f, 3)) + subtotal
   f = f + 1
Loop
subtotal = redondea(subtotal)
GridT.TextMatrix(1, 2) = ROLLOS
GridT.TextMatrix(1, 3) = subtotal



End Sub
Private Sub calcula_totales24()
Dim f As Integer
If SUT_LLAVE.EOF Then Exit Sub
Dim subtotal As Currency

f = 2
Do Until f = GridD.Rows
   subtotal = Val(GridD.TextMatrix(f, 6)) + subtotal
   f = f + 1
Loop
subtotal = redondea(subtotal)
GridD.TextMatrix(1, 6) = subtotal
GridD.TextMatrix(1, 2) = subtotal

End Sub
Private Sub calcula_totales25()
Dim f As Integer
If SUT_LLAVE.EOF Then Exit Sub
Dim ww_cantidad As Currency
Dim ww_req As Currency

f = 2
Do Until f = gridl.Rows - 1
   grid_fac.Rows = grid_fac.Rows + 1
   If Val(gridl.TextMatrix(f, 4)) > 0 Then
      gridl.TextMatrix(f, 5) = gridl.TextMatrix(f, 2) - gridl.TextMatrix(f, 4)
   Else
      gridl.TextMatrix(f, 5) = gridl.TextMatrix(f, 2)
   End If
   ww_req = Val(gridl.TextMatrix(f, 5)) + ww_req
   ww_cantidad = Val(gridl.TextMatrix(f, 2)) + ww_cantidad
   grid_fac.TextMatrix(f, 4) = Val(gridl.TextMatrix(f, 2))
   grid_fac.TextMatrix(f, 7) = Format(Val(grid_fac.TextMatrix(f, 4)) * Val(grid_fac.TextMatrix(f, 6)), "0.00")
   f = f + 1
Loop
gridl.TextMatrix(1, 2) = ww_cantidad
gridl.TextMatrix(1, 5) = ww_req


End Sub


Private Sub calcula_totales()
Dim ws_bruto_bak
Dim ww_descto2 As Currency
Dim ws_ex_igv As Currency
Dim WS_MONEDA As String
Dim WS_DESCTO2 As Currency
fila = 2
SUB_FLETE = 0
WS_DESCTO2 = 0
ww_descto2 = 0
SUB_CANT = 0
ws_bruto_bak = WS_BRUTO
WS_BRUTO = 0
SUB_UNIDAD = 0
pub_ojo = ""
If SUT_LLAVE.EOF Then Exit Sub
PUB_DESCTO = 0
PUB_DESCTO = 0
PUB_IMPTO = 0
Dim subtotal As Currency

Do Until fila = 999
   
   If PUB_TIPMOV = 10 Then
       If Trim(SUT_LLAVE!SUT_abreviado) = "CONS" Or LK_CODTRA = 2212 Or LK_CODTRA = 2103 Then
         grid_fac.TextMatrix(fila, 7) = Format(Val(grid_fac.TextMatrix(fila, 4)) * Val(grid_fac.TextMatrix(fila, 6)), "0.0000")
       Else
          If Val(grid_fac.TextMatrix(fila, 6)) > 0 Then grid_fac.TextMatrix(fila, 4) = Format(Val(grid_fac.TextMatrix(fila, 7)) / Val(grid_fac.TextMatrix(fila, 6)), "0.00")
          If Val(grid_fac.TextMatrix(fila, 6)) = 0 And Val(grid_fac.TextMatrix(fila, 35)) <> 0 Then grid_fac.TextMatrix(fila, 4) = Format(Val(grid_fac.TextMatrix(fila, 7)) / Val(grid_fac.TextMatrix(fila, 35)), "0.00")
       End If
      subtotal = Val(grid_fac.TextMatrix(fila, 7))
   End If
      
   If PUB_TIPMOV = 20 Then
         grid_fac.TextMatrix(fila, 7) = Format(Val(grid_fac.TextMatrix(fila, 4)) * Val(grid_fac.TextMatrix(fila, 6)), "0.0000")
         subtotal = Val(grid_fac.TextMatrix(fila, 7))
   End If
      
      
   If grid_fac.TextMatrix(fila, 23) = "A" Then
      If Not cli_llave.EOF And Val(i_codcli.Text) > 0 Then
         ws_ex_igv = redondea(subtotal * (LK_IGV - Val(grid_fac.TextMatrix(fila, 24)))) / 100 + ws_ex_igv
      End If
   End If

   
   If PUB_TIPMOV = 10 Then
      If LK_CODTRA = 2105 Then
         ww_descto2 = Val(grid_fac.TextMatrix(fila, 10))
      Else
         ww_descto2 = (Val(grid_fac.TextMatrix(fila, 35)) - Val(grid_fac.TextMatrix(fila, 6))) * Val(grid_fac.TextMatrix(fila, 4))
      End If
      ww_descto2 = redondea(ww_descto2)
      grid_fac.TextMatrix(fila, 10) = ww_descto2
      PUB_DESCTO = PUB_DESCTO + ww_descto2
   End If
   
   
   
   SUB_FLETE = Val(grid_fac.TextMatrix(fila, 8)) + SUB_FLETE
   
   If LK_CODTRA = 2210 Then
      If Val(grid_fac.TextMatrix(fila, 10)) = 9999 Then
         SUB_CANT = SUB_CANT - Val(grid_fac.TextMatrix(fila, 4))
      Else
         SUB_CANT = Val(grid_fac.TextMatrix(fila, 4)) + SUB_CANT
      End If
   Else
      SUB_CANT = Val(grid_fac.TextMatrix(fila, 4)) + SUB_CANT
   End If
   
      
   SUB_UNIDAD = Val(grid_fac.TextMatrix(fila, 3)) + SUB_UNIDAD
   If LK_CODTRA = 2414 And fila = 3 Then subtotal = subtotal * -1
   WS_BRUTO = WS_BRUTO + subtotal
   
'   WS_DESCTO2 = WS_DESCTO2 + Val(grid_fac.TextMatrix(fila, 4)) * (Val(grid_fac.TextMatrix(fila, 35)) - Val(grid_fac.TextMatrix(fila, 6)))
'   WS_DESCTO2 = redondea(WS_DESCTO2)
   fila = fila + 1
   If fila = grid_fac.Rows Then
      grid_fac.TextMatrix(1, 8) = SUB_FLETE
      grid_fac.TextMatrix(1, 3) = SUB_UNIDAD
      grid_fac.TextMatrix(1, 4) = SUB_CANT
      grid_fac.TextMatrix(1, 7) = WS_BRUTO
      grid_fac.TextMatrix(1, 10) = PUB_DESCTO
      fila = 999
   End If
   
   
Loop

   
PUB_DESCTO = 0
ws_igv = 0
If i_fbg.Visible = True Then
   ws_igv = LK_IGV
   If i_fbg.Text = "G" Or i_fbg.Text = "P" Then ws_igv = 0
End If




PUB_DESCTO = redondea(PUB_DESCTO)

WS_NETO = WS_BRUTO

If PUB_TIPMOV = 20 Then
   PUB_GASTOS = Val(i_gastos.Text)
If Not cli_llave.EOF And Val(i_codcli.Text) > 0 Then
      WS_NETO = WS_BRUTO
      PUB_IMPTO = (WS_BRUTO - PUB_DESCTO + PUB_GASTOS) * (ws_igv / 100)
      PUB_IMPTO = PUB_IMPTO - ws_ex_igv
      If ws_igv = 0 Then
         PUB_IMPTO = 0
      Else
         PUB_IMPTO = redondea(PUB_IMPTO)
      End If
      WS_NETO = WS_BRUTO - PUB_DESCTO + PUB_GASTOS + PUB_IMPTO
End If
End If



If PUB_TIPMOV = 10 Or PUB_TIPMOV = 93 Or PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
   WS_NETO = WS_BRUTO - PUB_DESCTO
   If par_llave!PAR_VENTAS_IGV = "A" Then
      If PUB_DESCTO <> 0 Then
         PUB_DESCTO = redondea((PUB_DESCTO) / (1 + ws_igv / 100))
      End If
      
      WS_BRUTO = (WS_NETO) / (1 + ws_igv / 100)
      WS_BRUTO = WS_BRUTO + PUB_DESCTO
      WS_BRUTO = redondea(WS_BRUTO)
      PUB_IMPTO = WS_NETO - WS_BRUTO + PUB_DESCTO
      If PUB_IMPTO = 0 And ws_igv > 0 Then
         PUB_IMPTO = redondea((WS_BRUTO + PUB_GASTOS) * (ws_igv / 100))
      End If
   Else
      PUB_IMPTO = (WS_BRUTO - PUB_DESCTO + PUB_GASTOS) * (ws_igv / 100)
      If ws_igv = 0 Then
         PUB_IMPTO = 0
      Else
         PUB_IMPTO = redondea(PUB_IMPTO)
      End If
      WS_NETO = WS_BRUTO - PUB_DESCTO + PUB_GASTOS + PUB_IMPTO
   End If
End If
If par_llave!PAR_VENTAS_IGV = "E" Then
     ws_igv = 0
     ' If PUB_DESCTO <> 0 Then
     '    PUB_DESCTO = redondea((PUB_DESCTO) / (1 + WS_IGV / 100))
     ' End If
     ' WS_NETO = WS_NETO + PUB_FLETE
     ' WS_BRUTO = (WS_NETO) / (1 + WS_IGV / 100)
      WS_BRUTO = WS_BRUTO + PUB_DESCTO
      WS_BRUTO = redondea(WS_BRUTO)
      WS_NETO = WS_BRUTO
      PUB_IMPTO = 0
      'If PUB_IMPTO = 0 And WS_IGV > 0 Then
      '   PUB_IMPTO = redondea((WS_BRUTO + PUB_GASTOS) * (WS_IGV / 100))
      'End If
      PUB_IMPTO = PUB_IMPTO + ws_ex_igv
      WS_BRUTO = WS_BRUTO + Abs(ws_ex_igv)
End If


      

   i_impto.Text = PUB_IMPTO
   i_descto.Text = PUB_DESCTO
   i_subtotal.Text = WS_BRUTO
   
   
   
   i_neto.Text = WS_NETO

   

If LK_CODTRA = 1403 And Val(i_importe.Text) > 0 Then
      i_importe_amort.Text = WS_NETO - Val(i_importe.Text)
      If Val(i_importe_amort.Text) <> 0 Then
      If (pub_signo_car > 0 And Val(i_importe_amort.Text) > 0) Or (pub_signo_car < 0 And Val(i_importe_amort.Text) < 0) Then
      Else
         MsgBox "NO procede ...Cambiar opcion Aumento/disminucion..."
         cancela_todo
         Exit Sub
      End If
      End If
End If

If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then i_importe_amort.Text = WS_NETO

If PUB_TIPMOV <> 10 Then Exit Sub

If cli_llave.EOF Then Exit Sub

grid_fac.TextMatrix(1, 0) = ""
'If LK_CODTRA = 2101 Then grid_fac.TextMatrix(1, 10) = WS_DESCTO2


If SUT_LLAVE!SUT_SIGNO_CAR <> 1 Then Exit Sub
If Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 1 Then Exit Sub


If LK_CODTRA <> 1111 Then
If LK_FLAG_EXED = "A" Then
If (pub_deuda + WS_NETO) > Nulo_Valor0(cli_llave!cli_limcre) Then
         If ws_bruto_bak <> WS_BRUTO Then MsgBox "Credito Excedido... "
End If
End If
End If

grid_fac.TextMatrix(1, 0) = "Disponible:   " & Format((Nulo_Valor0(cli_llave!cli_limcre) - pub_deuda - WS_NETO), "##,###,###.00")

End Sub

Private Sub calcula_totales2()
Dim FILAX
Dim FF As Integer
FILAX = 2
Dim subtotal As Currency
Dim subtotal_che As Currency
Dim saldocar_s As Currency
Dim saldocar_d As Currency

Dim WS_MONEDA As String * 1
FF = 0
subtotal = 0
subtotal_che = 0
saldocar_s = 0
saldocar_d = 0

If i_ds.Visible = True Then
   If i_ds.Text = "S" Or i_ds.Text = "D" Then
   Else
      MsgBox "Seleccione Moneda..."
      Exit Sub
   End If
End If
If i_ds.Visible = True And pub_signo_ccm <> 0 Then
   If i_ds.Text <> WS_MONEDA_CCM Then
      MsgBox "Moneda errada ..."
      Exit Sub
   End If
End If


Do Until FILAX = 999
         If LK_CODTRA = 1455 And Val(gridl.TextMatrix(FILAX, 5)) <> 0 Then
            If Val(gridl.TextMatrix(FILAX, 5)) <> Val(gridl.TextMatrix(FILAX, 25)) Then
               MsgBox "Importe debe ser Cancelación"
               gridl.TextMatrix(FILAX, 5) = 0
            End If
         End If
         If gridl.TextMatrix(FILAX, 24) = "S" Then
            saldocar_s = saldocar_s + Val(gridl.TextMatrix(FILAX, 25))
         Else
            saldocar_d = saldocar_d + Val(gridl.TextMatrix(FILAX, 25))
         End If
         
         If Val(gridl.TextMatrix(FILAX, 25)) < 0 Then
            subtotal = subtotal - Val(gridl.TextMatrix(FILAX, 5))
         Else
            subtotal = subtotal + Val(gridl.TextMatrix(FILAX, 5))
         End If
         If Val(gridl.TextMatrix(FILAX, 5)) <> 0 Then
            If FF = 0 Then
               WS_MONEDA = gridl.TextMatrix(FILAX, 24)
               FF = 1
            Else
               If WS_MONEDA <> gridl.TextMatrix(FILAX, 24) And LK_CODTRA <> 1455 Then
                  MsgBox "HAY DIFERENTES MONEDAS..."
                  cancelar_Click
                  Exit Sub
               End If
            End If
         End If
         
         If WS_MONEDA = "S" Then
            WS_MONEDA_CLI = "S"
         Else
            WS_MONEDA_CLI = "D"
         End If
         
         subtotal_che = subtotal_che + Val(gridl.TextMatrix(FILAX, 19))
         
   FILAX = FILAX + 1
   If FILAX = gridl.Rows Then
      FILAX = 999
   End If
Loop
   gridl.TextMatrix(1, 4) = "S/." & saldocar_s
   gridl.TextMatrix(1, 3) = "$" & saldocar_d
   
   gridl.TextMatrix(1, 5) = subtotal
   gridl.TextMatrix(1, 19) = subtotal_che
   
   i_importe.Text = subtotal + subtotal_che
   
   If Tipo_Cambio.Visible = False Then Exit Sub
   
   If pub_signo_ccm = 0 Then Exit Sub
   
   
   If Trim(WS_MONEDA_CCM) = Trim(WS_MONEDA_CLI) Then Exit Sub
   
   
   
   
End Sub
Private Sub CALCULA_TOTALES_CANJE()
Dim FILAX
FILAX = 2
Dim subtotal As Currency

Do Until FILAX = 999
   subtotal = subtotal + Val(grid_canje.TextMatrix(FILAX, 3))
   grid_canje.TextMatrix(FILAX, 1) = FILAX - 1 & ".-"
   FILAX = FILAX + 1
   If FILAX = grid_canje.Rows Then
      FILAX = 999
   End If
Loop
grid_canje.TextMatrix(1, 3) = subtotal
i_importe = subtotal + Val(gridl.TextMatrix(1, 5))
End Sub


Private Sub TOMA_DATOS()
'INICIALIZA_INPUT
Dim pos2 As Integer
Dim CARAC As String


pub_numplan = Val(i_numplan.Text)
PUB_NUM_OPER = Val(i_num_oper.Text)

PUB_CODCLIE = Val(i_codcli.Text)
PUB_CODVEN = Val(i_codven.Text)

PUB_IMPORTE = Val(i_importe.Text)

PUB_NETO = Val(i_neto.Text)
PUB_SUBTOTAL = Val(i_subtotal.Text)

PUB_IMPTO = Val(i_impto.Text)
PUB_FLETE = Val(i_flete.Text)

PUB_DESCTO = Val(i_descto.Text)
PUB_GASTOS = Val(i_gastos.Text)


PUB_SERDOC = Val(i_serdoc.Text)

PUB_NUMSER = Val(i_numser.Text)


PUB_NUMFAC = Val(i_numfac.Text)

PUB_NUMSER_C = Val(i_numser_c.Text)
PUB_NUMFAC_C = Val(i_numfac_c.Text)


'PUB_LIMCRE_ACT = Val(i_limcre.Text)
'PUB_LIMCRE_ANT = Val(i_limcre_ant.Text)
pub_dias = Val(i_dias.Text)
PUB_CONCEPTO = i_concepto.Text
If i_fbg.Visible = False Then
   PUB_FBG = ""
Else
   PUB_FBG = i_fbg.Text
End If
If PUB_TIPMOV = 97 Then PUB_FBG = "N"
If PUB_TIPMOV = 98 Then PUB_FBG = "D"
If PUB_TIPMOV = 20 Then PUB_FBG = ""

PUB_CHESEC = 0
PUB_NUMGUIA = Val(i_numguia.Text)
PUB_IMPORTE_AMORT = Val(i_importe_amort.Text)
PUB_DS = i_ds.Text


PUB_NUMDOC = Val(i_numdoc.Text)
If IsDate(i_fecha_vcto.Text) Then
    PUB_FECHA_VCTO = CDate(i_fecha_vcto.Text)
Else
    PUB_FECHA_VCTO = 0
End If
If i_def.ListCount = 0 Then
   PUB_SECUENCIA = 0
Else
   pos2 = InStr(1, i_def.Text, ".", 1)
   CARAC = Mid(i_def.Text, 1, pos2)
   PUB_SECUENCIA = Val(CARAC)
End If

End Sub
Private Sub Boton_Cons_Click()
Dim xf As Integer
Dim wv
Dim wvar  As Currency
If LK_CODTRA = 2103 And PUB_SECUENCIA = 1 Then
  GoTo PROGRA
End If
If LK_CODTRA = 2406 Or LK_CODTRA = 2408 Then
  GoTo ENVIODEV
End If


grid_fac.Clear
pasa_cabeza
If i_codcli.Visible = False Then i_codcli.Text = ""

  PSCONS_LLAVE.rdoParameters(0) = LK_CODCIA
  PSCONS_LLAVE.rdoParameters(1) = Val(i_codcli.Text)
  CONS_llave.Requery
  If CONS_llave.EOF Then
    MsgBox "Cliente No Tiene Guías a Consignación.O Canje", 48, Pub_Titulo
    If i_codcli.Visible Then Azul i_codcli, i_codcli
    Exit Sub
  End If

   fila = 2
   Frame4.Visible = True
   Do Until CONS_llave.EOF
        grid_fac.Rows = fila + 2
        
        If CONS_llave!far_estado = "E" Then GoTo OTRA_FILA
        
        If Val(i_codcli.Text) = 0 Then
           If CONS_llave!far_SIGNO_aRM <> pub_signo_arm * -1 Then GoTo OTRA_FILA
        End If
        
        PUB_KEY = CONS_llave!far_codart
        pu_codcia = LK_CODCIA
        SQ_OPER = 1
        LEER_ART_LLAVE
        If art_LLAVE.EOF And PUB_KEY <> 0 Then
           MsgBox "Error Grave en arti..."
           Exit Sub
        End If
        PUB_CODART = CONS_llave!far_codart
        pu_codcia = LK_CODCIA
        SQ_OPER = 1
        LEER_ARM_LLAVE
        If arm_llave.EOF Then
           MsgBox "Error Grave en arti..."
           Exit Sub
        End If
      
   If PUB_KEY <> 0 Then grid_fac.TextMatrix(fila, 0) = art_LLAVE!ART_NOMBRE
   grid_fac.RowHeight(fila) = 285
   grid_fac.TextMatrix(fila, 1) = art_LLAVE!ART_ALTERNO
   grid_fac.TextMatrix(fila, 16) = CONS_llave!far_codart
   grid_fac.TextMatrix(fila, 33) = CONS_llave!far_codart
   grid_fac.TextMatrix(fila, 2) = CONS_llave!far_JABAS
   grid_fac.TextMatrix(fila, 3) = CONS_llave!far_UNIDADES
   grid_fac.TextMatrix(fila, 11) = CONS_llave!FAR_COSPRO
   If Nulo_Valor0(CONS_llave!FAR_equiv) > 0 Then
   grid_fac.TextMatrix(fila, 4) = (CONS_llave!FAR_CANTIDAD - CONS_llave!FAR_CONSIG) / Nulo_Valor0(CONS_llave!FAR_equiv)
   grid_fac.TextMatrix(fila, 22) = (CONS_llave!FAR_CANTIDAD - CONS_llave!FAR_CONSIG) / Nulo_Valor0(CONS_llave!FAR_equiv)
   Else
   grid_fac.TextMatrix(fila, 4) = CONS_llave!FAR_CANTIDAD - CONS_llave!FAR_CONSIG
   grid_fac.TextMatrix(fila, 22) = CONS_llave!FAR_CANTIDAD - CONS_llave!FAR_CONSIG
   End If
   
   grid_fac.TextMatrix(fila, 37) = grid_fac.TextMatrix(fila, 22)
   grid_fac.TextMatrix(fila, 6) = CONS_llave!FAR_PRECIO
   grid_fac.TextMatrix(fila, 7) = redondea(CONS_llave!FAR_PRECIO * Val(grid_fac.TextMatrix(fila, 4)))
   
   grid_fac.TextMatrix(fila, 36) = CONS_llave!FAR_PRECIO
   
   grid_fac.TextMatrix(fila, 14) = Nulo_Valor0(CONS_llave!FAR_equiv)
   grid_fac.TextMatrix(fila, 5) = Nulo_Valors(CONS_llave!FAR_descri)
   grid_fac.TextMatrix(fila, 8) = Nulo_Valor0(CONS_llave!FAR_FLETE)
   grid_fac.TextMatrix(fila, 10) = CONS_llave!FAR_DESCTO
   grid_fac.TextMatrix(fila, 11) = CONS_llave!FAR_COSPRO
   grid_fac.TextMatrix(fila, 12) = CONS_llave!far_SIGNO_aRM
   grid_fac.TextMatrix(fila, 13) = arm_llave!ARM_STOCK
   grid_fac.TextMatrix(fila, 38) = arm_llave!ARM_Saldo_n
   grid_fac.TextMatrix(fila, 39) = arm_llave!ARM_saldo_s
   
   grid_fac.TextMatrix(fila, 35) = CONS_llave!FAR_PRECIO
   grid_fac.TextMatrix(fila, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
   grid_fac.TextMatrix(fila, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
   grid_fac.TextMatrix(fila, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
   grid_fac.TextMatrix(fila, 26) = CONS_llave!FAR_TIPMOV
   grid_fac.TextMatrix(fila, 27) = CONS_llave!FAR_CODCIA
   grid_fac.TextMatrix(fila, 28) = CONS_llave!far_numser
   grid_fac.TextMatrix(fila, 29) = CONS_llave!FAR_FBG
   grid_fac.TextMatrix(fila, 30) = CONS_llave!FAR_NUMFAC
   grid_fac.TextMatrix(fila, 31) = CONS_llave!FAR_NUMSEC
   
   fila = fila + 1
OTRA_FILA:
   CONS_llave.MoveNext
Loop
calcula_totales
grid_fac.SetFocus

Exit Sub
PROGRA:
wv = InputBox("Nro.de Pedidos mayores a: ", "Filtro... ", "0")
wvar = Val(wv)

grid_liq.Cols = 4
grid_liq.Rows = 1
grid_liq.Width = 4000
grid_liq.Height = 3200
grid_liq.Left = 8500
grid_liq.Top = 2300
grid_liq.ColWidth(0) = 2000
grid_liq.ColWidth(1) = 400
grid_liq.ColWidth(2) = 800
grid_liq.ColWidth(3) = 0

grid_liq.TextMatrix(0, 0) = "Cliente"
grid_liq.TextMatrix(0, 1) = "P/Serie"
grid_liq.TextMatrix(0, 2) = "P/Nro."

grid_liq.ColWidth(1) = 400
grid_liq.ColWidth(2) = 800
grid_liq.ColWidth(3) = 0

Exit Sub

ENVIODEV:

grid_liq.Cols = 6
grid_liq.Rows = 1
grid_liq.Width = 4000
grid_liq.Height = 3200
grid_liq.Left = 8500
grid_liq.Top = 2300
grid_liq.ColWidth(0) = 2000
grid_liq.ColWidth(1) = 400
grid_liq.ColWidth(2) = 800
grid_liq.ColWidth(3) = 0
grid_liq.ColWidth(4) = 0
grid_liq.ColWidth(5) = 0

grid_liq.TextMatrix(0, 0) = "MAq."
grid_liq.TextMatrix(0, 1) = "Tejido"
grid_liq.TextMatrix(0, 2) = "P/Nro."

grid_liq.ColWidth(0) = 400
grid_liq.ColWidth(1) = 2500
grid_liq.ColWidth(2) = 0
grid_liq.ColWidth(3) = 0


grid_liq.Visible = True
grid_liq.Rows = 1
''PED_CODCIA = ? AND PED_TIPMOV = ? AND PED_NUMSER = ? AND PED_NUMFAC = ? AND PED_CODCLIE = ? and PED_CODVEN = ?"
PSDAT_LLAVE(0) = LK_CODCIA
PSDAT_LLAVE(1) = PUB_TIPMOV '180
PSDAT_LLAVE(2) = Val(i_numser_c.Text)
PSDAT_LLAVE(3) = Val(i_numfac_c.Text)
PSDAT_LLAVE(4) = Val(i_codcli.Text)
PSDAT_LLAVE(5) = Val(i_codven.Text)
Set pdat_llave = PSDAT_LLAVE.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)
pdat_llave.Requery
If pdat_llave.EOF Then
  Frame4.Visible = False
  MsgBox "No Existe Documento.....!!!", 48, Pub_Titulo
  If i_numser_c.Visible Then i_numser_c.SetFocus
  Exit Sub
End If
grid_liq.Rows = 1
xf = 0
Do Until pdat_llave.EOF
  xf = xf + 1
  grid_liq.Rows = grid_liq.Rows + 1
  grid_liq.TextMatrix(xf, 0) = pdat_llave!PED_CODVEN
  SQ_OPER = 1
  PUB_KEY = pdat_llave!PED_CODART
  pu_codcia = LK_CODCIA
  LEER_ART_LLAVE
  grid_liq.TextMatrix(xf, 1) = art_LLAVE!ART_NOMBRE
  grid_liq.TextMatrix(xf, 2) = pdat_llave!PED_TIPMOV
  grid_liq.TextMatrix(xf, 3) = pdat_llave!PED_NUMSER
  grid_liq.TextMatrix(xf, 4) = pdat_llave!PED_NUMFAC
pdat_llave.MoveNext
Loop
grid_liq.COL = 1
grid_liq.Row = 1
grid_liq.SetFocus





End Sub

Private Sub Boton_Cons_GotFocus()
Boton_Cons_Click
End Sub

Private Sub Boton_Letras_Click()

grid_fac.Clear
pasa_cabeza

'If Val(gridl.TextMatrix(1, 5)) = 0 And Val(gridl.TextMatrix(1, 4)) = 0 Then Exit Sub

   FILAX = 2
   fila = 2
   Do Until FILAX > gridl.Rows - 1
      SQ_OPER = 1
      PU_TIPMOV = 10
      PU_NUMSER = Val(gridl.TextMatrix(FILAX, 2))
      PU_NUMFAC = Val(gridl.TextMatrix(FILAX, 3))
      pu_codcia = LK_CODCIA
      PU_FBG = gridl.TextMatrix(FILAX, 1)
      LEER_FAR_LLAVE
      If far_llave.EOF = True Then
         MsgBox "No existe Documento ...", 48, Pub_Titulo
         GoTo fuera
     End If
     
     Do Until far_llave.EOF
        grid_fac.Rows = fila + 2
        If far_llave!far_estado = "E" Then GoTo OTRA_FILA
        PUB_KEY = far_llave!far_codart
        pu_codcia = LK_CODCIA
        SQ_OPER = 1
        LEER_ART_LLAVE
        If art_LLAVE.EOF And PUB_KEY <> 0 Then
           MsgBox "Error Grave en arti..."
           GoTo fuera
        End If
        PUB_CODART = far_llave!far_codart
        pu_codcia = LK_CODCIA
        SQ_OPER = 1
        LEER_ARM_LLAVE
        If arm_llave.EOF Then
           MsgBox "Error Grave en arti..."
           GoTo fuera
        End If
      
   If PUB_KEY <> 0 Then grid_fac.TextMatrix(fila, 0) = art_LLAVE!ART_NOMBRE
   grid_fac.TextMatrix(fila, 1) = art_LLAVE!ART_ALTERNO
   grid_fac.TextMatrix(fila, 16) = far_llave!far_codart
   grid_fac.TextMatrix(fila, 33) = far_llave!far_codart
   grid_fac.TextMatrix(fila, 2) = far_llave!far_JABAS
   grid_fac.TextMatrix(fila, 3) = far_llave!far_UNIDADES
   grid_fac.TextMatrix(fila, 11) = far_llave!FAR_COSPRO
   
   If Nulo_Valor0(far_llave!FAR_equiv) > 0 Then
   grid_fac.TextMatrix(fila, 4) = (far_llave!FAR_CANTIDAD - far_llave!FAR_CONSIG) / Nulo_Valor0(far_llave!FAR_equiv)
   grid_fac.TextMatrix(fila, 22) = (far_llave!FAR_CANTIDAD - far_llave!FAR_CONSIG) / Nulo_Valor0(far_llave!FAR_equiv)
   Else
   grid_fac.TextMatrix(fila, 4) = far_llave!FAR_CANTIDAD - far_llave!FAR_CONSIG
   grid_fac.TextMatrix(fila, 22) = CONS_llave!FAR_CANTIDAD - far_llave!FAR_CONSIG
   End If
   
   grid_fac.TextMatrix(fila, 37) = grid_fac.TextMatrix(fila, 22)
   
   grid_fac.TextMatrix(fila, 6) = far_llave!FAR_PRECIO
   
   If LK_FLAG_GRIFO = "A" Then
      grid_fac.TextMatrix(fila, 7) = far_llave!FAR_SUBTOTAL
   Else
      grid_fac.TextMatrix(fila, 7) = redondea(far_llave!FAR_PRECIO * Val(grid_fac.TextMatrix(fila, 4)))
   End If
   
   grid_fac.TextMatrix(fila, 36) = far_llave!FAR_PRECIO
   
   grid_fac.TextMatrix(fila, 14) = Nulo_Valor0(far_llave!FAR_equiv)
   grid_fac.TextMatrix(fila, 5) = Nulo_Valors(far_llave!FAR_descri)
   grid_fac.TextMatrix(fila, 8) = Nulo_Valor0(far_llave!FAR_FLETE)
   grid_fac.TextMatrix(fila, 10) = far_llave!FAR_DESCTO
   grid_fac.TextMatrix(fila, 11) = far_llave!FAR_COSPRO
   grid_fac.TextMatrix(fila, 12) = far_llave!far_SIGNO_aRM
   grid_fac.TextMatrix(fila, 13) = arm_llave!ARM_STOCK
   grid_fac.TextMatrix(fila, 38) = arm_llave!arm_stock_n
   grid_fac.TextMatrix(fila, 39) = arm_llave!arm_stock_s
   
   grid_fac.TextMatrix(fila, 35) = far_llave!FAR_PRECIO
   grid_fac.TextMatrix(fila, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
   grid_fac.TextMatrix(fila, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
   grid_fac.TextMatrix(fila, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
   grid_fac.TextMatrix(fila, 26) = far_llave!FAR_TIPMOV
   grid_fac.TextMatrix(fila, 27) = far_llave!FAR_CODCIA
   grid_fac.TextMatrix(fila, 28) = far_llave!far_numser
   grid_fac.TextMatrix(fila, 29) = far_llave!FAR_FBG
   grid_fac.TextMatrix(fila, 30) = far_llave!FAR_NUMFAC
   grid_fac.TextMatrix(fila, 31) = far_llave!FAR_NUMSEC
   
   fila = fila + 1
OTRA_FILA:
   far_llave.MoveNext
Loop
   FILAX = FILAX + 1
Loop
calcula_totales
Frame4.Visible = True
Exit Sub
fuera:
End Sub

Private Sub cancelar_Click()
FILAX = 0
cancela_todito
gridl.Clear
pasa_cabeza_L
pasa_cabeza
If grid_canje.Visible = True Then pasa_cabeza_canje
grid_fac.Rows = 3
grabar.Enabled = True

If LK_CODTRA = 2101 Then
 i_cambio.Value = 0
End If


If i_def.Visible And i_def.Enabled Then
  i_def.SetFocus
End If

If LK_CODTRA = 2105 Then Frame4.Visible = False

End Sub
Private Sub Compras2_Click()
Dim FF As Integer
Dim WW_IMPORTE As Currency
Dim wcodclie As Currency
Dim wfecha As Date
BolFac_Click
gridC.Clear
gridC.TextMatrix(0, 0) = "Kardex "
gridC.TextMatrix(0, 1) = "Proveedor"
gridC.TextMatrix(0, 2) = "Cantidad"
gridC.TextMatrix(0, 3) = "Precio"

gridC.Visible = True
PSFAR_TEXT(0) = 20
PSFAR_TEXT(1) = LK_CODCIA
PSFAR_TEXT(2) = i_fecha_compra.Text
FAR_TEXT.Requery
If FAR_TEXT.EOF Then
   MsgBox "No hay Ingresos del dia ..."
   Exit Sub
End If
FF = 0
Do Until FAR_TEXT.EOF
   If PUB_NUMFAC <> FAR_TEXT!FAR_NUMFAC Then
      FF = FF + 1
      FF = FF + 1
      gridC.Rows = FF + 1

      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      pu_codclie = FAR_TEXT!far_codclie
      pu_cp = "P"
      LEER_CLI_LLAVE
      gridC.TextMatrix(FF, 0) = FAR_TEXT!FAR_fecha
      gridC.TextMatrix(FF, 1) = cli_llave!cli_nombre
      gridC.TextMatrix(FF, 2) = "Kardex :" & FAR_TEXT!FAR_NUMFAC
      WW_IMPORTE = Val(FAR_TEXT!far_bruto) + Val(FAR_TEXT!far_IMPTO) + Val(FAR_TEXT!FAR_TOT_DESCTO) + Val(FAR_TEXT!FAR_GASTOS)
      gridC.TextMatrix(FF, 3) = "Neto:" & WW_IMPORTE
   End If

   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_KEY = FAR_TEXT!far_codart
   LEER_ART_LLAVE

    FF = FF + 1
    gridC.Rows = FF + 1
    gridC.RowHeight(FF) = 320
    gridC.TextMatrix(FF, 1) = art_LLAVE!ART_NOMBRE
    gridC.TextMatrix(FF, 2) = FAR_TEXT!FAR_CANTIDAD
    gridC.TextMatrix(FF, 3) = FAR_TEXT!FAR_PRECIO
    gridC.TextMatrix(FF, 4) = FAR_TEXT!FAR_estado2
    gridC.TextMatrix(FF, 5) = FAR_TEXT!FAR_FBG
    gridC.TextMatrix(FF, 6) = FAR_TEXT!far_numser
    gridC.TextMatrix(FF, 7) = FAR_TEXT!FAR_NUMFAC
    If FAR_TEXT!FAR_estado2 <> "P" Then
       gridC.Row = FF
       gridC.COL = 0
       gridC.CellBackColor = vbCyan
       gridC.COL = 1
       gridC.CellBackColor = vbCyan
       gridC.COL = 2
       gridC.CellBackColor = vbCyan
    End If
    PUB_NUMFAC = FAR_TEXT!FAR_NUMFAC
    wfecha = FAR_TEXT!FAR_fecha
'    WW_IMPORTE = Val(FAR_TEXT!FAR_BRUTO) + Val(FAR_TEXT!far_impto) + Val(FAR_TEXT!FAR_TOT_DESCTO) + Val(FAR_TEXT!FAR_GASTOS)

    FAR_TEXT.MoveNext
Loop
 
PASALO:
f = f + 1


End Sub

Private Sub Diario_Click()

'MsgBox "XXX"
End Sub

Private Sub grid_liq_KeyPress(KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub

PSFAR_TRANS3.rdoParameters(0) = Val(grid_liq.TextMatrix(grid_liq.Row, 3))
PSFAR_TRANS3.rdoParameters(1) = Val(grid_liq.TextMatrix(grid_liq.Row, 4))
PSFAR_TRANS3.rdoParameters(2) = PUB_TIPMOV '180
PSFAR_TRANS3.rdoParameters(3) = LK_CODCIA
FAR_TRANS3.Requery
If FAR_TRANS3.EOF Then
  MsgBox "No Exiten Datos ..", 48, Pub_Titulo
  Exit Sub
End If
pasa_cabeza
fila = 2
grid_fac.ColWidth(3) = 0
grid_fac.ColWidth(4) = 900
Do Until FAR_TRANS3.EOF
'       If FAR_TRANS3!far_SIGNO_ARM = 1 Then GoTo OTRA_FILA
        grid_fac.Rows = fila + 1
       
        PUB_KEY = FAR_TRANS3!far_codart
        pu_codcia = LK_CODCIA
        SQ_OPER = 1
        LEER_ART_LLAVE
        If art_LLAVE.EOF And PUB_KEY <> 0 Then
           MsgBox "Error Grave en arti..."
           Exit Sub
        End If
        PUB_CODART = FAR_TRANS3!far_codart
        pu_codcia = LK_CODCIA
        SQ_OPER = 1
        LEER_ARM_LLAVE
        If arm_llave.EOF Then
           MsgBox "Error Grave en arti..."
           Exit Sub
        End If
      
      
   If PUB_KEY <> 0 Then grid_fac.TextMatrix(fila, 0) = art_LLAVE!ART_NOMBRE
   grid_fac.RowHeight(fila) = 285
   grid_fac.TextMatrix(fila, 1) = art_LLAVE!ART_ALTERNO
   grid_fac.TextMatrix(fila, 16) = FAR_TRANS3!far_codart
   grid_fac.TextMatrix(fila, 33) = FAR_TRANS3!far_codart
   grid_fac.TextMatrix(fila, 2) = FAR_TRANS3!far_JABAS
   grid_fac.TextMatrix(fila, 3) = ""
   grid_fac.TextMatrix(fila, 11) = arm_llave!arm_cospro
   grid_fac.TextMatrix(fila, 4) = Format(FAR_TRANS3!FAR_CANTIDAD, "0.00")
   grid_fac.TextMatrix(fila, 5) = FAR_TRANS3!FAR_descri
   grid_fac.TextMatrix(fila, 22) = 0
   grid_fac.TextMatrix(fila, 37) = grid_fac.TextMatrix(fila, 22)
   grid_fac.TextMatrix(fila, 6) = 0
   grid_fac.TextMatrix(fila, 7) = 0
   
   grid_fac.TextMatrix(fila, 36) = 0
   
   grid_fac.TextMatrix(fila, 14) = 1
   grid_fac.TextMatrix(fila, 5) = ""
   grid_fac.TextMatrix(fila, 8) = 0
   grid_fac.TextMatrix(fila, 10) = 0
   grid_fac.TextMatrix(fila, 11) = 0
   grid_fac.TextMatrix(fila, 12) = 0
   grid_fac.TextMatrix(fila, 13) = arm_llave!ARM_STOCK
   grid_fac.TextMatrix(fila, 38) = arm_llave!ARM_Saldo_n
   grid_fac.TextMatrix(fila, 39) = arm_llave!ARM_saldo_s
   
   grid_fac.TextMatrix(fila, 35) = 0
   grid_fac.TextMatrix(fila, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
   grid_fac.TextMatrix(fila, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
   grid_fac.TextMatrix(fila, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
   grid_fac.TextMatrix(fila, 26) = 100
   grid_fac.TextMatrix(fila, 27) = LK_CODCIA
   grid_fac.TextMatrix(fila, 28) = Val(grid_liq.TextMatrix(grid_liq.Row, 3))
   grid_fac.TextMatrix(fila, 29) = 0
   grid_fac.TextMatrix(fila, 30) = Val(grid_liq.TextMatrix(grid_liq.Row, 4))
   grid_fac.TextMatrix(fila, 31) = 0
   
   fila = fila + 1
OTRA_FILA:
   FAR_TRANS3.MoveNext
Loop

calcula_totales
Frame4.Visible = True
grid_fac.Visible = True
grid_fac.SetFocus


End Sub

Private Sub gridC_KeyPress(KeyAscii As Integer)
Dim WS_NUMOPER
If KeyAscii <> 13 Then Exit Sub

pub_mensaje = "¿Desea complementar un Ingreso de manera ADICIONAL  ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   gridC.Clear
   GoTo SS1
End If


If gridC.TextMatrix(gridC.Row, 4) = "Q" Then
   MsgBox "Ya procesado..."
   Exit Sub
End If

PU_NUMFAC = Val(gridC.TextMatrix(gridC.Row, 7))
PUB_NUMFAC = Val(gridC.TextMatrix(gridC.Row, 7))

pu_codcia = LK_CODCIA
PU_NUMSER = Val(gridC.TextMatrix(gridC.Row, 6))
PUB_FBG = ""
PU_TIPMOV = PUB_TIPMOV
PU_FBG = PUB_FBG

SQ_OPER = 1
LEER_FAR_LLAVE
If far_llave.EOF = True Then Exit Sub

far_llave.MoveLast
WS_NUMOPER = far_llave!FAR_NUMOPER

fila = 2
far_llave.MoveFirst


Do Until far_llave.EOF
   grid_fac.Rows = fila + 2
   
   If WS_NUMOPER = far_llave!FAR_NUMOPER And far_llave!far_codart <> 0 And (far_llave!far_estado <> "E") Then
      PUB_KEY = far_llave!far_codart
      pu_codcia = LK_CODCIA
      SQ_OPER = 1
      LEER_ART_LLAVE
      If art_LLAVE.EOF And PUB_KEY <> 0 Then
         MsgBox "Error Grave en arti..."
         Exit Sub
      End If
      PUB_CODART = far_llave!far_codart
      pu_codcia = LK_CODCIA
      SQ_OPER = 1
      LEER_ARM_LLAVE
      If arm_llave.EOF Then
         MsgBox "Error Grave en arti..."
         Exit Sub
      End If
      
   
   i_dias.Text = far_llave!far_DIAS
   i_concepto.Text = far_llave!far_concepto
   i_numguia.Text = far_llave!far_NUMGUIA
   i_numfac_c.Text = far_llave!far_numfac_c
   i_numser_c.Text = far_llave!far_numser_c
   i_fecha_compra.Text = far_llave!far_FECHA_COMPRA
   i_fecha_vcto.Text = Str(DateAdd("d", Val(i_dias.Text), i_fecha_compra.Text))
   
   If PUB_KEY <> 0 Then grid_fac.TextMatrix(fila, 0) = art_LLAVE!ART_NOMBRE
   
   grid_fac.TextMatrix(fila, 1) = art_LLAVE!ART_ALTERNO
   grid_fac.TextMatrix(fila, 16) = far_llave!far_codart
   grid_fac.TextMatrix(fila, 33) = far_llave!far_codart
   grid_fac.TextMatrix(fila, 2) = far_llave!far_JABAS
   grid_fac.TextMatrix(fila, 3) = far_llave!far_UNIDADES
   grid_fac.TextMatrix(fila, 11) = far_llave!FAR_COSPRO
   
   If Nulo_Valor0(far_llave!FAR_equiv) > 0 Then
   grid_fac.TextMatrix(fila, 4) = far_llave!FAR_CANTIDAD / Nulo_Valor0(far_llave!FAR_equiv)
   Else
   grid_fac.TextMatrix(fila, 4) = far_llave!FAR_CANTIDAD
   End If
   grid_fac.TextMatrix(fila, 6) = far_llave!FAR_PRECIO
   grid_fac.TextMatrix(fila, 7) = redondea(far_llave!FAR_PRECIO * far_llave!FAR_CANTIDAD)
   
   grid_fac.TextMatrix(fila, 14) = Nulo_Valor0(far_llave!FAR_equiv)
   grid_fac.TextMatrix(fila, 5) = Nulo_Valors(far_llave!FAR_descri)
   grid_fac.TextMatrix(fila, 8) = Nulo_Valor0(far_llave!FAR_FLETE)
   grid_fac.TextMatrix(fila, 10) = far_llave!FAR_DESCTO
   grid_fac.TextMatrix(fila, 11) = far_llave!FAR_COSPRO
   grid_fac.TextMatrix(fila, 12) = far_llave!far_SIGNO_aRM
   grid_fac.TextMatrix(fila, 13) = arm_llave!ARM_STOCK
   grid_fac.TextMatrix(fila, 38) = arm_llave!ARM_Saldo_n
   grid_fac.TextMatrix(fila, 39) = arm_llave!ARM_saldo_s
   
   grid_fac.TextMatrix(fila, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
   grid_fac.TextMatrix(fila, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
   grid_fac.TextMatrix(fila, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
   i_codcli.Text = far_llave!far_codclie
   
   fila = fila + 1
   End If
   far_llave.MoveNext
Loop
SS1:
calcula_totales
pub_flag_cambio = 0
gridC.Visible = False
Frame4.Visible = True
' CAMTEX
grid_fac.COL = 1
grid_fac.SetFocus
flag_salto = 1
flag_salto = 0
End Sub

Private Sub GridD_EnterCell()
If SUT_LLAVE.EOF Then Exit Sub

textovar3.Visible = False
If GridD.Text = "" Then
   PROCESA_CELDAS3
End If

End Sub

Private Sub GridD_KeyPress(KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub

If GridD.COL = 3 Or GridD.COL = 5 Or GridD.COL = 6 Then PROCESA_CELDAS3


End Sub
Private Sub GridD_KeyUp(KeyCode As Integer, Shift As Integer)
Dim fil As Integer
Dim xx As Integer
If KeyCode <> 45 And KeyCode <> 46 Then Exit Sub

If KeyCode = 46 Then GoTo elimina
fil = GridD.Row
GridD.Rows = GridD.Rows + 1
GridD.RowHeight(GridD.Rows - 1) = 315

xx = GridD.Rows - 1
Do Until xx = fil
   GridD.TextMatrix(xx, 0) = GridD.TextMatrix(xx - 1, 0)
   GridD.TextMatrix(xx, 1) = GridD.TextMatrix(xx - 1, 1)
   GridD.TextMatrix(xx, 2) = GridD.TextMatrix(xx - 1, 2)
   GridD.TextMatrix(xx, 3) = GridD.TextMatrix(xx - 1, 3)
   GridD.TextMatrix(xx, 4) = GridD.TextMatrix(xx - 1, 4)
   GridD.TextMatrix(xx, 5) = GridD.TextMatrix(xx - 1, 5)
   GridD.TextMatrix(xx, 6) = GridD.TextMatrix(xx - 1, 6)
   GridD.TextMatrix(xx, 7) = GridD.TextMatrix(xx - 1, 7)
   
   xx = xx - 1
Loop
GridD.TextMatrix(fil, 3) = ""
GridD.TextMatrix(fil, 4) = ""
GridD.TextMatrix(fil, 5) = 0
GridD.TextMatrix(fil, 6) = ""
Exit Sub

elimina:
   GridD.CellBackColor = vbBlue
   
   pub_mensaje = "¿Desea Eliminar la fila seleccionada ... ?"
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   If Pub_Respuesta = vbNo Then
      GridD.CellBackColor = vbWhite
      Exit Sub
   End If
   If GridD.Row = 2 And GridD.Rows = 3 Then
   Else
    GridD.RemoveItem (GridD.Row)
   End If
   RECALCULA
   GridD.SetFocus
   Option5_Click
End Sub

Private Sub i_cias_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
NUMERO = i_cias.WhatsThisHelpID
avanza_grifo

End Sub

Private Sub i_destino_Click()
grid_fac.Clear
End Sub

Private Sub i_destino_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Val(Right(i_destino.Text, 5)) = 4 Then
      grid_fac.TextMatrix(0, 2) = SUT_LLAVE!SUT_JABAS
      grid_fac.ColWidth(2) = 600
Else
      grid_fac.ColWidth(2) = 0
End If

NUMERO = i_destino.WhatsThisHelpID
avanza_grifo

End Sub

Private Sub i_num_lote_GotFocus()
If i_cambio.Value = 0 And LK_CODTRA = 2103 Then llena_pedido

End Sub

Private Sub i_num_lote_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If LK_CODTRA = 2103 And Text1(1).Visible = True Then Text1(1).SetFocus
NUMERO = i_num_lote.WhatsThisHelpID
avanza_grifo

End Sub

Private Sub i_num_lote_LostFocus()
Dim WF As String * 1
Dim i As Integer
Dim ww_flag As String * 1
Dim WS_FECHITA As Date
If i_cambio.Value = 0 Then Exit Sub
If Val(i_num_lote.Text) = 0 Then Exit Sub

Dim WS_NUMFAC, WS_NUMOPER

If LK_CODTRA <> 2103 Then Exit Sub

If Not i_cambio.Visible Then Exit Sub

PUB_PEDFAC = Val(i_num_lote.Text)
pu_codcia = LK_CODCIA
PUB_PEDSER = Val(i_ser_lote.Text)
SQ_OPER = 1
LEER_PED_LLAVE
If ped_llave.EOF = True Then
pub_mensaje = "Desea saltar a este numero...? "
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   llena_pedido
End If
Exit Sub
End If


pub_mensaje = "Existe datos en este Documento : " & Trim(i_ser_lote.Text) & " - " & i_num_lote.Text & " ¿Desea Reemplazar sus datos por el actual ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If



GridT.Clear
GridT.Rows = 3
' QUITE CAMTEX
'WS_NUMOPER = ped_llave!PED_NUMOPER
fila = 2
Do Until ped_llave.EOF
   GridT.Rows = fila + 2
   
      PUB_KEY = ped_llave!PED_CODART
      pu_codcia = LK_CODCIA
      SQ_OPER = 1
      LEER_ART_LLAVE
      If art_LLAVE.EOF And PUB_KEY <> 0 Then
         MsgBox "Error Grave en arti..."
         Exit Sub
      End If
      PUB_CODART = ped_llave!PED_CODART
      pu_codcia = LK_CODCIA
      SQ_OPER = 1
      LEER_ARM_LLAVE
      If arm_llave.EOF Then
         MsgBox "Error Grave en arti..."
         Exit Sub
      End If
      
   i_codcli.Text = ped_llave!PED_CODCLIE
   
   GridT.TextMatrix(fila, 0) = art_LLAVE!ART_NOMBRE
   If LK_FLAG_ORIGINAL = "A" Then
     GridT.TextMatrix(fila, 1) = art_LLAVE!ART_KEY
   Else
     GridT.TextMatrix(fila, 1) = art_LLAVE!ART_ALTERNO
   End If
   
   GridT.TextMatrix(fila, 3) = ped_llave!PED_CANTIDAD
   GridT.TextMatrix(fila, 5) = ped_llave!PED_CODART
   GridT.TextMatrix(fila, 2) = ped_llave!PED_ROLLOS
   GridT.TextMatrix(fila, 6) = ped_llave!PED_CONTACTO
   GridT.TextMatrix(fila, 7) = ped_llave!PED_numfac2
   
   For i = 1 To 14
      Text1(i).Text = ped_llave(30 + i)
   Next i
   fila = fila + 1
   ped_llave.MoveNext
Loop

MUE:
calcula_totales22
'If i_codcli.Visible Then
'  i_codcli.SetFocus
'  i_codcli_KeyPress 13
'End If
   If WS_MONEDA_CLI = "S" Then
      i_ds.ListIndex = 0
   ElseIf WS_MONEDA_CLI = "D" Then
      If Val(i_ds.ListCount) = 1 Then
        i_ds.ListIndex = 0
      Else
        i_ds.ListIndex = 1
      End If
   Else
      i_ds.ListIndex = -1
   End If
textovar.Visible = False
pub_flag_cambio = 0
Exit Sub

End Sub

Private Sub i_origen_Click()
grid_fac.Clear
End Sub

Private Sub i_origen_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

NUMERO = i_origen.WhatsThisHelpID
If i_origen.ListIndex < i_destino.ListCount - 1 Then i_destino.ListIndex = i_origen.ListIndex + 1
avanza_grifo

End Sub

Private Sub i_ser_lote_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
NUMERO = i_ser_lote.WhatsThisHelpID
avanza_grifo

End Sub

Private Sub i_serguia_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
NUMERO = i_serguia.WhatsThisHelpID
avanza_grifo
End Sub

Private Sub lotes_Click()
Dim ww_numfac

Frame4.Visible = False
grid_fac.Clear
pasa_cabeza
grid_trans.Clear
pasa_cabeza_grid_trans
Screen.MousePointer = 11
' CAMTEX
' MODIFICO TODA LA INSTRUCCION
If i_codcli.Text <> "" Then
   pub_cadena = "SELECT PED_CODVEN, PED_NUMFAC2, PED_NUMGUIA, PED_FECHA, PED_CODART, PED_NUMSEC, PED_INGRESOS,PED_SALIDAS, PED_NUMSER, PED_NUMFAC , PED_CANTIDAD, PED_ROLLOS, PED_CODCLIE, ART_FAMILIA  FROM PEDIDOS, ARTI WHERE (PED_CODCIA = ART_CODCIA) AND (PED_CODART = ART_KEY) AND PED_CODCIA = ? AND PED_CODCLIE = ? AND PED_TIPMOV = ? AND ART_FAMILIA = ? ORDER BY PED_FECHA, PED_NUMSER, PED_NUMFAC "
End If
Set PSPED_MAYOR = CN.CreateQuery("", pub_cadena)

If i_codcli.Text <> "" Then
   PSPED_MAYOR.rdoParameters(0) = 0
   PSPED_MAYOR.rdoParameters(1) = 0
   PSPED_MAYOR.rdoParameters(2) = 0
   PSPED_MAYOR.rdoParameters(3) = 0
Else
   Screen.MousePointer = 0
   Exit Sub
End If
Set ped_mayor = PSPED_MAYOR.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PSPED_MAYOR.rdoParameters(0) = LK_CODCIA
PSPED_MAYOR.rdoParameters(1) = Val(i_codcli.Text)
If LK_CODTRA = 2212 Then
 PSPED_MAYOR.rdoParameters(2) = 177
Else
 PSPED_MAYOR.rdoParameters(2) = PUB_TIPMOV
End If
PSPED_MAYOR.rdoParameters(3) = SUT_LLAVE!SUT_art_gru
' MODIFICO TODA LA INSTRUCCION HASTA AQUI
ped_mayor.Requery
If ped_mayor.EOF Then
   Screen.MousePointer = 0
   MsgBox "No hay Registros ...Verificar datos de Busqueda"
   Exit Sub
End If
Screen.MousePointer = 0
' CAMTEX
If PUB_SECUENCIA = 1 Then
WGUIA = InputBox("Filtra por Guia", "BUSCAR...", WGUIA)
Else
WGUIA = ""
End If
Screen.MousePointer = 11
fila = 0
ww_numfac = 0
PUB_CHENUM = ped_mayor!PED_NUMFAC
Do Until ped_mayor.EOF
   ' CAMTEX
   If Trim(WGUIA) <> "" Then
     If (WGUIA <> ped_mayor!PED_NUMGUIA) Then GoTo pasa
   End If
   
   GoSub muestra
pasa:
   PUB_CHENUM = ped_mayor!PED_NUMFAC
   ped_mayor.MoveNext
Loop
grid_trans.Visible = True
grid_trans.Height = Frame4.Height - 1000
grid_trans.Width = Frame4.Width
grid_trans.Top = Frame4.Top + 400
label_grid.Top = Frame4.Top
grid_trans.Left = 0
label_grid.Left = 0
label_grid.Width = Frame4.Width
label_grid.Visible = True
grid_trans.Left = 0

grid_trans.SetFocus
Frame4.Visible = False
Screen.MousePointer = 0
Exit Sub

muestra:
      If PUB_CHENUM <> ped_mayor!PED_NUMFAC Then fila = fila + 1
     
      fila = fila + 1
      grid_trans.Rows = fila + 1
      grid_trans.TextMatrix(fila, 0) = Format(ped_mayor!PED_FECHA, "dd/mm/yy")
      grid_trans.TextMatrix(fila, 10) = ped_mayor!PED_NUMSER
      grid_trans.TextMatrix(fila, 1) = ped_mayor!PED_NUMFAC
      PUB_KEY = ped_mayor!PED_CODART
      pu_codcia = LK_CODCIA
      SQ_OPER = 1
      LEER_ART_LLAVE
      grid_trans.TextMatrix(fila, 2) = art_LLAVE!ART_NOMBRE
      grid_trans.TextMatrix(fila, 3) = ped_mayor!PED_ROLLOS
      grid_trans.TextMatrix(fila, 4) = ped_mayor!PED_CANTIDAD
      grid_trans.TextMatrix(fila, 5) = PUB_KEY
      grid_trans.TextMatrix(fila, 6) = ped_mayor!PED_NUMSEC
      grid_trans.TextMatrix(fila, 11) = ped_mayor!PED_NUMGUIA
      grid_trans.TextMatrix(fila, 12) = Nulo_Valor0(ped_mayor!PED_numfac2)
      grid_trans.TextMatrix(fila, 13) = Nulo_Valor0(ped_mayor!PED_CODVEN)
      
      Return


End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   grid_fac.Clear
   pasa_cabeza
   MsgBox "Ingresar el numero de Kardex , con el que se realizó el Ingreso , Si no conoce el nombre lo puede averiguar en la Opción Herramientas-Consultas-Kardex"
   i_numser.SetFocus
Else
   cancela_todo
End If
End Sub

Private Sub Diario_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = Diario.WhatsThisHelpID
avanza_grifo
fin:

End Sub

Private Sub Form_Click()
WhatsThisMode
End Sub
Private Sub Form_Load()
Dim ws_indice As Integer
Dim CONTA
Dim ws_ancho As Integer
If LK_EMP = "3AA" Then
 Screen.MousePointer = 0
 MsgBox "No Activo para este tipo de Empresa", 48, Pub_Titulo
 Unload FORM_GRIFO
 Exit Sub
End If
'DoEvents
If LK_FECHA_DIA = #1/1/1900# Then
  Screen.MousePointer = 0
  MsgBox "La Compañia No Tiene Definida su Fecha de Trabajo. Verificar!!!", 48, Pub_Titulo
  Unload FORM_GRIFO
  Exit Sub
End If
grid_fac.FontName = "Arial"
grid_fac.FontSize = 8

grid_fac.Cols = 40
grid_fac.ColWidth(0) = 2000 'nombre
grid_fac.ColWidth(1) = 800 'codigo
grid_fac.ColWidth(2) = 600 'jabas
grid_fac.ColWidth(3) = 400 'unidades
grid_fac.ColWidth(4) = 600 'cantidad
grid_fac.ColWidth(5) = 1000 'unidades
grid_fac.ColWidth(6) = 1000 'precio
grid_fac.ColWidth(7) = 1300 'subtotal
grid_fac.ColWidth(8) = 600 'flete
grid_fac.ColWidth(9) = 2000 'mortalidad
grid_fac.ColWidth(10) = 1000 'FLAG_DESCTO=D
grid_fac.ColWidth(11) = 900 'COSTO PROMEDIO
grid_fac.ColWidth(12) = 900 'signo del arm + -1
grid_fac.ColWidth(13) = 0 'STOCK
grid_fac.ColWidth(14) = 0 'EQUIV
grid_fac.ColWidth(15) = 900 'DESCRIP. UNIDAD
grid_fac.ColWidth(16) = 0 'codigo original
grid_fac.ColWidth(17) = 0 'NUMERO DE PRECIO 1.2,3,4,5
grid_fac.ColWidth(18) = 0 'relacion con un pedido
grid_fac.ColWidth(19) = 0 'orden en i_precios
grid_fac.ColWidth(20) = 0 'NUMSEC PARA UPDATE DE PEDIDOS
grid_fac.ColWidth(21) = 0 'FLAG DE STOCK
grid_fac.ColWidth(22) = 0 'numero de autorizacion
grid_fac.ColWidth(23) = 0 'FLAG DE EXONERACION IGV
grid_fac.ColWidth(24) = 0 '% IGV DE EXONERACION
grid_fac.ColWidth(25) = 0 'FLAG PARA COMPRA SUBTOTAL DIGITADO
grid_fac.ColWidth(26) = 0 'RECEP. TIPMOV
grid_fac.ColWidth(27) = 0 'RECEP. CODCIA
grid_fac.ColWidth(28) = 0 'RECEP. NUMSER
grid_fac.ColWidth(29) = 0 'RECEP. FBG
grid_fac.ColWidth(30) = 0 'RECEP. NUMFAC
grid_fac.ColWidth(31) = 0 'RECEP. NUMSEC
grid_fac.ColWidth(32) = 0 'precio de venta cash dpv O MINIMO
grid_fac.ColWidth(33) = 0 'PRE_CODART
grid_fac.ColWidth(34) = 0 'DISPONIBLE
grid_fac.ColWidth(35) = 0 'precio BASE
grid_fac.ColWidth(36) = 0 'PRECIO LISTA
grid_fac.ColWidth(37) = 0 'CANTI. CONSIG.
grid_fac.ColWidth(38) = 0 'CANTI. A PEDIR PARA REPOSICION(2103)
grid_fac.ColWidth(39) = 0 'NRO. DE PED. SEGUNDO

gridC.Cols = 8
gridC.ColWidth(0) = 1200 '
gridC.ColWidth(1) = 2800 'codigo
gridC.ColWidth(2) = 1400 'jabas
gridC.ColWidth(3) = 1400
gridC.ColWidth(4) = 0
gridC.ColWidth(5) = 0
gridC.ColWidth(6) = 0
gridC.ColWidth(7) = 0



Frame5.Width = 11000
Frame5.Height = 4000
GridT.Width = 8500
GridT.Height = 4500
GridD.Height = 4500

GridT.Cols = 8
GridT.ColWidth(0) = 4500 'nombre
GridT.ColWidth(1) = 700 'codigo
GridT.ColWidth(2) = 900 'rollos
GridT.ColWidth(3) = 900 'cantidad
GridT.ColWidth(4) = 0 'STOCK
GridT.ColWidth(5) = 0 'codigo original
GridT.ColWidth(6) = 500 'ok
GridT.ColWidth(7) = 700 'Nro. Pedido

GridD.Clear
GridD.Cols = 9
GridD.MergeCells = flexMergeFree
GridD.MergeCol(0) = True
GridD.MergeCol(1) = True
GridD.MergeCol(2) = True
GridD.MergeCol(3) = False

GridD.ColWidth(0) = 2800
GridD.ColWidth(1) = 700
GridD.ColWidth(2) = 700
GridD.ColWidth(3) = 900
GridD.ColWidth(4) = 3600
GridD.ColWidth(5) = 600
GridD.ColWidth(6) = 1000
GridD.ColWidth(7) = 0
GridD.ColWidth(8) = 0

GridD.Width = 12500






gridl.Cols = 27
gridl.ColWidth(0) = 2000
gridl.ColWidth(1) = 500
gridl.ColWidth(2) = 500
gridl.ColWidth(3) = 500
gridl.ColWidth(4) = 1000
gridl.ColWidth(5) = 1000
gridl.ColWidth(6) = 1000
gridl.ColWidth(7) = 500
gridl.ColWidth(8) = 1000
gridl.ColWidth(9) = 1000
gridl.ColWidth(10) = 0
gridl.ColWidth(11) = 0
gridl.ColWidth(12) = 0
gridl.ColWidth(13) = 0
gridl.ColWidth(14) = 0
gridl.ColWidth(15) = 0
gridl.ColWidth(16) = 0
gridl.ColWidth(17) = 0
gridl.ColWidth(18) = 400
gridl.ColWidth(19) = 1500
gridl.ColWidth(20) = 1500
gridl.ColWidth(21) = 1500
gridl.ColWidth(22) = 1500
gridl.ColWidth(23) = 2500
gridl.ColWidth(24) = 0 ' CAR_MONEDA
gridl.ColWidth(25) = 0 ' CAR_IMPORTE ESPEJO
gridl.ColWidth(26) = 0 ' CAR_fecha vcto espejo


gridl.RowHeight(0) = 350
gridl.RowHeight(1) = 350

pasa_cabeza_L

pasa_cabeza_L2

grid_canje.Cols = 8
grid_canje.ColWidth(0) = 0 'codigo cliente
grid_canje.ColWidth(1) = 450 'orden
grid_canje.ColWidth(2) = 400 'tipo doc
grid_canje.ColWidth(3) = 1000 'importe
grid_canje.ColWidth(4) = 1100 'fecha
grid_canje.ColWidth(5) = 1300 'banco
grid_canje.ColWidth(6) = 800 'n.doc
grid_canje.ColWidth(7) = 1200 'anotacion

grid_canje.RowHeight(2) = 250


grid_canje.RowHeight(0) = 250

grid_fac.RowHeight(1) = 400
grid_fac.RowHeight(2) = 350


RAYA.Y1 = 1200
RAYA.Y2 = 1200
   ' CAMTEX
   Frame4.Width = 8400 ' 7200  '7200
   Frame4.Height = 3800 '3100
   Frame2.Top = 3000  ' 2300
   grid_fac.Width = 8400 ' 7300    '7100
   grid_fac.Height = 2700  '2000

label_nomart.Top = 2500
label_nomart.Left = 0
label_precio.Left = 6000
label_nomart.Visible = False
label_precio.Visible = False
label_nomart.Top = Frame2.Top - 200
label_precio.Top = Frame2.Top - 200

i_image_llave.Left = 7000
i_image_llave.Top = 3000
i_image_llave.Visible = True

grid_fac.ColAlignment(0) = 1
grid_fac.ColAlignment(1) = 1
grid_fac.ColAlignment(5) = 1

grid_autorizacion.FontName = "Sans Serif"
grid_autorizacion.FontSize = 10
grid_canje.FontName = "Courier New"
grid_canje.FontSize = 10

Grid_all.FontName = "Sans Serif"
Grid_all.FontSize = 7



grid_trans.Cols = 14
grid_trans.ColWidth(0) = 1000
grid_trans.ColWidth(1) = 700
grid_trans.ColWidth(2) = 2500
grid_trans.ColWidth(3) = 800
grid_trans.ColWidth(4) = 1200
grid_trans.ColWidth(5) = 0
grid_trans.ColWidth(6) = 0
grid_trans.ColWidth(7) = 0
grid_trans.ColWidth(8) = 0
grid_trans.ColWidth(9) = 0
grid_trans.ColWidth(10) = 0
grid_trans.ColWidth(11) = 800
grid_trans.ColWidth(11) = 900



Grid_all.Top = 1200
Grid_all.Left = 300
Grid_all.Width = 8400
Grid_all.Height = 3800

i_concepto.Width = 3000

nomtra.Width = 2500




'DoEvents




gridl.Top = 1500
gridl.Left = 0
gridl.Width = 7500
gridl.Height = 2000

grid_autorizacion.Left = 0
grid_autorizacion.Top = 2500
grid_autorizacion.Width = 7400
grid_autorizacion.Height = 1500

numarchi = 2
PS_GEN(0) = 0
GEN.Requery
'On Error GoTo SALE100

LK_COBRADOR = 1
PUB_FECHA = LK_FECHA_DIA

PUB_TIPREG = 998
PUB_CODCIA = "00"
SQ_OPER = 2
LEER_TAB_LLAVE
Do Until tab_mayor.EOF
         ws_indice = tab_mayor!TAB_NUMTAB
         TABLA_TAG(ws_indice) = Val(tab_mayor!tab_nomcorto)
         tab_mayor.MoveNext
Loop
    
    
    
    CONTA = -1
    PUB_TIPREG = 122
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
        i_origen.AddItem Left(tab_mayor!TAB_NOMLARGO, 20) & "  " & tab_mayor!TAB_NUMTAB
        i_destino.AddItem Left(tab_mayor!TAB_NOMLARGO, 20) & "  " & tab_mayor!TAB_NUMTAB
        tab_mayor.MoveNext
    Loop
    
    
    CONTA = -1
    PUB_TIPREG = 52
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
        i_mortal.AddItem Left(tab_mayor!TAB_NOMLARGO, 20) & "  " & tab_mayor!TAB_NUMTAB
        tab_mayor.MoveNext
    Loop
    CONTA = 0
    PUB_TIPREG = 45
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    Do Until CONTA = 5
        CONTA = CONTA + 1
        If tab_mayor.EOF Then
           tab_precioss(CONTA) = "....."
        Else
           tab_precioss(CONTA) = tab_mayor!TAB_NOMLARGO
           tab_mayor.MoveNext
        End If
    Loop
    
    
   
If LK_FLAG_FACTURACION = "V" Then i_fbg.AddItem "P"

   textovar_canje.Width = 30
   textovar_canje.Height = 30


i_def.Visible = False
grabar.Enabled = False
cancelar.Enabled = False
tab_derL(5) = 6
tab_derL(6) = 7
tab_derL(7) = 5

  If i_ds.ListCount > 0 Then
    i_ds.ListIndex = 0
  End If

      pub_cadena = "SELECT * FROM TABLAS WHERE TAB_TIPREG = 2101 AND TAB_CODCLIE = ? AND TAB_CODCIA = ? ORDER BY TAB_NOMLARGO"
      Set PS_PLACAS = CN.CreateQuery("", pub_cadena)
      PS_PLACAS.rdoParameters(0) = 0
      PS_PLACAS.rdoParameters(1) = 0
      Set TAB_PLACAS = PS_PLACAS.OpenResultset(rdOpenKeyset, rdConcurValues)
    
      pub_cadena = "SELECT * FROM TABLAS WHERE TAB_TIPREG = 2301 AND TAB_CODCLIE = ? AND TAB_CODCIA = ? AND TAB_CODART = ? "
      Set PS_PRECIOS = CN.CreateQuery("", pub_cadena)
      PS_PRECIOS.rdoParameters(0) = 0
      PS_PRECIOS.rdoParameters(1) = 0
      PS_PRECIOS.rdoParameters(2) = 0
      Set TAB_PRECIOS = PS_PRECIOS.OpenResultset(rdOpenKeyset, rdConcurValues)
      
  If LK_FLAG_GRIFO = "A" Then
      pub_cadena = "SELECT * FROM VEMAEST2 WHERE VE2_CODCIA = ? AND VE2_ISLA =? AND VE2_TURNO=? "
      Set PS_VE2 = CN.CreateQuery("", pub_cadena)
      PS_VE2.rdoParameters(0) = 0
      PS_VE2.rdoParameters(1) = 0
      PS_VE2.rdoParameters(2) = 0
      Set VE2_LLAVE = PS_VE2.OpenResultset(rdOpenKeyset, rdConcurValues)
  End If

  PUB_TIPREG = 2102
  PUB_CODCIA = LK_CODCIA
  SQ_OPER = 2
  LEER_TAB_LLAVE
  If tab_mayor.EOF Then
    If LK_FLAG_GRIFO = "A" Then
      MsgBox "Aperturar Turnos....", 48, Pub_Titulo
    End If
  End If
  tab_turnos(2) = ""
  tab_turnos(3) = ""
  tab_turnos(4) = ""
  Do Until tab_mayor.EOF
     tab_turnos(tab_mayor!TAB_NUMTAB) = tab_mayor!TAB_NOMLARGO
     tab_mayor.MoveNext
  Loop
  
  PUB_TIPREG = 2103
  PUB_CODCIA = LK_CODCIA
  SQ_OPER = 2
  LEER_TAB_LLAVE
  If tab_mayor.EOF Then
    If LK_FLAG_GRIFO = "A" Then
      MsgBox "No hay Responsables por Isla...", 48, Pub_Titulo
    End If
  End If
  Do Until tab_mayor.EOF
     i_responsable.AddItem tab_mayor!TAB_NOMLARGO & "               " & tab_mayor!TAB_NUMTAB
     tab_mayor.MoveNext
  Loop
If LK_FLAG_GRIFO = "A" Then
  PUB_TIPREG = 2105
  PUB_CODCIA = LK_CODCIA
  SQ_OPER = 2
  LEER_TAB_LLAVE
  If tab_mayor.EOF Then
      MsgBox "Relacionar Islas con Productos ...", 48, Pub_Titulo
      Exit Sub
  Else
  FILAX = 1
  Do Until tab_mayor.EOF
     tab_items(Val(tab_mayor!TAB_contable2), FILAX) = tab_mayor!tab_codART
     FILAX = FILAX + 1
     tab_mayor.MoveNext
  Loop
  End If
End If
  
  
i_fbg.ListIndex = 0
i_numfac.BackColor = QBColor(7)
i_numser.BackColor = QBColor(7)
'If LK_FLAG_GRIFO = "A" Then
  pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ?  AND FAR_CODCLIE = ?  AND FAR_CANTIDAD <> FAR_CONSIG AND ( FAR_ESTADO2='C' OR FAR_ESTADO2='F') ORDER BY FAR_TIPMOV, FAR_NUMSER, FAR_NUMFAC" ' LE QUITO FAR_TIPMOV=10   'M ES PORQUE SOLO DESEA CANJES DE GUIAS POR SU FACTURA QUE RECLAMA
  Set PSCONS_LLAVE = CN.CreateQuery("", pub_cadena)
  PSCONS_LLAVE.rdoParameters(0) = 0
  PSCONS_LLAVE.rdoParameters(1) = 0
  Set CONS_llave = PSCONS_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
'End If
' SOLO PARA TEXTIL...


  pub_cadena = "SELECT * FROM FACART WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_FECHA = ? and  FAR_ESTADO2='P'  ORDER BY FAR_NUMSER, FAR_NUMFAC "
  Set PF_BUSCA = CN.CreateQuery("", pub_cadena)
  PF_BUSCA(0) = 0
  PF_BUSCA(1) = ""
  PF_BUSCA(2) = LK_FECHA_DIA
  Set pff_buscar = PF_BUSCA.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)



  pub_cadena = "SELECT * FROM FACART WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_FECHA = ? and  FAR_ESTADO2='P'  ORDER BY FAR_NUMSER, FAR_NUMFAC "
  Set PSFAR_TEXT = CN.CreateQuery("", pub_cadena)
  PSFAR_TEXT(0) = 0
  PSFAR_TEXT(1) = ""
  PSFAR_TEXT(2) = LK_FECHA_DIA
  Set FAR_TEXT = PSFAR_TEXT.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

' FIN DE SOLO PARA TEXTIL




'  pub_cadena = "SELECT PED_NUMFAC FROM PEDIDOS WHERE PED_TIPMOV = ? AND PED_CODCIA = ?  AND PED_NUMSER = ? ORDER BY  PED_NUMFAC DESC"
'  Set PSPED_MENOR = CN.CreateQuery("", pub_cadena)

'  PSPED_MENOR.MaxRows = 1
'  Set ped_menor = PSPED_MENOR.OpenResultset(rdOpenKeyset, rdConcurValues)

  
  
pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ? AND TAB_CODCLIE = ? and TAB_NOMLARGO = ?  ORDER BY TAB_NOMLARGO"
Set PSPLAC_LLAVE = CN.CreateQuery("", pub_cadena)
PSPLAC_LLAVE.rdoParameters(0) = 0
PSPLAC_LLAVE.rdoParameters(1) = 0
PSPLAC_LLAVE.rdoParameters(2) = 0
PSPLAC_LLAVE.rdoParameters(3) = 0
Set cliplac_llave = PSPLAC_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
    
pub_cadena = "SELECT TAB_NUMTAB FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ? ORDER BY TAB_NUMTAB DESC"
Set PSPLAC_SECU = CN.CreateQuery("", pub_cadena)
PSPLAC_SECU.rdoParameters(0) = 0
PSPLAC_SECU.rdoParameters(1) = 0
PSPLAC_SECU.MaxRows = 1
Set cliplac_secu = PSPLAC_SECU.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

pub_cadena = "SELECT PED_NUMFAC FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_TIPMOV = ? AND PED_NUMSER = ? ORDER BY  PED_NUMFAC DESC"
Set PSPED_MENOR = CN.CreateQuery("", pub_cadena)
PSPED_MENOR(0) = 0
PSPED_MENOR(1) = 0
PSPED_MENOR(2) = 0
PSPED_MENOR.MaxRows = 1
Set ped_menor = PSPED_MENOR.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

pub_cadena = "SELECT FAR_CODART ,SUM(FAR_UNIDADES*FAR_SIGNO_ARM),SUM(FAR_CANTIDAD*FAR_SIGNO_ARM)  FROM facart WHERE  FAR_ESTADO <> 'E' AND FAR_PEDSER = ? AND FAR_PEDFAC = ?  AND FAR_NUM_LOTE= ?  AND FAR_CODCIA= ? GROUP BY FAR_CODART, FAR_SIGNO_ARM"
Set PSFAR_TRANS = CN.CreateQuery("", pub_cadena)
PSFAR_TRANS.rdoParameters(0) = 0
PSFAR_TRANS.rdoParameters(1) = 0
PSFAR_TRANS.rdoParameters(2) = 0
PSFAR_TRANS.rdoParameters(3) = 0
Set FAR_TRANS = PSFAR_TRANS.OpenResultset(rdOpenKeyset, rdConcurValues)


pub_cadena = "SELECT FAR_CODART ,SUM(FAR_UNIDADES*FAR_SIGNO_ARM),SUM(FAR_CANTIDAD*FAR_SIGNO_ARM)  FROM facart WHERE  FAR_ESTADO <> 'E' AND FAR_NUMFAC_C = ?  AND FAR_NUM_LOTE= ?  AND FAR_CODCIA= ? GROUP BY FAR_CODART, FAR_SIGNO_ARM"
Set PSFAR_TRANS2 = CN.CreateQuery("", pub_cadena)
PSFAR_TRANS2.rdoParameters(0) = 0
PSFAR_TRANS2.rdoParameters(1) = 0
PSFAR_TRANS2.rdoParameters(2) = 0
Set FAR_TRANS2 = PSFAR_TRANS2.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT FAR_CODART , FAR_CANTIDAD, FAR_DESCRI, FAR_JABAS FROM FACART WHERE  FAR_ESTADO <> 'E' AND FAR_PEDSER = ? AND FAR_PEDFAC = ?  AND FAR_TIPMOV = ?  AND FAR_CODCIA= ? AND FAR_TRANSITO = 'P' AND FAR_ESTADO <> 'E'   ORDER BY FAR_NUMSEC"
Set PSFAR_TRANS3 = CN.CreateQuery("", pub_cadena)
PSFAR_TRANS3.rdoParameters(0) = 0
PSFAR_TRANS3.rdoParameters(1) = 0
PSFAR_TRANS3.rdoParameters(2) = 0
PSFAR_TRANS3.rdoParameters(3) = 0
Set FAR_TRANS3 = PSFAR_TRANS3.OpenResultset(rdOpenKeyset, rdConcurValues)


pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_TIPMOV = ? AND PED_NUMSER = ? AND PED_NUMFAC = ? AND PED_CODCLIE = ? and PED_CODVEN = ?"
Set PSDAT_LLAVE = CN.CreateQuery("", pub_cadena)
PSDAT_LLAVE(0) = 0
PSDAT_LLAVE(1) = 0
PSDAT_LLAVE(2) = 0
PSDAT_LLAVE(3) = 0
PSDAT_LLAVE(4) = 0
PSDAT_LLAVE(5) = 0
Set pdat_llave = PSDAT_LLAVE.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

i_ser_lote.Locked = True
i_num_lote.Locked = True
i_ser_lote.BackColor = QBColor(7)
i_num_lote.BackColor = QBColor(7)


Exit Sub
SALE100:
MsgBox "Procesar la 5555"
MsgBox Err.Description
Resume Next


End Sub

Public Sub PROCESA_CELDASL()
'If gridl.Col < 3 Then Exit Sub
'If gridl.Col = 5 Then Exit Sub
If gridl.CellWidth < 0 Then Exit Sub
textovarl.Locked = False


    textovarl.Left = gridl.Left + gridl.CellLeft
    textovarl.Width = gridl.CellWidth
    textovarl.Height = gridl.CellHeight
    textovarl.Top = gridl.Top + gridl.CellTop
    textovarl.Text = gridl.Text
    If textovarl.Visible = False Then textovarl.Visible = True
    textovar_bak = textovarl.Text
    textovarl.SetFocus
 
End Sub
Private Sub ACT_FORM_GRIFO()
 Dim i
 Dim C As Integer
    PUB_TIPREG = 998
    PUB_CODCIA = "00"
    SQ_OPER = 2
    LEER_TAB_LLAVE
     Do Until tab_mayor.EOF
        tab_mayor.Delete
        tab_mayor.MoveNext
     Loop
 C = 0
 For i = 0 To FORM_GRIFO.Controls.count - 1
     Print "                " & FORM_GRIFO.Controls(i).Name
       If Not (FORM_GRIFO.Controls(i).Tag = 0 Or FORM_GRIFO.Controls(i).Tag > 200) Then
        tab_mayor.AddNew
        tab_mayor!TAB_CODCIA = "00"
        tab_mayor!TAB_TIPREG = 998
        tab_mayor!TAB_NUMTAB = FORM_GRIFO.Controls(i).Tag
        
        tab_mayor!tab_nomcorto = i
        tab_mayor!TAB_NOMLARGO = FORM_GRIFO.Controls(i).Name
        tab_mayor.Update
        C = C + 1
      End If
 Next i
MsgBox "Cantidad de Controles Actualizados....  : " & C & "  O K !!", 48, Pub_Titulo
Unload FORM_GRIFO

End Sub
Private Sub grabar_Click()
'On Error GoTo Error_fatal
Dim FLAG_2406 As String * 1
Dim dispositivo As Integer
Dim WW_NUM_OPER As Integer
Dim ws_fecha_ant As Date
Dim ww_numdoc As Long
Dim ww_numser As Integer
Dim wformula, wformula1, wformula2
Dim ext_conta2 As Integer
Dim DIA, MES, ANO
Dim er As rdoError
Dim pub_mensaje As String
Dim ws_cospro, ws_compra, ws_valor_ant As Currency
Dim flag_repite As Integer
Dim ws_fecha_anterior As Date
Dim wS_saldo_car As Currency
Dim WS_NUM_OPER2 As Integer
Dim WW_NUM_OPER2 As Integer
Dim ws_codcia As String * 2
Dim WS_IMPORTE_AMORT As Currency
Dim WS_IMPORTE_AMORT_DOLL As Currency
Dim SS_IMPORTE_AMORT As Currency
Dim SS_IMPORTE_AMORT_DOLL As Currency
Dim WS_TOT As Currency
Dim FILAXX As Integer
Dim WZONA As String
Dim WSUBZONA As String
Dim WSLUGAR As String
Dim WS_AUTOCON As String
Dim WS_MONEDA As String * 1
Dim ww_numsec As Integer
Dim ws_mon_canje As String * 1
Const ingre = 2
Const MODIF = 1
Dim i As Integer
Dim N As Integer
Dim LOC_SALDO_CAR As Currency
Dim WS_DESCTO1, WS_DESCTO2 As Currency
Dim WS_CANTIDAD As Currency
Dim count As Integer
Dim ww_concepto As String * 150
Dim zona_nombre As String
Dim subzona_nombre  As String
Dim WS_IMPRESION_LET As Long
Dim WS_CHENUM As Currency
Dim WS_NUM_OPER  As Integer
Dim WS_DATOS As String * 2
Dim WS_FLAG As Integer
Dim WS_TRANSITO As String * 1
Dim WS_FLAG2 As Integer
Dim subtotal As Currency
Dim WS_BRUTO2 As Currency
Dim ws_diferencia As Currency
Dim WS_CORRELATIVO As Double
Dim NETO As Currency
Dim fx As Integer
Dim FLAG As Boolean

Dim ws_flag_arm As Integer
Dim ws_flag_che As Integer
Dim ws_flag_far As Integer
Dim ws_flag_car As Integer
Dim ws_flag_par As Integer
Dim ws_flag_ccm As Integer
Dim ws_flag_ven As Integer
Dim filafac As Integer
Dim WS_FLAG_FACART As String * 1
WS_FLAG_FACART = ""
Dim RUTA
Dim TEM_SIGNO_CAR As Integer
Dim plaza_tasa1
Dim plaza_tasa2
pub_responsable = Val(Right(i_responsable.Text, 5))

'i_ds.ListIndex = 0
'i_moneda.Text = "S/."
'WS_MONEDA_CLI = "S"
'PUB_FECHA_VCTO = LK_FECHA_DIA
ext_conta2 = 1
inicio_ext:   'inicio loop de extornos


TEM_SIGNO_CAR = pub_signo_car

If (LK_CODTRA = 2410 Or LK_CODTRA = 2412) And LK_EMP = "PIU" And Val(SUT_LLAVE!SUT_SECUENCIA) = 0 Then
  If Val(gridl.TextMatrix(1, 5)) = 0 Then
    pub_signo_car = 0
  End If
End If

If consis() = False Then Exit Sub

ww_concepto = ""
ww_codtra_ext = 0
filafac = 1
WS_NUMSEC = 0

If Val(GridT.TextMatrix(2, 3)) = 0 And LK_CODTRA = 2103 Then
 pub_mensaje = "Es Anulacion ¿Desea Continuar ... ?"
 Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
 If Pub_Respuesta = vbNo Then
     Exit Sub
 End If
   'MsgBox " Falta Procesar Detalle , Resumen ..."
  ' Exit Sub
End If



If grid_fac.Visible = True And LK_CODTRA <> 1111 And LK_CODTRA <> 1122 Then
   subtotal = 0
   fila = 2
   Do Until fila > 900
     If (grid_fac.TextMatrix(fila, 21)) = "A" Then GoTo mas
     
     If Nulo_Valors(SUT_LLAVE!SUT_CERO) = "A" And LK_CODTRA = 2210 Then GoTo mas
     
     If (Val(grid_fac.TextMatrix(fila, 4)) = 0 Or Trim(grid_fac.TextMatrix(fila, 6)) = "") And Val(grid_fac.TextMatrix(fila, 16)) <> 0 Then
            fila = 999
            Exit Do
     End If
     If Val(grid_fac.TextMatrix(fila, 16)) <> Val(grid_fac.TextMatrix(fila, 33)) And Val(grid_fac.TextMatrix(fila, 16)) > 0 Then
            fila = 888
            Exit Do
     End If

     If Nulo_Valors(SUT_LLAVE!SUT_CERO) = "A" Then GoTo mas
     
         If (Val(grid_fac.TextMatrix(fila, 4)) = 0 Or Val(grid_fac.TextMatrix(fila, 6)) = 0 Or Val(grid_fac.TextMatrix(fila, 7)) = 0) And Val(grid_fac.TextMatrix(fila, 16)) <> 0 Then
            fila = 999
            Exit Do
         End If
         If (Trim(grid_fac.TextMatrix(fila, 4)) <> "") And Trim(grid_fac.TextMatrix(fila, 0)) = "" Then
            fila = 999
            Exit Do
         End If

mas:
      fila = fila + 1
      If fila = grid_fac.Rows Then fila = 998
   Loop
   If fila = 999 Then
      MsgBox " Falta ingresar los datos en algun articulo...", 48, Pub_Titulo
      If grid_fac.Visible Then grid_fac.SetFocus
      Exit Sub
   End If
   If fila = 888 Then
      MsgBox " revisar unidades y precios  de algun articulo .."
      Exit Sub
   End If
   
   If fila = 997 Then
      Exit Sub
   End If
   
End If


PASO:

If PUB_TIPMOV <> 0 Then
      i_importe_amort.Text = Val(i_neto.Text)
End If


exito = True

TOMA_DATOS
'CAPTURA_DATOS



SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_PAR_LLAVE
If par_llave!par_flag_cierre = 9 Then
   MsgBox "!!! Compañia ... Cerró Operaciones ... Llamar al Administrador ", 48, Pub_Titulo
   Exit Sub
End If
If LK_FLAG_GRIFO <> "A" Then
If par_llave!PAR_FECHA_DIA <> LK_FECHA_DIA Then
   MsgBox "!!!FECHA YA NO COINCIDE CON LA ACTUAL , OTRO USUARIO A CERRADO EL DIA!!! SALGA Y REINICIE SU SISTEMA...", 48, Pub_Titulo
   Exit Sub
End If
End If
'YO AGREGO
'If LK_CODTRA = 2409 And Right(i_cias.text, 2) = "00" Then
'  If par_llave!par_flag_COSTOS <> 9 Then
'    MsgBox "!!! Compañia ... debe Costear Productos antes de Enviar ", 48, Pub_Titulo
'    Exit Sub
'  End If
'End If





pasa:


'If Tipo_Cambio.Visible = True Then
'   If WS_TIPO_CAMBIO <> LK_TIPO_CAMBIO And WS_TIPO_CAMBIO = 0 Then Tipo_Cambio_Click
'End If

If PUB_TIPMOV <> 0 And PUB_TIPMOV = 20 Then   'QUITAR EL TIPMOV..PARA HER
   If Val(i_flete.Text) <> SUB_FLETE Then
      SUB_FLETE = 0
      MsgBox " Los Fletes no coinciden ....Revisar " & SUB_FLETE & " - " & i_flete.Text
      Exit Sub
    End If
    If Val(i_descto.Text) <> PUB_DESCTO Then
       MsgBox " Los Desctos no coinciden ....Revisar ", 48, Pub_Titulo
       Exit Sub
    End If
End If



grabar.Enabled = False
Screen.MousePointer = 11

pub_cadena = "SELECT * FROM CONTROLL"
Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)

'repite: ' quite

barra.Visible = True
'Barra
barra.max = 90
barra.Min = 62

barra.Value = 62



N = 62
Do Until Val(tra_llave(N)) = 0 Or N = 76 Or exito = False


NUMERO = tra_llave(N)
On NUMERO GoSub CON1, CON2, CON3, CON4, CON5, CON6, CON7, CON8, CON9, CON10, CON11, CON12, CON13, CON14, CON15, CON16, CON17, CON18, CON19, CON20, CON21, CON22, CON23
barra.Value = N

N = N + 1
Loop
If exito = False Then
   con_llave.Close
   barra.Visible = False
   Screen.MousePointer = 0
   MsgBox pub_mensaje_err, 48, Pub_Titulo
   PUB_TIPDOC = Nulo_Valors(SUT_LLAVE!SUT_tipdoc)
   grabar.Enabled = True
   grabar.SetFocus
   GoTo errorr
End If
FLAG_2406 = "X"
If LK_CODTRA = 2406 Then FLAG_2406 = "A"
CN.Execute "Begin Transaction", rdExecDirect
repite:
GoSub manda_numero


N = 77
Do Until Val(tra_llave(N)) = 0 Or N = 91
NUMERO = tra_llave(N)
 On NUMERO GoSub ACT1, ACT2, ACT3, ACT4, ACT5, ACT6, ACT7, ACT8, ACT9, ACT10, ACT11, ACT12, ACT13, ACT14, ACT15, ACT16, ACT17, ACT18, ACT19, ACT20, ACT21
barra.Value = N
N = N + 1
Loop
barra.Value = 90




' PARA PIDA PROCESO DE COSTEO
If (LK_CODTRA = 2407 And LK_EMP_PTO = "A") Or PUB_TIPMOV = 20 Then
  SQ_OPER = 1
  PUB_CODCIA = LK_CODCIA
  LEER_PAR_LLAVE
  par_llave.Edit
  par_llave!par_flag_costos = "8"
  par_llave.Update
End If

CN.Execute "Commit Transaction", rdExecDirect
con_llave.Close


' <<<< PROCEDIMIENTO PARA LOS REPORTES >>>>

'dispositivo = REP_TRANSAC




i_cambio.Value = 0
N = 103
Do Until Val(tra_llave(N)) = 0 Or N = 107
NUMERO = tra_llave(103)
 On NUMERO GoSub REP1, REP2, REP3

N = N + 1
Loop

'i_numfac.Locked = True
'i_numser.Locked = True
i_num_lote.Locked = True
i_ser_lote.Locked = True
i_numfac.Text = ""
i_num_lote.Text = ""
i_ser_lote.Text = ""
i_importe.Text = ""
i_importe_amort.Text = ""
cancela_todito
pasa_cabeza
grid_fac.Rows = 3
FILAX = 0
barra.Visible = False
grabar.Enabled = True
label_nomart.Caption = ""
label_precio.Caption = ""
If LK_CODTRA = 2105 Or LK_CODTRA = 2107 Then i_fbg_Click

' AQUI INICIALIZO VARIABLES QUE NO SE CEREAN  AL ENTRAR AL COMMAND1.CLICK


WS_BRUTO = 0
SUB_CANT = 0
SUB_FLETE = 0
SUB_JABAS = 0
SUB_UNIDAD = 0
'WS_MONEDA_CLI = ""
WS_MONEDA_CCM = ""
fila = 0
grabar.Enabled = True
i_placas.Clear
i_concepto.Text = ""
If gridl.Visible Then pasa_cabeza_L
If LK_CODTRA = 2103 Then
   GridT.Clear
   GridD.Clear
   Option7.Value = True
   i_def.SetFocus
   flag_textil = 0
   GoTo fin
End If

If LK_CODTRA = 2105 Or LK_CODTRA = 2107 Then
   ' quite acv
   i_codcli.SetFocus
   GoTo fin
End If
If i_fbg.Visible = True Then  ' OJO 2125 AGREGALO
If i_fbg.ListCount = 1 Then
   If i_cambio.Value = 0 Then
      i_numfac.SetFocus
   Else
      i_numser.SetFocus
   End If
Else
   i_fbg.SetFocus
End If
End If

GoTo fin

CON1:
exito = ANEXO_1()
Return

CON2:
exito = ANEXO_2()
Return


CON3:
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_VEN_LLAVE
If ven_llave.EOF Then
   pub_mensaje_err = "Isla no existe... "
   exito = False
End If

If Not cli_llave.EOF And Nulo_Valors(par_llave!par_flag_cred) <> "A" And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 0 Then
   If SUT_LLAVE!SUT_SIGNO_CAR = 1 Then
      pu_codcia = LK_CODCIA
      pub_deuda = CAR_TOT_CPX2("C", pu_codcia, cli_llave!cli_codclie)
      If PUB_FLAG_VENCIDO = 1 And LK_FLAG_LIMITE <> "A" And LK_FLAG_LIMITE <> "C" Then
         pub_mensaje_err = "CLIENTE TIENE OBLIGACIONES VENCIDAS ... "
         exito = False
      End If
      If LK_FLAG_EXED = "A" Then
      If (pub_deuda + Val(i_neto.Text)) > Nulo_Valor0(cli_llave!cli_limcre) And LK_FLAG_LIMITE <> "B" And LK_FLAG_LIMITE <> "C" Then
         pub_mensaje_err = "LIMITE DE CREDITO EXCEDIDO ...SALDO POR ATENDER : " & Val(Nulo_Valor0(cli_llave!cli_limcre) - (pub_deuda))
         exito = False
      End If
      End If
   End If
End If

If Nulo_Valors(SUT_LLAVE!SUT_CERO) <> "A" Then
If WS_BRUTO = 0 Then
   pub_mensaje_err = "Importe errado... "
   exito = False
End If
End If
Return

CON4:

Return

CON5:
Return

CON6:
If LK_CODTRA <> 1402 Then Return
'  CAMTEX Or gridC.TextMatrix(2, 1) = ""

PU_NUMFAC = Val(gridC.TextMatrix(gridC.Row, 7))
PUB_NUMFAC = Val(gridC.TextMatrix(gridC.Row, 7))
pu_codcia = LK_CODCIA
PU_NUMSER = Val(gridC.TextMatrix(gridC.Row, 6))
PUB_FBG = ""
PU_TIPMOV = PUB_TIPMOV
PU_FBG = PUB_FBG

SQ_OPER = 1
LEER_FAR_LLAVE
If far_llave.EOF = True Then
   pub_mensaje_err = "Primero llame al ingreso Original "
   exito = False
   Return
End If
PUB_NUMSER = Val(i_numser.Text)
PUB_NUMFAC = Val(i_numfac.Text)

Return

CON7:
   
Return
   
CON8:
Return

CON9:
Return

CON10:
Return

CON11:

Return

CON12:
If i_cambio.Value = 1 And i_cambio.Visible = True Then Return


If Val(i_numguia.Text) <> 0 And PUB_TIPMOV = 10 Then
SQ_OPER = 6
PUB_SERGUIA = Val(i_serguia.Text)
PUB_NUMGUIA = Val(i_numguia.Text)
pu_codcia = LK_CODCIA
LEER_FAR_LLAVE
If far_guia.EOF = False Then
   pub_mensaje_err = "Guia ya existe en Doc: " & far_guia!FAR_FBG & " - " & far_guia!FAR_NUMFAC
   exito = False
   Return
End If
End If

If LK_CODTRA <> 2105 And LK_CODTRA <> 2107 Then
                                                        ' CAMTEX
If (i_cambio.Visible = False Or i_cambio.Value = 0) And LK_CODTRA <> 2211 Then Return
End If

SQ_OPER = 1
PU_TIPMOV = PUB_TIPMOV
PU_NUMSER = Val(i_numser.Text)
PU_NUMFAC = Val(i_numfac.Text)
' CAMTEX
If LK_CODTRA = 2211 Then
 PU_NUMSER = Val(i_ser_lote.Text)
 PU_NUMFAC = Val(i_num_lote.Text)
End If
pu_codcia = LK_CODCIA
PU_FBG = PUB_FBG
LEER_FAR_LLAVE

If far_llave.EOF = False Then
   Do Until far_llave.EOF
      WS_NUMSEC = far_llave!FAR_NUMSEC
                                        ' CAMTEX
      If far_llave!far_estado <> "E" And LK_CODTRA <> 2211 Then
         pub_mensaje_err = "Documento no esta extornado..."
         exito = False
      End If
      far_llave.MoveNext
   Loop
   WS_NUMSEC = WS_NUMSEC + 200
End If

Return

CON13:

If LK_CODTRA = 1133 Then PUB_FBG = "N"

SQ_OPER = 1
PU_TIPMOV = PUB_TIPMOV
PU_NUMSER = PUB_NUMSER_C
PU_NUMFAC = PUB_NUMFAC_C
pu_codcia = LK_CODCIA
PU_FBG = PUB_FBG
LEER_FAR_LLAVE
If far_llave.EOF = True Then
   pub_mensaje_err = "No existe Documento ..."
   LKCHEK = True
   exito = False
   Return
End If

If LK_CODTRA <> 1122 And LK_CODTRA <> 1133 Then Return
fx = 0
Do Until fx = 99
   If far_llave!far_estado <> "E" Then fx = 99
   If fx <> 99 Then
      far_llave.MoveNext
   End If
   If far_llave.EOF Then
      pub_mensaje_err = "DOCUMENTO YA  ANULADO ..."
      exito = False
      Return
   End If
Loop

PUB_CODCIA = far_llave!FAR_CODCIA
PUB_CODCIAL = far_llave!FAR_CODCIA
PUB_CODCLIE = far_llave!far_codclie
WS_IMPORTE_AMORT = Val(far_llave!far_bruto) + Val(far_llave!far_IMPTO) - Val(far_llave!FAR_TOT_DESCTO)
PUB_IMPORTE_AMORT = WS_IMPORTE_AMORT
PUB_CP = far_llave!FAR_CP
PUB_TIPMOV = far_llave!FAR_TIPMOV
PUB_CODBAN = 0
PUB_CHENUM = 0
PUB_CHESER = 0
PUB_CHESEC = 0
PUB_SERDOC = 0
PUB_NUMDOC = 0
PUB_CANTIDAD = 0
PUB_NUMSER = far_llave!far_numser_c
PUB_FBG = far_llave!FAR_FBG
PUB_CODVEN = far_llave!FAR_codven
PUB_NUMFAC = far_llave!far_numfac_c
i_numfac.Text = far_llave!FAR_NUMFAC

PUB_SECUENCIA = far_llave!FAR_NUMSEC
PUB_NUM_INI = 0
pub_signo_car = Nulo_Valor0(far_llave!far_signo_car) * -1
If pub_signo_car <> 0 Then PUB_TIPDOC = "FA"

pub_signo_arm = 1
pub_signo_caja = 0
PUB_CONCEPTO = "Anulado - " & far_llave!far_concepto


If pub_signo_car <> 0 Then
    SQ_OPER = 2
    pu_codcia = LK_CODCIA
    pu_codclie = PUB_CODCLIE
    pu_cp = PUB_CP
    fx = 0
    LEER_CAR_LLAVE
    Do Until car_mayor.EOF Or fx = 99
    'revisar cambie pub_numfac por pu_numfac para anul de ventas
        If PU_NUMFAC = Val(car_mayor!CAR_numfac) And PU_NUMSER = Val(car_mayor!CAR_numSER) Then
           fx = 99
        Else
           car_mayor.MoveNext
        End If
    Loop
    If Not car_mayor.EOF Then
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    pu_codclie = PUB_CODCLIE
    PUB_SERDOC = car_mayor!car_SERDOC
    PUB_NUMDOC = car_mayor!car_NUMDOC
    PUB_TIPDOC = car_mayor!CAR_TIPDOC
    pu_cp = PUB_CP
    LEER_CAR_LLAVE
'no importa si cancelo...
'    If car_llave!CAR_IMPORTE <> car_llave!CAR_IMP_INI Then
'       pub_mensaje_err = "DOCUMENTO NO PUEDE SER ANULADO...SALDO DIFERENTE "
'       exito = False
'       Return
'    End If
    
    End If
    
    
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    pu_codclie = PUB_CODCLIE
    pu_cp = PUB_CP
    PUB_FECHA = far_llave!FAR_fecha
    PUB_NUM_OPER = far_llave!FAR_NUMOPER
    LEER_CAA_LLAVE
    If caa_histo.EOF = True Then
       pub_mensaje_err = "DOCUMENTO EN CARACU NO EXISTE..."
       exito = False
    End If


End If

SQ_OPER = 5
PUB_FECHA = far_llave!FAR_fecha
PUB_NUM_OPER_EXT = far_llave!FAR_NUMOPER
LEER_FAR_LLAVE
'Print far_menor3!far_codclie
If far_menor3.EOF = True Then
   pub_mensaje_err = "!!! NO HAY ESE DOCUMENTO o YA ANULADO..."
   exito = False
   Return
End If

If LK_CODTRA = 1122 Then
If far_menor3!FAR_fecha = LK_FECHA_DIA Then
   pub_mensaje_err = "!!! Use el extorno de Operaciones ... "
   exito = False
   Return
End If
End If
fila = 2
Do Until far_menor3.EOF

   grid_fac.Rows = fila + 2
   WS_NUMSEC = far_menor3!FAR_NUMSEC
   
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_KEY = far_menor3!far_codart
   LEER_ART_LLAVE
   If art_LLAVE.EOF Then Exit Do
   grid_fac.TextMatrix(fila, 0) = art_LLAVE!ART_NOMBRE
   
   grid_fac.TextMatrix(fila, 16) = far_menor3!far_codart
   grid_fac.TextMatrix(fila, 33) = far_menor3!far_codart
   grid_fac.TextMatrix(fila, 2) = far_menor3!far_JABAS
   grid_fac.TextMatrix(fila, 3) = far_menor3!far_UNIDADES
   grid_fac.TextMatrix(fila, 11) = far_menor3!FAR_COSPRO
   grid_fac.TextMatrix(fila, 4) = far_menor3!FAR_CANTIDAD / far_menor3!FAR_equiv
   grid_fac.TextMatrix(fila, 6) = far_menor3!FAR_PRECIO
   grid_fac.TextMatrix(fila, 14) = far_menor3!FAR_equiv
   grid_fac.TextMatrix(fila, 5) = far_menor3!FAR_descri
   grid_fac.TextMatrix(fila, 8) = far_menor3!FAR_FLETE
   grid_fac.TextMatrix(fila, 10) = far_menor3!FAR_DESCTO
   grid_fac.TextMatrix(fila, 11) = far_menor3!FAR_COSPRO
   grid_fac.TextMatrix(fila, 12) = far_menor3!far_SIGNO_aRM * -1
   
   far_menor3.MoveNext
   fila = fila + 1
   
Loop


WS_NUMSEC = WS_NUMSEC + 200
calcula_totales



Return

CON14:
'COMPRAS

SQ_OPER = 1
pu_codcia = LK_CODCIA
pu_codclie = PUB_CODCLIE
pu_cp = PUB_CP
LEER_CAR_LLAVE
If car_llave.EOF = False Then
   pub_mensaje_err = "CARTERA YA EXISTE..."
   exito = False
End If
Return

CON15:

   
Return

CON16:
SQ_OPER = 1
LEER_TAB_LLAVE



Return

CON17:
If PUB_FBG = "B" Or PUB_FBG = "F" Then
   If PUB_IMPTO = 0 Then
       pub_mensaje_err = "Falta el Igv.."
       exito = False
       Return
   End If
End If

Return

CON18:

SQ_OPER = 1
pu_cp = PUB_CP
pu_codclie = PUB_CODCLIE
pu_codcia = LK_CODCIA
LEER_CAR_LLAVE
If car_llave.EOF = True Then
   pub_mensaje_err = "DOCUMENTO EN CARTERA NO EXISTE..."
   exito = False
   Return
End If




Return

CON19:
If LK_CODTRA <> 2407 Then Return
If i_cias.ListIndex = -1 Then
   pub_mensaje_err = "Seleccion Destino ..."
   exito = False
End If


Return

CON20:
CONTADOR = 0
Do Until CONTADOR = PUB_CANT_CHEQ
PUB_CHESEC = 0
PUB_CHENUM = CONTADOR + PUB_NUM_INI
CONTADOR = CONTADOR + 1
SQ_OPER = 1
LEER_CHE_LLAVE
If Not che_llave.EOF Then
       pub_mensaje_err = "Numero de Serie y Cheque ya existe..." & PUB_CHESER & "-" & PUB_CHENUM
       exito = False
       Return
End If
Loop
Return

CON21:
PUB_FECHA = LK_FECHA_DIA
If LK_CODTRA = 1122 Then
   SQ_OPER = 1
   PU_TIPMOV = PUB_TIPMOV
   PU_NUMSER = PUB_NUMSER_C
   PU_NUMFAC = PUB_NUMFAC_C
   pu_codcia = LK_CODCIA
   PU_FBG = PUB_FBG
   LEER_FAR_LLAVE
   If far_llave.EOF = True Then
      pub_mensaje_err = "No existe Documento ..."
      LKCHEK = True
      exito = False
      Return
   End If
   Do Until far_llave.EOF
    If far_llave!far_estado = "N" Then GoTo ESTA_BIEN
    far_llave.MoveNext
   Loop
   pub_mensaje_err = "Ya Extornado  ..."
   LKCHEK = True
   exito = False
   Return
ESTA_BIEN:
   PUB_FECHA = far_llave!FAR_fecha
   PUB_NUM_OPER = far_llave!FAR_NUMOPER
End If

SQ_OPER = 1
pu_codcia = LK_CODCIA
LEER_ALL_LLAVE
Do Until all_llave.EOF
  If PUB_NUM_OPER = all_llave!ALL_NUMOPER Then
    Exit Do
  Else
    all_llave.MoveNext
  End If
Loop
     

If all_llave.EOF Then
   pub_mensaje_err = "Numero de Operacion Incorrecto...."
   exito = False
   Return
End If

If all_llave!all_flag_ext = "E" Then
   pub_mensaje_err = "Operacion ya Extornada........"
   exito = False
   Return
End If
If all_llave!all_flag_ext = "X" Then
   pub_mensaje_err = "Operacion NO Extornable........"
   exito = False
   Return
End If


If all_llave!all_codCIA <> LK_CODCIA Then
   pub_mensaje_err = "Cia. no coincide..."
   exito = False
   Return
End If




PUB_NUM_OPER = all_llave!ALL_NUMOPER
PUB_NUM_OPER_EXT = all_llave!ALL_NUMOPER
PUB_CODCIA = all_llave!all_codCIA
PUB_CODCIAL = all_llave!all_codCIA
PUB_CODCLIE = all_llave!all_CODCLIE
ww_codtra_ext = all_llave!ALL_CODTRA

PUB_NUMDOC = Nulo_Valor0(all_llave!ALL_NUMDOC)
PUB_CODART = all_llave!ALL_CODART
PUB_IMPORTE_AMORT = all_llave!ALL_IMPORTE_AMORT
WS_IMPORTE_AMORT = all_llave!ALL_IMPORTE_AMORT

If all_llave!all_tipmov = 20 Then
   If Val(all_llave!ALL_IMPORTE_DOLL) <> 0 Then WS_IMPORTE_AMORT = all_llave!ALL_IMPORTE_DOLL
End If

PUB_IMPORTE = all_llave!ALL_importe
PUB_CP = all_llave!all_cp
PUB_TIPDOC = all_llave!ALL_TIPDOC
PUB_TIPMOV = all_llave!all_tipmov
PUB_CODBAN = all_llave!ALL_CODBAN
PUB_CHENUM = all_llave!ALL_CHENUM
PUB_CHESER = Nulo_Valor0(all_llave!ALL_CHESER)
PUB_CHESEC = Nulo_Valor0(all_llave!ALL_CHESEC)
PUB_SERDOC = Nulo_Valor0(all_llave!ALL_serdoc)
PUB_NUMDOC = Nulo_Valor0(all_llave!ALL_NUMDOC)
ws_fecha_anterior = Nulo_Valor0(all_llave!all_fecha_ant)
PUB_CANTIDAD = all_llave!ALL_CANTIDAD
PUB_NUMSER = all_llave!all_numser
PUB_FBG = all_llave!ALL_FBG
PUB_CODVEN = all_llave!ALL_CODVEN
PUB_NUMFAC = all_llave!all_numfac
PUB_SECUENCIA = all_llave!all_SECUENCIA
pub_signo_ccm = Nulo_Valor0(all_llave!ALL_SIGNO_CCM) * -1
pub_signo_car = Nulo_Valor0(all_llave!ALL_SIGNO_CAR) * -1
pub_signo_arm = Nulo_Valor0(all_llave!all_sIGNO_ARM) * -1
pub_signo_caja = Nulo_Valor0(all_llave!ALL_SIGNO_CAJA) * -1
WS_FLAG_FACART = Nulo_Valors(all_llave!ALL_FACART)
PUB_CONCEPTO = "Extorno - " & all_llave!ALL_CONCEPTO


If pub_signo_ccm <> 0 Then
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_CCM_LLAVE
   If ccm_llave.EOF Then
      pub_mensaje_err = "banco no existe..."
      exito = False
   End If
End If



If pub_signo_car <> 0 Then
SQ_OPER = 1
pu_codcia = LK_CODCIA
pu_codclie = PUB_CODCLIE
pu_cp = PUB_CP
LEER_CAR_LLAVE
If car_llave.EOF = True Then
   pub_mensaje_err = "DOCUMENTO NO EXISTE..."
   exito = False
End If
SQ_OPER = 1
pu_codcia = LK_CODCIA
pu_codclie = PUB_CODCLIE
pu_cp = PUB_CP
LEER_CAA_LLAVE
If caa_histo.EOF = True Then
   pub_mensaje_err = "DOCUMENTO EN CARACU NO EXISTE..."
   exito = False
End If
End If
If LK_CODTRA = 1122 Then
If WS_IMPORTE_AMORT <> Val(i_importe.Text) Then
   pub_mensaje_err = "!!! IMPORTE NO CORRESPONDE ..."
   exito = False
   Return
End If
End If


If PUB_TIPMOV = 0 And Trim(WS_FLAG_FACART) = "" Then
   Return
End If
'If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then Return
SQ_OPER = 5
LEER_FAR_LLAVE
' en caso de notas de credito, puede ser sin facart
'If far_menor3.EOF = True Then
'   MsgBox "!!! NO HAY ESE DOCUMENTO ...", 48, Pub_Titulo
'   exito = False
'   Return
'End If
If far_menor3.EOF Then Return



If far_menor3!far_codart = 0 Then
       Return
End If


fila = 2
Do Until far_menor3.EOF

   grid_fac.Rows = fila + 2
   WS_NUMSEC = far_menor3!FAR_NUMSEC
   
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_KEY = far_menor3!far_codart
   LEER_ART_LLAVE
   'SOLO PARA LA 2499
   If PUB_TIPMOV <> 80 Then grid_fac.TextMatrix(fila, 0) = art_LLAVE!ART_NOMBRE
   
   grid_fac.TextMatrix(fila, 16) = far_menor3!far_codart
   grid_fac.TextMatrix(fila, 33) = far_menor3!far_codart
   grid_fac.TextMatrix(fila, 2) = far_menor3!far_JABAS
   grid_fac.TextMatrix(fila, 3) = far_menor3!far_UNIDADES
   grid_fac.TextMatrix(fila, 11) = far_menor3!FAR_COSPRO
   If far_menor3!FAR_equiv > 0 Then
      grid_fac.TextMatrix(fila, 4) = far_menor3!FAR_CANTIDAD / far_menor3!FAR_equiv
   Else
      grid_fac.TextMatrix(fila, 4) = far_menor3!FAR_CANTIDAD
   End If
   grid_fac.TextMatrix(fila, 6) = far_menor3!FAR_PRECIO
   grid_fac.TextMatrix(fila, 14) = Nulo_Valor0(far_menor3!FAR_equiv)
   grid_fac.TextMatrix(fila, 5) = Nulo_Valors(far_menor3!FAR_descri)
   grid_fac.TextMatrix(fila, 8) = Nulo_Valor0(far_menor3!FAR_FLETE)
   grid_fac.TextMatrix(fila, 10) = Nulo_Valor0(far_menor3!FAR_DESCTO)
   grid_fac.TextMatrix(fila, 11) = far_menor3!FAR_COSPRO
   grid_fac.TextMatrix(fila, 14) = far_menor3!FAR_equiv
   grid_fac.TextMatrix(fila, 18) = Nulo_Valor0(far_menor3!far_pedsec)
   grid_fac.TextMatrix(fila, 19) = Nulo_Valor0(far_menor3!FAR_orden_unidades)
   grid_fac.TextMatrix(fila, 12) = far_menor3!far_SIGNO_aRM * -1
   WS_MONEDA_CLI = Nulo_Valors(far_menor3!far_MONEDA)
   If WS_MONEDA_CLI = "S" Then
      i_ds.ListIndex = 0
   ElseIf WS_MONEDA_CLI = "D" Then
      i_ds.ListIndex = 1
   Else
      i_ds.ListIndex = -1
   End If
   
   i_moneda.Text = Nulo_Valors(far_menor3!far_MONEDA)
   lmoneda.Caption = Nulo_Valors(far_menor3!far_MONEDA)
   PUB_PEDSER = Nulo_Valor0(far_menor3!FAR_PEDSER)
   PUB_PEDFAC = Nulo_Valor0(far_menor3!FAR_PEDFAC)
   far_menor3.MoveNext
   fila = fila + 1
   
Loop
WS_NUMSEC = WS_NUMSEC + 200
calcula_totales

PUB_FECHA = LK_FECHA_DIA
Return
CON22:
SQ_OPER = 1
pu_codcia = LK_CODCIA
pu_cp = PUB_CP
pu_codclie = PUB_CODCLIE
LEER_CAR_LLAVE
If car_llave.EOF = True Then
   pub_mensaje_err = "DOCUMENTO NO EXISTE..."
   exito = False
   Return
End If

If car_llave!CAR_FECHA_INGR <> LK_FECHA_DIA Then
   pub_mensaje_err = "NO PUEDE MODIFICAR DOCS... DE OTRO DIA..."
   exito = False
   Return
End If
   
Return

CON23:
Return


GoTo fin

ACT1:

GoSub ACT2

flag_salto = 1


If (PUB_TIPMOV = 0 Or pub_signo_arm = 0) And LK_CODTRA <> 2406 Then
    Return
End If
    
If LK_CODTRA = 1111 Or LK_CODTRA = 1122 Then
   If PUB_TIPMOV = 10 And PUB_IMPORTE_AMORT = 0 And Val(grid_fac.TextMatrix(2, 4)) = 0 Then Return
End If

If Val(i_neto.Text) = 0 And Trim(grid_fac.TextMatrix(2, 4)) = "ANULACION" Then
   GoTo ACT13
End If


'If i_cambio.Visible = False Or i_cambio.Value = 0 Then llena_numfac_grifo


SUB_CANT = 0
fila = 2
WS_BRUTO2 = 0
FLAG = False
Do While FLAG = False
   If Val(grid_fac.TextMatrix(fila, 16)) = 0 Or (Val(grid_fac.TextMatrix(fila, 3)) = 0 And Val(grid_fac.TextMatrix(fila, 4)) = 0) Then
      GoTo OTROMAS
   End If
   
   PUB_CODART = Val(grid_fac.TextMatrix(fila, 16))
   
   PUB_JABAS = Val(grid_fac.TextMatrix(fila, 2))
   PUB_UNIDAD = Val(grid_fac.TextMatrix(fila, 2)) 'PARA LA TEXTIL
   
   PUB_PRECIO2 = Val(grid_fac.TextMatrix(fila, 6))
   
   PUB_COSPRO = Val(grid_fac.TextMatrix(fila, 11))
   
sigue:
   WS_CANTIDAD = Val(grid_fac.TextMatrix(fila, 4)) * Val(grid_fac.TextMatrix(fila, 14))

   If pub_signo_arm = 0 And LK_CODTRA <> 2212 Then
      GoTo arm_no
   End If
   
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ARM_LLAVE
   If arm_llave.EOF Then
     CN.Execute "Rollback Transaction", rdExecDirect
     MsgBox "error grave en ARTICULO ..."
     GoTo fin
   End If
   

If LK_CODTRA <> 1122 And LK_CODTRA <> 1111 Then
 If Nulo_Valors(SUT_LLAVE!SUT_UNIDADES) <> "A" And Val(grid_fac.TextMatrix(fila, 14)) <> 1 Then
   CN.Execute "Rollback Transaction", rdExecDirect
   MsgBox "error grave en ARTICULO ..."
   GoTo fin
 End If
End If
                                                
   If Trim(Nulo_Valors(SUT_LLAVE!sut_is)) = "A" Or LK_CODTRA = 2425 Or Check1.Value = 1 Then 'ok 2425 unico caso
      pub_signo_arm = Val(grid_fac.TextMatrix(fila, 12))
   End If
   If LK_CODTRA = 2406 And PUB_SECUENCIA = 0 And Trim(FLAG_2406) = "" Then pub_signo_arm = -1
   
      
   
   
   If (pub_signo_arm = -1 Or pub_signo_arm = 1) Then
   Else
      CN.Execute "Rollback Transaction", rdExecDirect
      MsgBox "error grave en SIGNO ARM ..."
      GoTo fin
   End If
  
      If LK_CODTRA <> 2105 And LK_CODTRA <> 2107 And LK_CODTRA <> 1402 Then
         CONTADOR = CONTADOR + 1
         arm_llave.Edit
         arm_llave!ARM_STOCK = Val(arm_llave!ARM_STOCK) + WS_CANTIDAD * pub_signo_arm
         arm_llave!ARM_stock2 = Val(arm_llave!ARM_stock2) + PUB_JABAS * pub_signo_arm
         
         If PUB_SO = "A" Then
            arm_llave!ARM_saldo_s = Val(arm_llave!ARM_saldo_s) + WS_CANTIDAD * pub_signo_arm
            arm_llave!ARM_SALDO_S2 = Val(arm_llave!ARM_SALDO_S2) + PUB_JABAS * pub_signo_arm
         Else
            arm_llave!ARM_Saldo_n = Val(arm_llave!ARM_Saldo_n) + WS_CANTIDAD * pub_signo_arm
            arm_llave!ARM_SALDO_N2 = Val(arm_llave!ARM_SALDO_N2) + PUB_JABAS * pub_signo_arm
         End If
         If pub_signo_arm = -1 Then
            arm_llave!ARM_SALIDAS = Val(arm_llave!ARM_SALIDAS) + Val(WS_CANTIDAD)
         ElseIf pub_signo_arm = 1 Then
            arm_llave!ARM_INGRESOS = Val(arm_llave!ARM_INGRESOS) + Val(WS_CANTIDAD)
         End If
         arm_llave.Update
      End If
      
         pu_codcia = LK_CODCIA
         PUB_KEY = arm_llave!ARM_CODART
         LEER_ART_LLAVE
         If art_LLAVE.EOF Then
            CN.Execute "Rollback Transaction", rdExecDirect
            MsgBox "error grave en ARTICULO ..."
            GoTo fin
         End If
         
      If Trim(art_LLAVE!ART_NOMBRE) = Trim(grid_fac.TextMatrix(fila, 0)) Then
      Else
         CN.Execute "Rollback Transaction", rdExecDirect
         MsgBox "error grave con codigos alternos ..."
         GoTo fin
      End If
      
      If LK_CODTRA = 2403 And PUB_SECUENCIA = 10 Then
         arm_llave.Edit
         arm_llave!arm_cospro = PUB_PRECIO2
         PUB_COSPRO = PUB_PRECIO2
         arm_llave.Update
      End If
         
         
      
      ws_flag_arm = MODIF
      
arm_no:
      far_llave.AddNew
      far_llave!FAR_TIPMOV = PUB_TIPMOV
      far_llave!FAR_CODCIA = LK_CODCIA
      far_llave!far_COD_SUNAT = 0 'Val(Right(i_codsunat.Text, 5))
      far_llave!far_numser = PUB_NUMSER
      ' quite CAMTEX
      If PUB_TIPMOV = 181 Then
       far_llave!FAR_codven = Val(Left(grid_fac.TextMatrix(1, 0), 2))
      Else
        far_llave!FAR_codven = PUB_CODVEN
      End If
      
      far_llave!FAR_NUMFAC = PUB_NUMFAC
      ' CAMTEX
      If LK_CODTRA = 2103 Or LK_CODTRA = 2211 Then
         far_llave!far_numser = Val(i_ser_lote.Text)
         far_llave!FAR_NUMFAC = Val(i_num_lote.Text)
      End If
      WS_NUMSEC = WS_NUMSEC + 1
      far_llave!FAR_NUMSEC = WS_NUMSEC
      If pub_signo_arm = 0 Then
         far_llave!far_codart = PUB_CODART
         far_llave!FAR_STOCK = 0
      Else
         far_llave!FAR_STOCK = Val(arm_llave!ARM_STOCK)
         far_llave!far_codart = Val(arm_llave!ARM_CODART)
      End If
      far_llave!FAR_CANTIDAD = Val(WS_CANTIDAD)
      far_llave!far_PESO = Val(grid_fac.TextMatrix(fila, 4))
      
      SUB_CANT = SUB_CANT + WS_CANTIDAD
      
      far_llave!far_signo_car = pub_signo_car
      
      If LK_CODTRA = 2101 And PUB_SECUENCIA = 15 Then
         far_llave!far_signo_car = 0
      End If

      far_llave!far_SIGNO_aRM = pub_signo_arm
      
      far_llave!far_codclie = PUB_CODCLIE
      WS_MONEDA_CLI = i_ds.Text
      
      far_llave!far_MONEDA = WS_MONEDA_CLI
      far_llave!FAR_EX_IGV = Nulo_Valors(grid_fac.TextMatrix(fila, 23))
      
      far_llave!FAR_CP = PUB_CP
      If i_fecha_compra.Visible = True Then
         far_llave!far_FECHA_COMPRA = i_fecha_compra.Text
      Else
         far_llave!far_FECHA_COMPRA = LK_FECHA_DIA
      End If
      
      far_llave!far_estado = "N"
      far_llave!FAR_estado2 = "N"
      
      If LK_CODTRA = 1111 Or LK_CODTRA = 1122 Or LK_CODTRA = 1133 Then
         far_llave!far_estado = "E"
      End If
      
      If LK_CODTRA = 2105 Or LK_CODTRA = 2107 Or LK_CODTRA = 1402 Then
         far_llave!far_estado = "M"
         far_llave!FAR_estado2 = "M"
      End If
      
      
      
      far_llave!FAR_COSPRO = PUB_COSPRO
      
      'YO AGREGO
      If LK_CODTRA = 2407 Then
      far_llave!FAR_COSPRO_ANT = Val(arm_llave!arm_cospro)
      Else
      far_llave!FAR_COSPRO_ANT = PUB_COSPRO
      End If
      far_llave!FAR_PRECIO = PUB_PRECIO2
      
      far_llave!FAR_equiv = Val(Nulo_Valor0(grid_fac.TextMatrix(fila, 14)))
      far_llave!FAR_FBG = PUB_FBG
      far_llave!far_IMPTO = PUB_IMPTO
      far_llave!FAR_TOT_FLETE = PUB_FLETE
      far_llave!FAR_FLETE = Val(Nulo_Valor0(grid_fac.TextMatrix(fila, 8)))
'      If PUB_DESCTO <> 0 Then
         far_llave!FAR_DESCTO = Val(Nulo_Valor0(grid_fac.TextMatrix(fila, 10)))
'      Else
'        far_llave!FAR_DESCTO = 0
'      End If
       If LK_CODTRA = 2210 Or LK_CODTRA = 2211 Or LK_CODTRA = 2212 Then
         far_llave!FAR_DESCTO = 0
       End If
      far_llave!FAR_TOT_DESCTO = PUB_DESCTO
      far_llave!FAR_GASTOS = PUB_GASTOS
      If LK_CODTRA = 2211 And PUB_SECUENCIA = 1 Then
        far_llave!far_bruto = Val(Nulo_Valor0(grid_fac.TextMatrix(fila, 6)))
      Else
        far_llave!far_bruto = PUB_SUBTOTAL
      End If
      far_llave!FAR_NUMDOC = PUB_NUMDOC
      
      far_llave!far_NUMGUIA = Val(i_numguia.Text)
      far_llave!far_serguia = Val(i_serguia.Text)
      
   
      far_llave!FAR_pordescto1 = 0 ' Val(Nulo_Valor0(grid_fac.TextMatrix(fila, 10)))
      
      far_llave!FAR_COSTEO = " "
      far_llave!FAR_COSTEO_REAL = " "
      
      far_llave!FAR_tipo_cambio = 1
      
      far_llave!far_DIAS = pub_dias
      
      far_llave!FAR_fecha = LK_FECHA_DIA
      
      far_llave!far_numser_c = PUB_NUMSER_C
      far_llave!far_numfac_c = PUB_NUMFAC_C
      If LK_CODTRA = 2211 Then
         far_llave!far_numfac_c = Val(grid_fac.TextMatrix(fila, 39))
      End If
      far_llave!FAR_NUMOPER = PUB_NUM_OPER_XXX
      far_llave!far_NUMOPER2 = PUB_NUM_OPER_XXX
      
      'QUITE ACV =  far_llave!far_precio_neto = Val(grid_fac.TextMatrix(fila, 36))
      far_llave!far_precio_neto = Val(grid_fac.TextMatrix(fila, 35))
      
      'CAMPO ELIMINADO...
'      far_llave!far_limcre_act = Val(grid_fac.TextMatrix(fila, 35))
      
      
      
      far_llave!FAR_CONSIG = 0
      
      far_llave!far_otra_cia = " "
      far_llave!far_subtra = SUT_LLAVE!SUT_DESCRIPCION
      
      If i_cias.Visible = True Then
         far_llave!far_transito = "T"
         far_llave!far_subtra = Left(SUT_LLAVE!SUT_DESCRIPCION, 7) & " A " & i_cias.Text
         far_llave!far_otra_cia = Right(i_cias.Text, 2)
      Else
         far_llave!far_transito = " "
      End If
      
      If LK_CODTRA = 2103 And PUB_SECUENCIA = 10 Then far_llave!far_transito = "P"
      
      If LK_CODTRA = 2406 And PUB_SECUENCIA = 0 And FLAG_2406 = "A" Then
         FLAG_2406 = ""
         far_llave!far_transito = "A"
         far_llave!far_subtra = ""
         far_llave!far_otra_cia = ""
      End If
      If Boton_Recepcion.Visible Then
         far_llave!far_subtra = Left(SUT_LLAVE!SUT_DESCRIPCION, 10) & " De " & grid_trans.TextMatrix(1, 3)
         far_llave!far_otra_cia = grid_trans.TextMatrix(grid_trans.Row, 6)
         far_llave!far_transito = "Y"
      End If
      
      far_llave!far_subtra = i_placa.Text
      If grid_fac.TextMatrix(fila, 1) <> "" And Trim(grid_fac.TextMatrix(fila, 15)) <> "" Then
         far_llave!far_subtra = grid_fac.TextMatrix(fila, 15)
      End If
      
      far_llave!FAR_descri = Nulo_Valors(grid_fac.TextMatrix(fila, 5))
      far_llave!far_JABAS = PUB_JABAS
      far_llave!far_UNIDADES = PUB_UNIDAD
      far_llave!far_mortal = Val(Nulo_Valor0(Mid(grid_fac.TextMatrix(fila, 9), 23, 5)))
      far_llave!far_num_precio = Val(grid_fac.TextMatrix(fila, 17))
      far_llave!FAR_orden_unidades = Val(grid_fac.TextMatrix(fila, 19))
      far_llave!FAR_SUBTOTAL = Val(grid_fac.TextMatrix(fila, 7))
      far_llave!far_ISLA = PUB_CODVEN
      far_llave!far_turno = Val(i_turno.Text)
      far_llave!far_concepto = " "
      If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
         far_llave!far_concepto = Nulo_Valors(i_TEXTONCRE.Text)
      End If
      If PUB_TIPMOV = 181 Then
         far_llave!far_concepto = Mid(Trim(i_def.Text), 4, Len(i_def.Text))
      End If
      far_llave!far_codusu = LK_CODUSU
'      far_llave!far_CODTRA = SUT_LLAVE!SUT_CODTRA
'      far_llave!far_SUBCODTRA = SUT_LLAVE!SUT_SECUENCIA
      far_llave!FAR_HORA = Format(Now, "hh:mm:ss AMPM")
      far_llave!far_NUM_LOTE = art_LLAVE!art_familia ' ES PARA LA TEXTIL
'      If i_codsunat.ListIndex >= 0 Then far_llave!far_COD_SUNAT = Val(Right(i_codsunat.Text, 5))
      If LK_CODTRA = 2103 Or LK_CODTRA = 2210 Or LK_CODTRA = 2212 Or LK_CODTRA = 2211 Then
         far_llave!FAR_PEDSER = Val(i_ser_lote.Text)
         far_llave!FAR_PEDFAC = Val(i_num_lote.Text)
         far_llave!far_pedsec = 0
         far_llave!far_fbg2 = "O"
      Else
         far_llave!FAR_PEDSER = Val(grid_fac.TextMatrix(fila, 28))
         far_llave!FAR_PEDFAC = Val(grid_fac.TextMatrix(fila, 30))
         far_llave!far_pedsec = Val(grid_fac.TextMatrix(fila, 31))
         far_llave!far_fbg2 = grid_fac.TextMatrix(fila, 29)
      End If
      
      
      far_llave!FAR_FLAG_SO = PUB_SO

      'far_llave!FAR_pedsec = Val(grid_fac.TextMatrix(fila, 18))
      far_llave!FAR_TIPDOC = PUB_TIPDOC
      If LK_CODTRA = 2101 And PUB_SECUENCIA = 5 Then far_llave!FAR_TIPDOC = "VO"
      If LK_CODTRA = 2101 And PUB_SECUENCIA = 25 Then far_llave!FAR_TIPDOC = "AU"
      
      
      
      far_llave.Update
      
      
      ws_flag_far = ingre
OTROMAS:

   fila = fila + 1
   If fila > grid_fac.Rows - 1 Then
      FLAG = True
   End If
Loop


'If pub_autkey <> 0 Then
'   aut_llave.MoveFirst
'   Do Until aut_llave.EOF
'         aut_llave.Edit
'         aut_llave!aut_estado = "A"
'         aut_llave.Update
'         aut_llave.MoveNext
'   Loop
'   ' SENDMAIL "alancosme@almacen", "AUTORIZACIONES", "PASE AUTORIZADO.... EJECUTADA"'
'End If
If Val(gridl.TextMatrix(1, 5)) = 0 And (PUB_TIPMOV = 97 Or PUB_TIPMOV = 98) Then
   GoSub ACT7
   Return
End If

flag_salto = 0
i_subtotal.Text = Val(i_subtotal.Text)
PUB_SUBTOTAL = i_subtotal.Text
Return

ACT2:
' agrege acv 22-07
If i_cambio.Value = 0 Or i_cambio.Visible = False Then Return

WW_FECHA_ANT = LK_FECHA_DIA
WW_NUM_OPER = PUB_NUM_OPER_XXX


far_llave.MoveFirst

If far_llave.EOF Then Return

If LK_CODTRA = 1401 Then
   far_llave.MoveLast
   If Val(i_codcli.Text) <> far_llave!far_codclie Then pub_flag_cambio = 1
   far_llave.MoveFirst
End If

If pub_flag_cambio = 0 And LK_CODTRA = 1401 Then
   Do Until far_llave.EOF
      far_llave.Edit
      far_llave!far_numfac_c = PUB_NUMFAC_C
      far_llave!far_numser_c = PUB_NUMSER_C
      far_llave!far_NUMGUIA = PUB_NUMGUIA
      If i_fecha_compra.Visible Then
        far_llave!far_FECHA_COMPRA = i_fecha_compra.Text
      Else
        far_llave!far_FECHA_COMPRA = LK_FECHA_DIA
      End If
      If LK_FLAG_GRIFO = "A" And LK_CODTRA = 1401 Then
        far_llave!far_turno = Val(i_turno.Text)
        far_llave!FAR_PEDFAC = Val(i_TEXTONCRE.Text)
      End If
      far_llave!far_DIAS = pub_dias
      far_llave!far_serguia = Val(i_serguia.Text)
      far_llave!far_NUMGUIA = PUB_NUMGUIA
      far_llave.Update
      PUB_NUMFAC = far_llave!FAR_NUMFAC
      PUB_NUMSER = far_llave!far_numser
      If far_llave!far_estado <> "E" Then PUB_NUM_OPER_EXT = far_llave!FAR_NUMOPER

      
      far_llave.MoveNext
   Loop
Else
   far_llave.MoveLast      ' PARA VER CUAL ES EL VALIDO...CUANDO HAY VARIAS ANULACIONES DEL MISMO DOC.
   WS_NUMSEC = far_llave!FAR_NUMSEC
   If far_llave!far_estado = "E" Then Return

   SQ_OPER = 5
   PUB_FECHA = far_llave!FAR_fecha
   WW_FECHA_ANT = far_llave!far_FECHA_COMPRA
   WW_NUM_OPER = far_llave!far_NUMOPER2
   PUB_NUM_OPER_EXT = far_llave!FAR_NUMOPER
   
   far_llave.MoveFirst ' ES PARA SU EXTORNO EN ACT2
   
   LEER_FAR_LLAVE
   If far_menor3.EOF = True Then
      CN.Execute "Rollback Transaction", rdExecDirect
      MsgBox "error grave en ARTICULO ..."
      GoTo fin
   End If
   
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_FECHA = PUB_FECHA
    LEER_ALL_LLAVE
    Do Until all_llave.EOF
      ' quite en arequipa
      ' far_llave!far_NUMOPER2
      If PUB_NUM_OPER_EXT = all_llave!ALL_NUMOPER Then
        Exit Do
      Else
        all_llave.MoveNext
      End If
    Loop
    If all_llave.EOF Then
          CN.Execute "Rollback Transaction", rdExecDirect
          MsgBox "error grave en ALLOG ..."
          GoTo fin
    End If
    all_llave.Edit
    all_llave!all_flag_ext = "E"
    all_llave.Update
 
'End If
 

   
Do Until far_llave.EOF
   If far_llave!far_estado = "E" Then GoTo SALTA_OTRO
   i = 0
   far_menor3.AddNew
   Do Until i = 81
      If Not IsNull(far_llave(i)) Then far_menor3(i) = far_llave(i)
      i = i + 1
   Loop
   WS_NUMSEC = WS_NUMSEC + 1
   far_menor3!FAR_NUMSEC = WS_NUMSEC
   
   far_menor3!far_estado = "E"
   
   far_menor3!far_SIGNO_aRM = far_llave!far_SIGNO_aRM * -1
   pub_signo_arm = far_menor3!far_SIGNO_aRM
   
   'AGREGE ACV
   'far_menor3!far_NUMOPER = PUB_NUM_OPER_XXX
   'Print PUB_NUM_OPER_XXX
   far_menor3.Update
   
   PUB_CODART = far_llave!far_codart
   far_llave.Edit
   far_llave!far_estado = "E"
   far_llave.Update
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ARM_LLAVE
   If arm_llave.EOF Then
     CN.Execute "Rollback Transaction", rdExecDirect
     MsgBox "error grave en ARTICULO ..."
     GoTo fin
   End If
   
   WS_CANTIDAD = far_llave!FAR_CANTIDAD
   PUB_JABAS = far_llave!far_JABAS
   
   'AGREGE ACV
   pub_dias = far_llave!far_DIAS
'   pub_signo_arm = far_llave!far_SIGNO_aRM * -1
   
   If pub_signo_arm = -1 Or pub_signo_arm = 1 Then
   Else
      CN.Execute "Rollback Transaction", rdExecDirect
      MsgBox "error grave en SIGNO ARM ..."
      GoTo fin
   End If
   If LK_CODTRA <> 2105 And LK_CODTRA <> 2107 And LK_CODTRA <> 1402 Then
        arm_llave.Edit
        arm_llave!ARM_STOCK = Val(arm_llave!ARM_STOCK) + WS_CANTIDAD * pub_signo_arm
        arm_llave!ARM_stock2 = Val(arm_llave!ARM_stock2) + PUB_JABAS * pub_signo_arm
        
        If PUB_SO = "A" Then
           arm_llave!ARM_saldo_s = Val(arm_llave!ARM_saldo_s) + WS_CANTIDAD * pub_signo_arm
           arm_llave!ARM_SALDO_S2 = Val(arm_llave!ARM_SALDO_S2) + PUB_JABAS * pub_signo_arm
        Else
           arm_llave!ARM_Saldo_n = Val(arm_llave!ARM_Saldo_n) + WS_CANTIDAD * pub_signo_arm
           arm_llave!ARM_SALDO_N2 = Val(arm_llave!ARM_SALDO_N2) + PUB_JABAS * pub_signo_arm
        End If

      
        If pub_signo_arm = -1 Then
           arm_llave!ARM_SALIDAS = Val(arm_llave!ARM_SALIDAS) + Val(WS_CANTIDAD)
        Else
        If pub_signo_arm = 1 Then
           arm_llave!ARM_INGRESOS = Val(arm_llave!ARM_INGRESOS) + Val(WS_CANTIDAD)
        End If
        End If
        If LK_CODTRA = 1401 Then
           If i_fecha_compra.Visible Then
               arm_llave!ARM_FECHA_ULT = i_fecha_compra.Text
           Else
               arm_llave!ARM_FECHA_ULT = LK_FECHA_DIA
           End If
        End If
        arm_llave.Update
     End If
       'AGREGE ACV
      If LK_CODTRA = 1401 And i_cambio.Value = 1 Then
         pub_signo_arm = 1
      End If
      
      pub_signo_car = far_llave!far_signo_car


   PUB_NUMFAC = far_llave!FAR_NUMFAC
   PUB_NUMSER = far_llave!far_numser
SALTA_OTRO:
   far_llave.MoveNext
  Loop
  
End If

If pub_signo_car = 0 Then GoTo ffin


If LK_CODTRA = 2414 Then GoTo ffin


SQ_OPER = 5
pu_codcia = LK_CODCIA
LEER_CAR_LLAVE
If car_far2.EOF = True Then
      pub_mensaje = "!!! ERROR EN ACTUALIZACION de CANJE..."
      exito = False
      GoTo Error_fatal
End If


SQ_OPER = 1
pu_cp = cli_llave!CLI_CP
pu_codclie = car_far2!CAR_CODCLIE
pu_codcia = LK_CODCIA
PUB_SERDOC = car_far2!car_SERDOC
PUB_NUMDOC = car_far2!car_NUMDOC
PUB_TIPDOC = car_far2!CAR_TIPDOC
LEER_CAR_LLAVE
If Val(car_llave!CAR_IMP_INI) = 0 Then GoTo SSS
SQ_OPER = 1
pu_codcia = LK_CODCIA
pu_cp = PUB_CP
PUB_NUM_OPER = PUB_NUM_OPER_EXT
LEER_CAA_LLAVE
If caa_histo.EOF = True Then
   MsgBox "DOCUMENTO EN CARACU NO EXISTE..."
   exito = False
End If
Do Until caa_histo.EOF
    caa_histo.Edit
    caa_histo!CAA_ESTADO = "E"
    caa_histo.Update
    caa_histo.MoveNext
Loop
SSS:
' AGREGE ACV PARA CARACU
If Val(i_codcli.Text) <> car_llave!CAR_CODCLIE Then pub_flag_cambio = 1

If pub_flag_cambio = 0 And LK_CODTRA = 1401 Then
car_llave.Edit
car_llave!CAR_NUMSER_C = PUB_NUMSER_C
car_llave!CAR_NUMFAC_C = PUB_NUMFAC_C
car_llave!car_numguia = PUB_NUMGUIA
car_llave!car_SERguia = i_serguia.Text
If i_fecha_compra.Visible Then
  car_llave!CAR_FECHA_SUNAT = i_fecha_compra.Text
Else
  car_llave!CAR_FECHA_SUNAT = LK_FECHA_DIA
End If
car_llave!car_fecha_vcto = PUB_FECHA_VCTO
car_llave!car_fecha_vcto_orig = PUB_FECHA_VCTO
car_llave.Update
pub_signo_car = 0

GoTo ffin
End If

SQ_OPER = 1
pu_cp = all_llave!all_cp
pu_codcia = all_llave!all_codCIA
pu_codclie = all_llave!all_CODCLIE
PUB_CODCLIE = all_llave!all_CODCLIE
LEER_CLI_LLAVE



pub_signo_car = -1
PUB_IMPORTE_AMORT = Val(car_far2!CAR_IMP_INI)

fx = 32000
GoSub ACT8
fx = 0
'GoSub ACT7
pub_signo_car = SUT_LLAVE!SUT_SIGNO_CAR
PUB_IMPORTE_AMORT = Val(i_neto.Text)
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1

SQ_OPER = 1
pu_cp = SUT_LLAVE!SUT_cp
pu_codcia = LK_CODCIA
pu_codclie = Val(i_codcli.Text)
PUB_CODCLIE = pu_codclie
LEER_CLI_LLAVE


ffin:
pub_signo_arm = SUT_LLAVE!SUT_SIGNO_ARM
pub_signo_car = SUT_LLAVE!SUT_SIGNO_CAR


Return


ACT3:
' EL CAMPO SIGNO_CAR SOLO VALE EN LA 2745 Y PARA CONSULTAS
If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
   If Val(gridl.TextMatrix(gridl.Row, 5)) <> 0 Then Return
End If

If PUB_IMPORTE_AMORT = 0 Then Return

If pub_signo_car = 0 Then
   Return
End If

SQ_OPER = 3
pu_codcia = LK_CODCIA
pu_cp = PUB_CP
LEER_CAR_LLAVE
If car_menor.EOF Then
   PUB_NUMDOC = 1
Else
   PUB_NUMDOC = car_menor!car_NUMDOC + 1
End If

car_llave.AddNew
car_llave!CAR_CODCLIE = PUB_CODCLIE
car_llave!CAR_CODCIA = LK_CODCIA
car_llave!car_numguia = PUB_NUMGUIA
car_llave!car_SERguia = Val(Nulo_Valor0(i_serguia.Text))
car_llave!CAR_TIPDOC = PUB_TIPDOC
car_llave!CAR_cp = PUB_CP
car_llave!car_SERDOC = PUB_SERDOC
car_llave!car_NUMDOC = PUB_NUMDOC
car_llave!CAR_FECHA_INGR = LK_FECHA_DIA
car_llave!car_fecha_vcto = PUB_FECHA_VCTO
car_llave!car_fecha_vcto_orig = PUB_FECHA_VCTO

car_llave!CAR_COMISION = 0

car_llave!CAR_SITUACION = " "
car_llave!CAR_NUM_REN = 0
car_llave!car_concepto = PUB_CONCEPTO
car_llave!car_nombre_banco = PUB_NOMBRE_BANCO
car_llave!CAR_NUM_CHEQUE = PUB_NUM_CHEQUE
car_llave!car_SIGNO_CAJA = pub_signo_caja
car_llave!car_numguia = PUB_NUMGUIA
WS_IMPORTE_AMORT = PUB_IMPORTE_AMORT

WS_TOT = WS_IMPORTE_AMORT
car_llave!CAR_IMP_INI = WS_IMPORTE_AMORT * pub_signo_car
car_llave!CAR_IMPORTE = WS_IMPORTE_AMORT * pub_signo_car
If LK_CODTRA = 2101 And PUB_SECUENCIA = 15 Then
   car_llave!CAR_IMPORTE = 0
   PUB_IMPORTE_AMORT = 0
   WS_IMPORTE_AMORT = 0
End If
If i_fecha_compra.Visible Then
   car_llave!CAR_FECHA_SUNAT = i_fecha_compra.Text
Else
   car_llave!CAR_FECHA_SUNAT = LK_FECHA_DIA
End If


LOC_SALDO_CAR = WS_IMPORTE_AMORT
wS_saldo_car = car_llave!CAR_IMPORTE
car_llave!car_codtra = LK_CODTRA
'If PUB_PRECIO2 <> 0 Then CREO QUE NO SE USA
'   car_llave!car_PRECIO = 1 / PUB_PRECIO2
'Else
car_llave!car_PRECIO = 0
'End If

car_llave!CAR_MONEDA = WS_MONEDA_CLI
If LK_EMP = "PLA" And LK_CODTRA = 2745 Then
   car_llave!CAR_MONEDA = SUT_LLAVE!SUT_MONEDA_CAJA
End If
car_llave!CAR_signo_car = pub_signo_car
car_llave!CAR_numSER = PUB_NUMSER
car_llave!CAR_numfac = PUB_NUMFAC

'If PUB_TIPMOV = 10 And LK_CODTRA = 2105 Then
'   PUB_NUMSER = PUB_NUMSER_C
'   PUB_NUMFAC = PUB_NUMFAC_C
'End If
car_llave!CAR_NUMOPER = PUB_NUM_OPER_XXX

car_llave!CAR_TIPMOV = PUB_TIPMOV
car_llave!CAR_FBG = PUB_FBG
car_llave!CAR_codven = PUB_CODVEN
car_llave!CAR_COBRADOR = PUB_CODVEN
car_llave!CAR_NUMSER_C = PUB_NUMSER_C
car_llave!CAR_NUMFAC_C = PUB_NUMFAC_C
car_llave!CAR_codban = PUB_CODBAN
car_llave!CAR_PLACA = i_placa.Text
If LK_CODTRA = 2101 Then
   car_llave!CAR_CODART = grid_fac.TextMatrix(2, 16)
End If
If PUB_SECUENCIA = 5 And LK_CODTRA = 2101 Then
   car_llave!CAR_VOUCHER = "V=" & Trim(i_concepto.Text)
   car_llave!car_concepto = "V=" & Trim(i_concepto.Text)
End If
car_llave.Update
ws_flag_car = ingre

GoSub registra_caa

Return

ACT4:
pub_signo_arm = 0  'asegura no en facart cuando se extorna

FILAX = 2
pub_signo_car = -1
Do Until FILAX > gridl.Rows - 1
   If Val(gridl.TextMatrix(FILAX, 5)) = 0 And Val(gridl.TextMatrix(FILAX, 4)) = 0 Then GoTo pasando

   PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
   PUB_NUMDOC = gridl.TextMatrix(FILAX, 12)
   PUB_SERDOC = gridl.TextMatrix(FILAX, 11)
   PUB_TIPDOC = gridl.TextMatrix(FILAX, 13)
   PUB_IMPORTE_AMORT = Val(gridl.TextMatrix(FILAX, 5))
   PUB_CONCEPTO = "Cancel. X Canje"
   SQ_OPER = 1
   pu_cp = PUB_CP
   pu_codclie = PUB_CODCLIE
   pu_codcia = LK_CODCIA
   LEER_CAR_LLAVE
   If car_llave.EOF Then
      pub_mensaje = "!!! ERROR EN ACTUALIZACION de CANJE..."
      exito = False
      GoTo Error_fatal
   End If
   car_llave.Edit
   
WS_TOT = PUB_IMPORTE_AMORT
WS_IMPORTE_AMORT = PUB_IMPORTE_AMORT
car_llave!CAR_IMPORTE = car_llave!CAR_IMPORTE + WS_IMPORTE_AMORT * pub_signo_car
car_llave!CAR_NUMSER_C = PUB_NUMSER
car_llave!CAR_NUMFAC_C = PUB_NUMFAC
car_llave!car_concepto = "Canje x " & BolFac.Caption & i_numser.Text & "-" & i_numfac.Text

wS_saldo_car = car_llave!CAR_IMPORTE
LOC_SALDO_CAR = car_llave!CAR_IMPORTE
WS_MONEDA_CLI = car_llave!CAR_MONEDA
ws_mon_canje = car_llave!CAR_MONEDA
ws_fecha_anterior = car_llave!car_fecha_vcto

car_llave.Update
GoSub registra_caa
ws_flag_car = MODIF
PUB_TIPMOV = 0
GoSub ACT7
pasando:
FILAX = FILAX + 1
Loop
PUB_TIPMOV = SUT_LLAVE!SUT_TIPMOV


PUB_IMPORTE_AMORT = Val(i_neto.Text)
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
pub_signo_car = SUT_LLAVE!SUT_SIGNO_CAR
pub_signo_arm = SUT_LLAVE!SUT_SIGNO_ARM

PUB_CONCEPTO = "Ingreso x Canje"
PUB_TIPDOC = "FA"

If Val(gridl.TextMatrix(1, 5)) = 0 Then
   pub_signo_car = 0
Else
   cli_llave.Edit
   cli_llave!cli_fecha_fac = DateAdd("d", Nulo_Valor0(cli_llave!cli_DIAS_FAC), LK_FECHA_DIA)
   cli_llave.Update
End If

Return

ACT5:
flag_salto = 1

fila = 2
FLAG = False
Do While FLAG = False
   If Val(grid_fac.TextMatrix(fila, 16)) = 0 Or (Val(grid_fac.TextMatrix(fila, 3)) = 0 And Val(grid_fac.TextMatrix(fila, 4)) = 0) Then
      GoTo OTROMASS
   End If
   
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   If LK_CODTRA = 2107 And i_cambio.Value = 1 Then
    PU_TIPMOV = PUB_TIPMOV
    PU_NUMSER = PUB_NUMSER
    PU_NUMFAC = PUB_NUMFAC
    PU_FBG = PUB_FBG
   Else
    PU_TIPMOV = Val(grid_fac.TextMatrix(fila, 26))
    PU_NUMSER = Val(grid_fac.TextMatrix(fila, 28))
    PU_NUMFAC = Val(grid_fac.TextMatrix(fila, 30))
    PU_FBG = grid_fac.TextMatrix(fila, 29)
   End If
   LEER_FAR_LLAVE
   Do Until far_llave.EOF
      If far_llave!FAR_NUMSEC = Val(grid_fac.TextMatrix(fila, 31)) Then
         far_llave.Edit
         far_llave!FAR_CONSIG = Val(grid_fac.TextMatrix(fila, 4)) + far_llave!FAR_CONSIG
         If LK_CODTRA <> 2125 Then far_llave!far_concepto = i_fbg.Text & "/ " & i_numser.Text & " - " & i_numfac.Text & " (Canje)"
         far_llave.Update
         Exit Do
      End If
      far_llave.MoveNext
   Loop
   
OTROMASS:
   fila = fila + 1
   If fila > grid_fac.Rows - 1 Then
      FLAG = True
   End If
Loop

Return

Return
ACT6:
If LK_CODTRA <> 1402 Then Return
'  CAMTEX Or gridC.TextMatrix(2, 1) = ""
  far_llave.MoveFirst
   Do Until far_llave.EOF
      If far_llave!FAR_estado2 = "P" Then
         far_llave.Edit
         far_llave!FAR_estado2 = "Q"
         far_llave.Update
      End If
      far_llave.MoveNext
   Loop
   
Return

ACT7:

If PUB_NUM_OPER_XXX = 0 Then Return
all_llave.AddNew
all_llave!ALL_NUMOPER = PUB_NUM_OPER_XXX
all_llave!all_codCIA = LK_CODCIA
all_llave!ALL_CODTRA = LK_CODTRA
all_llave!all_flag_ext = Nulo_Valors(tra_llave!TRA_FLAG_EXT)
If PUB_TIPMOV = 20 And Check1.Value = 1 Then
   all_llave!all_flag_ext = "E"
End If
all_llave!all_CODCLIE = PUB_CODCLIE
all_llave!ALL_CODART = PUB_CODART
all_llave!ALL_IMPORTE_AMORT = PUB_IMPORTE_AMORT
all_llave!all_codusu = LK_CODUSU
all_llave!ALL_FBG = PUB_FBG
all_llave!ALL_CODVEN = PUB_CODVEN
all_llave!ALL_importe = PUB_IMPORTE
all_llave!ALL_IMPORTE_DOLL = 0
If PUB_TIPMOV = 20 Then 'PENDIENTE
   If Tipo_Cambio.Visible = True And cli_llave!CLI_MONEDA = "D" Then
      If Mid(Tipo_Cambio.Caption, 2, 1) = "S" Then all_llave!ALL_IMPORTE_DOLL = WS_IMPORTE_AMORT
   End If
End If
all_llave!ALL_NUMDOC = PUB_NUMDOC
all_llave!all_cp = PUB_CP
all_llave!ALL_TIPDOC = PUB_TIPDOC
all_llave!ALL_NUMFAC_C = PUB_NUMFAC_C
all_llave!ALL_NUMSER_C = PUB_NUMSER_C
all_llave!ALL_CODBAN = PUB_CODBAN
all_llave!ALL_CONCEPTO = PUB_CONCEPTO
all_llave!ALL_CHENUM = PUB_CHENUM
all_llave!ALL_FECHA_DIA = LK_FECHA_DIA '**
If i_fecha_compra.Visible = True Then
  all_llave!ALL_FECHA_SUNAT = i_fecha_compra.Text
Else
  all_llave!ALL_FECHA_SUNAT = LK_FECHA_DIA
End If
all_llave!ALL_FECHA_VCTO = PUB_FECHA_VCTO '**
If IsDate(ws_fecha_anterior) Then all_llave!all_fecha_ant = ws_fecha_anterior

all_llave!ALL_CANTIDAD = PUB_CANTIDAD
If (LK_CODTRA = 2211 Or LK_CODTRA = 2103) And PUB_SECUENCIA = 1 Then
  all_llave!all_numser = PUB_PEDSER
  all_llave!all_numfac = PUB_PEDFAC
Else
 all_llave!all_numser = PUB_NUMSER
 all_llave!all_numfac = PUB_NUMFAC

End If
all_llave!all_neto = PUB_NETO  '**IMPORTE

all_llave!ALL_bruto = PUB_SUBTOTAL
all_llave!all_tipmov = PUB_TIPMOV
all_llave!ALL_impto = PUB_IMPTO
all_llave!ALL_flete = PUB_FLETE
all_llave!ALL_HORA = Now
all_llave!ALL_LIMCRE_ACT = PUB_LIMCRE_ACT
all_llave!ALL_LIMCRE_ANT = PUB_LIMCRE_ANT
all_llave!ALL_DESCTO = PUB_DESCTO
all_llave!ALL_GASTOS = PUB_GASTOS
all_llave!ALL_PRECIO = PUB_PRECIO2

all_llave!ALL_MONEDA_CLI = WS_MONEDA_CLI
all_llave!all_moneda_ccm = WS_MONEDA_CCM
all_llave!ALL_MONEDA_CAJA = PUB_DS
If LK_EMP = "PLA" And LK_CODTRA = 2725 Then
   all_llave!ALL_MONEDA_CAJA = SUT_LLAVE!SUT_MONEDA_CAJA
   all_llave!ALL_MONEDA_CLI = SUT_LLAVE!SUT_MONEDA_CAJA
End If

all_llave!all_SECUENCIA = PUB_SECUENCIA
all_llave!ALL_SIGNO_CAR = pub_signo_car
all_llave!ALL_SIGNO_CAJA = pub_signo_caja

all_llave!ALL_SIGNO_CCM = pub_signo_ccm
all_llave!all_sIGNO_ARM = pub_signo_arm

all_llave!ALL_CHENUM = PUB_CHENUM
all_llave!ALL_CHESEC = PUB_CHESEC
all_llave!ALL_CHESER = PUB_CHESER
all_llave!ALL_SUBTRA = SUT_LLAVE!SUT_DESCRIPCION
If car_llave.RowCount > 0 Then
   all_llave!ALL_SITUACION = Nulo_Valors(car_llave!CAR_SITUACION)
Else
   all_llave!ALL_SITUACION = ""
End If
all_llave!all_codtra_ext = Left(TRANS.Text, 4)

all_llave!ALL_TIPO_CAMBIO = WS_TIPO_CAMBIO
If WS_TIPO_CAMBIO <= 1 Then
  all_llave!ALL_TIPO_CAMBIO = LK_TIPO_CAMBIO
End If
all_llave!ALL_FACART = WS_FLAG_FACART

If PUB_TIPMOV = 10 Or PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
   If all_llave!ALL_FBG = "F" Then
          WS_AUTOCON = "Fact. " & all_llave!all_numser & "-" & all_llave!all_numfac
   ElseIf all_llave!ALL_FBG = "B" Then
          WS_AUTOCON = "Bolet." & all_llave!all_numser & "-" & all_llave!all_numfac
   ElseIf all_llave!ALL_FBG = "G" Or all_llave!ALL_FBG = "P" Then
          WS_AUTOCON = "Guia. " & all_llave!all_numser & "-" & all_llave!all_numfac
   ElseIf all_llave!ALL_FBG = "N" Then
          WS_AUTOCON = "N.CRED. " & all_llave!all_numser & "-" & all_llave!all_numfac
   ElseIf all_llave!ALL_FBG = "D" Then
          WS_AUTOCON = "N.DEB.  " & all_llave!all_numser & "-" & all_llave!all_numfac
   End If
   WS_AUTOCON = WS_AUTOCON & ":" & Trim(Left(cli_llave!cli_nombre, 25))
End If

If Val(all_llave!ALL_CODBAN) <> 0 And Not ccm_llave.EOF Then
   If all_llave!ALL_CHENUM <> 0 And SUT_LLAVE!SUT_cp = "P" And Not cli_llave.EOF And PUB_CODCLIE <> 0 Then
      WS_AUTOCON = SUT_LLAVE!SUT_abreviado & ":" & all_llave!ALL_CHENUM & ", Fact." & all_llave!ALL_NUMFAC_C & ":" & Trim(Left(cli_llave!cli_nombre, 18)) & ":" & Left(ccm_llave!ccm_nombre, 15)
   ElseIf all_llave!ALL_CHENUM <> 0 Then
      WS_AUTOCON = SUT_LLAVE!SUT_abreviado & ":" & all_llave!ALL_CHENUM & ":" & Trim(Left(ccm_llave!ccm_nombre, 15))
   Else
      WS_AUTOCON = Trim(ccm_llave!ccm_nombre) & ":" & i_concepto.Text
   End If
End If
      
If pub_signo_car = -1 And pub_signo_caja = 1 And Val(gridl.TextMatrix(1, 5)) <> 0 Then
   WS_AUTOCON = ""
   If car_llave!CAR_FBG = "F" Then
          WS_AUTOCON = "Fact. " & car_llave!CAR_numSER & "-" & car_llave!CAR_numfac
   ElseIf car_llave!CAR_FBG = "B" Then
          WS_AUTOCON = "Bolet." & car_llave!CAR_numSER & "-" & car_llave!CAR_numfac
   ElseIf car_llave!CAR_FBG = "G" Or car_llave!CAR_FBG = "P" Then
          WS_AUTOCON = "Guia. " & car_llave!CAR_numSER & "-" & car_llave!CAR_numfac
   End If
   WS_AUTOCON = WS_AUTOCON & " " & Trim(Left(cli_llave!cli_nombre, 18))
End If
 
If pub_signo_caja = 1 And Val(gridl.TextMatrix(1, 19)) <> 0 And PUB_TIPDOC = "CH" Then
   WS_AUTOCON = "Cheque N. " & PUB_NUM_CHEQUE
   WS_AUTOCON = WS_AUTOCON & ":" & Left(cli_llave!cli_nombre, 18)
End If

If Trim(PUB_CONCEPTO) <> "" Then
   WS_AUTOCON = WS_AUTOCON & ":" & PUB_CONCEPTO
End If
'   If ws_codclie <> 0 Then
'         WS_NOMCLI = Left(cli_llave!cli_nombre, 18) & ":"
'   End If
'   If all_llave!ALL_SIGNO_CAR <> 0 Then
'       WS_NOMCLI = Left(WS_NOMCLI, 30) & car_llave!CAR_fbg & "-." & car_llave!CAR_numSER & "-" & car_llave!CAR_numfac
'   End If
WS_AUTOCON = Left(SUT_LLAVE!SUT_DESCRIPCION, 18) & ":" & WS_AUTOCON
If Not cli_llave.EOF And PUB_CODCLIE <> 0 Then
   If Trim(WS_AUTOCON) = "" Then WS_AUTOCON = Left(SUT_LLAVE!SUT_DESCRIPCION, 25) & "  : " & Left(cli_llave!cli_nombre, 18) & ":" & all_llave!ALL_CONCEPTO
End If
all_llave!all_numoper2 = WS_NUM_OPER2
all_llave!ALL_autocon = WS_AUTOCON
all_llave!ALL_FLAG_SO = PUB_SO


all_llave.Update

Return


ACT8:
If pub_signo_car = 0 Then
   Return
End If

car_llave.Edit

WS_TOT = PUB_IMPORTE_AMORT
WS_IMPORTE_AMORT = PUB_IMPORTE_AMORT

If gridl.Visible = True Then
   WS_TOT = Val(gridl.TextMatrix(1, 9))
End If

car_llave!CAR_IMPORTE = car_llave!CAR_IMPORTE + WS_IMPORTE_AMORT * pub_signo_car
wS_saldo_car = car_llave!CAR_IMPORTE

LOC_SALDO_CAR = car_llave!CAR_IMPORTE

ws_fecha_anterior = car_llave!car_fecha_vcto


If (i_fecha_vcto.Visible = True And LK_CODTRA <> 1111) Or gridl.Visible = True Or grid_liq.Visible = True Then
   car_llave!car_fecha_vcto = PUB_FECHA_VCTO
End If



If i_concepto.Visible = True And LK_CODTRA <> 2412 And LK_CODTRA <> 2410 Then
   car_llave!car_concepto = PUB_CONCEPTO
End If

If PUB_IMPORTE_AMORT <> 0 And LK_CODTRA <> 1111 And LK_CODTRA <> 1122 And LK_CODTRA <> 1133 And car_llave!CAR_IMPORTE <> 0 Then
   car_llave!CAR_NUM_REN = car_llave!CAR_NUM_REN + 1
End If

If LK_EMP = "HER" Then car_llave!CAR_COBRADOR = Val(gridl.TextMatrix(FILAX, 7))
If LK_CODTRA = 1111 Or LK_CODTRA = 1122 Then  ' para mantener enlace del doc.
 PUB_NUMSER_C = car_llave!CAR_NUMSER_C
 PUB_NUMFAC_C = car_llave!CAR_NUMFAC_C
End If
If LK_CODTRA = 2410 Or LK_CODTRA = 2412 Then
  PUB_NUMSER_C = Val(i_numser_c.Text)
  PUB_NUMFAC_C = Val(i_numfac_c.Text)
End If

If LK_CODTRA = 1111 And ww_codtra_ext = 2105 Then ' para mantener enlace del doc.
 car_llave!car_concepto = ""
End If

car_llave.Update

GoSub registra_caa

ws_flag_car = MODIF
Return

ACT9:
If PUB_TIPMOV <> 180 Then Return
fila = 2
i = 0
Do Until fila = GridT.Rows - 1
    PUB_PEDFAC = Val(GridT.TextMatrix(fila, 7))
    pu_codcia = LK_CODCIA
    PUB_PEDSER = 0
    PUB_TIPMOV = 177
    SQ_OPER = 1
    LEER_PED_LLAVE
    PUB_TIPMOV = 180
    If ped_llave.EOF = True Then
       pub_mensaje_err = "El Nro. de Pedido no Existe"
       exito = False
    End If
    Do Until ped_llave.EOF
        If Val(ped_llave!PED_CODART) = Val(GridT.TextMatrix(fila, 1)) Then
             ped_llave.Edit
             ped_llave!PED_salidas = ped_llave!PED_salidas + Val(GridT.TextMatrix(fila, 3))
             ped_llave.Update
             GoTo estabien2
        End If
     ped_llave.MoveNext
    Loop
       pub_mensaje_err = "El Nro. de Pedido, no Corresponde al articulo"
       exito = False
estabien2:
fila = fila + 1
Loop


Return

ACT10:
If LK_CODTRA <> 2407 Then Return
FAR_TRANS.MoveFirst
Do Until FAR_TRANS.EOF
   If Val(FAR_TRANS!FAR_NUMFAC) = Val(grid_trans.TextMatrix(grid_trans.Row, 0)) Then
      FAR_TRANS.Edit
      FAR_TRANS!far_transito = "X"
      FAR_TRANS.Update
   End If
   FAR_TRANS.MoveNext
Loop
 
Return

ACT11:
If i_numplan.Visible = True And i_codcli.Visible = False Then
   par_llave.Edit
   par_llave!par_planilla = par_llave!par_planilla + 1
   i_numplan.Text = par_llave!par_planilla
   pub_numplan = i_numplan.Text
   par_llave.Update
End If

If Val(gridl.TextMatrix(1, 5)) = 0 And (PUB_TIPMOV = 97 Or PUB_TIPMOV = 98) Then
   GoSub ACT7
   Return
End If

If Val(gridl.TextMatrix(2, 12)) = 0 And Val(gridl.TextMatrix(1, 5)) = 0 Then GoTo continua

FILAX = 2

Do Until FILAX > gridl.Rows - 1
      If SUT_LLAVE!SUT_SIGNO_CAR = 2 Then
         If Val(gridl.TextMatrix(FILAX, 25)) < 0 Then
            pub_signo_car = 1
         Else
            pub_signo_car = -1
         End If
      Else
        pub_signo_car = SUT_LLAVE!SUT_SIGNO_CAR
      End If
      PUB_NUMSER_C = Val(gridl.TextMatrix(FILAX, 2))
      PUB_NUMFAC_C = Val(gridl.TextMatrix(FILAX, 3))
      PUB_CODCLIE = Val(gridl.TextMatrix(FILAX, 15))
      PUB_CP = gridl.TextMatrix(FILAX, 10)
      PUB_NUMDOC = Val(gridl.TextMatrix(FILAX, 12))
      PUB_SERDOC = Val(gridl.TextMatrix(FILAX, 11))
      PUB_TIPDOC = gridl.TextMatrix(FILAX, 13)
      If LK_EMP = "PIU" Then
         PUB_IMPORTE_AMORT = Val(gridl.TextMatrix(FILAX, 5)) + Val(gridl.TextMatrix(FILAX, 19))
      Else
         PUB_IMPORTE_AMORT = Val(gridl.TextMatrix(FILAX, 5))
      End If
      
      If IsDate(gridl.TextMatrix(FILAX, 6)) = True Then PUB_FECHA_VCTO = gridl.TextMatrix(FILAX, 6)
      If PUB_IMPORTE_AMORT <> 0 Then
         ww_concepto = Trim(gridl.TextMatrix(FILAX, 1)) & "./ " & gridl.TextMatrix(FILAX, 2) & "-" & gridl.TextMatrix(FILAX, 3)
      End If
      If Trim(gridl.TextMatrix(FILAX, 23)) <> "" Then
        ww_concepto = Trim(gridl.TextMatrix(FILAX, 23))
        'gridl.TextMatrix(FILAX, 23) = ""
        i_concepto.Text = ""
      End If
      If Trim(i_concepto.Text) = "" Then i_concepto.Text = ww_concepto
      
      SQ_OPER = 1
      pu_cp = PUB_CP
      pu_codclie = PUB_CODCLIE
      pu_codcia = LK_CODCIA
      LEER_CAR_LLAVE
      If car_llave.EOF Then
         pub_mensaje = "!!! ERROR EN ACTUALIZACION de CANJE..."
         exito = False
         GoTo Error_fatal
      End If
      WS_MONEDA_CLI = car_llave!CAR_MONEDA
      ws_mon_canje = car_llave!CAR_MONEDA

      pub_signo_ccm = 0
      If PUB_IMPORTE_AMORT = 0 Then
         If car_llave!car_fecha_vcto <> PUB_FECHA_VCTO Then
            car_llave.Edit
            car_llave!car_fecha_vcto = PUB_FECHA_VCTO
            car_llave.Update
         End If
         If car_llave!CAR_COBRADOR <> Val(gridl.TextMatrix(FILAX, 7)) And LK_CODTRA = 2770 Then
            car_llave.Edit
            car_llave!CAR_COBRADOR = Val(gridl.TextMatrix(FILAX, 7))
            car_llave.Update
         End If
      Else
         PUB_CONCEPTO = i_concepto.Text
'         PUB_CONCEPTO = Trim(gridl.TextMatrix(FILAX, 1)) & "./ " & gridl.TextMatrix(FILAX, 2) & "-" & gridl.TextMatrix(FILAX, 3)
         PUB_CODCLIE = car_llave!CAR_CODCLIE
         GoSub CON1
         GoSub ACT8
         GoSub ACT7
         PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
      End If
      
      FILAX = FILAX + 1
Loop
   pub_signo_ccm = SUT_LLAVE!SUT_signo_ccm
   If pub_signo_ccm = 0 Then GoTo continua
   PUB_IMPORTE_AMORT = 0
   If pub_signo_ccm = 0 Then
      PUB_NUM_OPER_XXX = 0
      Return
   End If
   
   pub_signo_caja = 0
   If pub_signo_ccm = 1 Then
      pub_signo_caja = -1
      PUB_CONCEPTO = "Dep.Bco. :" & ww_concepto
   ElseIf pub_signo_ccm = -1 Then
      PUB_CONCEPTO = Trim(ww_concepto) & "- " & Trim(Left(cli_llave!cli_nombre, 25))
   End If
   pub_signo_car = 0
   GoSub ACT5
   GoSub ACT7

continua:

If Val(gridl.TextMatrix(1, 19)) = 0 Then
   textovarl.Visible = False
   Return
End If

FILAX = 2
Do Until FILAX > gridl.Rows - 1
If Val(gridl.TextMatrix(FILAX, 19)) <> 0 Then
If LK_EMP = "PIU" Then
   PUB_TIPDOC = "CD"
   pub_signo_car = 1
Else
   PUB_TIPDOC = "CH"
   pub_signo_car = -1
   If LK_EMP = "HER" Then
     pub_signo_caja = 0
   End If
End If

PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
PUB_NUMGUIA = 0
PUB_FECHA_VCTO = gridl.TextMatrix(FILAX, 22)
PUB_CONCEPTO = gridl.TextMatrix(FILAX, 1) & "/." & gridl.TextMatrix(FILAX, 2) & "-" & gridl.TextMatrix(FILAX, 3)
PUB_IMPORTE_AMORT = gridl.TextMatrix(FILAX, 19)
PUB_NUM_CHEQUE = gridl.TextMatrix(FILAX, 20)
PUB_CODCLIE = gridl.TextMatrix(FILAX, 15)
'PUB_CODVEN = 0 ' lo quito por que quiero que grave el codigo de vendedor ** alan

GoSub CON1
PUB_NUMSER = 0
PUB_NUMFAC = 0
PUB_TIPMOV = 0
WS_MONEDA_CLI = gridl.TextMatrix(FILAX, 24)
ws_mon_canje = gridl.TextMatrix(FILAX, 24)

GoSub ACT3
GoSub ACT7
End If

FILAX = FILAX + 1
Loop
textovarl.Visible = False
PUB_NUM_OPER_XXX = 0
Return

ACT12:
cli_llave.Edit
If Frame1.Visible = True Then
   cli_llave!CLI_TIPO_BLOQ1 = BLOQ(0).Value
   cli_llave!CLI_TIPO_BLOQ2 = BLOQ(1).Value
   cli_llave!CLI_TIPO_BLOQ3 = BLOQ(2).Value
   cli_llave!CLI_TIPO_BLOQ4 = BLOQ(3).Value
End If

If LK_CODTRA = 2580 Then
If Nulo_Valor0(cli_llave!cli_limcre) <> PUB_LIMCRE_ACT Then
   cli_llave!CLI_FECHA_APROB = LK_FECHA_DIA
   cli_llave!cli_limcre = PUB_LIMCRE_ACT
End If
End If

cli_llave.Update
Return

ACT13:
      If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
         If Val(grid_fac.TextMatrix(1, 7)) <> 0 Then Return
      End If

      If i_cambio.Visible = False Or i_cambio.Value = 0 Then llena_numfac_grifo
      far_llave.AddNew
      far_llave!FAR_TIPMOV = PUB_TIPMOV
      far_llave!FAR_CODCIA = LK_CODCIA
'      far_llave!far_COD_SUNAT = Val(Right(i_codsunat.Text, 5))
      far_llave!far_numser = PUB_NUMSER
      far_llave!FAR_NUMFAC = PUB_NUMFAC
      far_llave!FAR_NUMSEC = WS_NUMSEC + 1
      far_llave!far_codart = 0
      far_llave!FAR_CANTIDAD = PUB_CANTIDAD
      far_llave!far_UNIDADES = 0
      far_llave!far_JABAS = 0
      far_llave!far_mortal = 0
      far_llave!far_FECHA_COMPRA = LK_FECHA_DIA
      far_llave!far_codclie = PUB_CODCLIE
      far_llave!FAR_CP = PUB_CP
      far_llave!far_transito = " "
     
      far_llave!FAR_STOCK = 0
      far_llave!FAR_COSPRO = PUB_COSPRO
      far_llave!FAR_PRECIO = PUB_PRECIO
      far_llave!FAR_FBG = PUB_FBG
      far_llave!far_IMPTO = PUB_IMPTO2
      far_llave!FAR_TOT_DESCTO = 0
      far_llave!FAR_DESCTO = 0
      far_llave!FAR_GASTOS = 0
      far_llave!far_bruto = PUB_BRUTO2
      PUB_SUBTOTAL = PUB_BRUTO2
      PUB_IMPTO = PUB_IMPTO2
      far_llave!FAR_NUMDOC = 0
      far_llave!far_NUMGUIA = PUB_NUMGUIA
      far_llave!FAR_tipo_cambio = 0
      far_llave!FAR_DESCTO = 0
      far_llave!FAR_TOT_DESCTO = 0
      far_llave!far_DIAS = 0
      far_llave!far_MONEDA = Nulo_Valors(cli_llave!CLI_MONEDA)
      far_llave!far_num_precio = "0"
      far_llave!far_otra_cia = ""
      far_llave!FAR_fecha = LK_FECHA_DIA
      far_llave!far_numser_c = PUB_NUMSER_C
      far_llave!far_numfac_c = PUB_NUMFAC_C
      far_llave!FAR_NUMOPER = PUB_NUM_OPER_XXX
      far_llave!far_tipo_bloq_act1 = PUB_TIPO_BLOQ_act1
      far_llave!far_tipo_bloq_act2 = PUB_TIPO_BLOQ_act2
      far_llave!far_tipo_bloq_act3 = PUB_TIPO_BLOQ_act3
      far_llave!far_tipo_bloq_act4 = PUB_TIPO_BLOQ_act4
      far_llave!FAR_TIPO_BLOQ_ANT1 = PUB_TIPO_BLOQ_ant1
      far_llave!far_tipo_bloq_ant2 = PUB_TIPO_BLOQ_ant2
      far_llave!far_tipo_bloq_ant3 = PUB_TIPO_BLOQ_ant3
      far_llave!far_tipo_bloq_ant4 = PUB_TIPO_BLOQ_ant4
      far_llave!far_estado = "N"
      If Val(i_neto.Text) = 0 And Trim(grid_fac.TextMatrix(2, 4)) = "ANULACION" And PUB_TIPMOV = 10 Then
         far_llave!far_estado = "E"
      End If
      far_llave!FAR_equiv = 1
      far_llave!FAR_descri = ""
      
      far_llave!far_ISLA = PUB_CODVEN
      far_llave!far_turno = Val(i_turno.Text)
      far_llave!far_NUMGUIA = Val(i_numguia.Text)
      far_llave!far_serguia = Val(i_serguia.Text)
      
      If LK_CODTRA = 2107 Then
       far_llave!FAR_codven = PUB_CODVEN
      Else
       far_llave!FAR_codven = pub_responsable
      End If
      
      
      far_llave!far_PESO = 0
      far_llave!far_codusu = LK_CODUSU
'      far_llave!far_CODTRA = SUT_LLAVE!SUT_CODTRA
'      far_llave!far_SUBCODTRA = SUT_LLAVE!SUT_SECUENCIA
'      If i_codsunat.ListIndex >= 0 Then far_llave!far_COD_SUNAT = Val(Right(i_codsunat.Text, 5))
      far_llave!far_concepto = PUB_CONCEPTO
      far_llave!far_SIGNO_aRM = pub_signo_arm
      'far_llave!FAR_CODVEN = PUB_CODVEN
      WS_FLAG_FACART = "A"
      If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
         far_llave!far_concepto = Nulo_Valors(i_TEXTONCRE.Text)
      End If
      
      far_llave.Update

Return
ACT14:
WS_IMPORTE_AMORT = 0
car_llave.Edit
car_llave!CAR_SITUACION = Left(i_situacion.Text, 1)
car_llave.Update
GoSub registra_caa
 
'Boton_Compras_Click 'no se usa
Return

ACT15:
If grid_canje.Visible = False Then Return
FILAX = 2
Do Until FILAX > grid_canje.Rows - 1
If Val(grid_canje.TextMatrix(FILAX, 3)) <> 0 Then
PUB_TIPDOC = grid_canje.TextMatrix(FILAX, 2)
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
PUB_NUMGUIA = 0
PUB_FECHA_VCTO = grid_canje.TextMatrix(FILAX, 6)
PUB_CONCEPTO = grid_canje.TextMatrix(FILAX, 4)
PUB_IMPORTE_AMORT = grid_canje.TextMatrix(FILAX, 3)
PUB_NUM_CHEQUE = grid_canje.TextMatrix(FILAX, 5)
PUB_CODCLIE = grid_canje.TextMatrix(FILAX, 8)
PUB_CODVEN = grid_canje.TextMatrix(FILAX, 9)

GoSub CON1
PUB_NUMSER = 0
PUB_NUMFAC = 0
PUB_TIPMOV = 0
i_fbg.ListIndex = -1
GoSub ACT3
GoSub ACT7
End If

FILAX = FILAX + 1
Loop

grid_canje.Top = 4300
grid_canje.Height = 700
grid_canje.Clear
'gridl.Clear

Return

ACT16:
'VIENE DE CONSIS_21 EXTORNOS
If pub_signo_car <> 0 Then
   caa_histo.Edit
   caa_histo!CAA_ESTADO = "E"
   ww_numdoc = caa_histo!CAa_numfac_C
   ww_numser = caa_histo!CAa_numser_C
   caa_histo.Update
End If

If (PUB_TIPMOV = 0 And Trim(WS_FLAG_FACART) = "") Or (ww_codtra_ext = 2103 And pub_signo_arm = 0) Then
   Return
End If
far_menor3.MoveFirst
Do Until far_menor3.EOF
   far_menor3.Edit
   far_menor3!far_estado = "E"
   PUB_NUMSER_C = far_menor3!far_numser_c
   PUB_NUMFAC_C = far_menor3!far_numfac_c
   far_menor3.Update
   far_menor3.MoveNext
Loop
far_llave.Requery
Return

ACT17:
CONTADOR = 0
Do Until CONTADOR = PUB_CANT_CHEQ
che_llave.AddNew
che_llave!CHE_CODBAN = PUB_CODBAN
If LK_EMP = "3AA" Then
    che_llave!CHE_CODCIA = par_llave!PAR_CIACCM
Else
    che_llave!CHE_CODCIA = LK_CODCIA
End If

che_llave!che_cheser = PUB_CHESER
che_llave!che_chesec = 0
che_llave!che_chenum = CONTADOR + PUB_NUM_INI
CONTADOR = CONTADOR + 1
che_llave!che_importe = 0
che_llave!CHE_FECHA = 0
che_llave!CHE_FECHA_EMISION = 0
che_llave!CHE_CODUSU = LK_CODUSU
che_llave!che_concepto = " "
che_llave!CHE_NUMOPER = CONTADOR + 30000
che_llave!che_saldo = 0
che_llave!CHE_SIGNO_CCM = 0
che_llave!che_estado = " "
che_llave!che_abreviado = ""
che_llave!CHE_CODTRA = LK_CODTRA
che_llave!che_chenum_ext = 0

che_llave.Update
Loop
PUB_CHENUM = 0
Return

ACT18:
SQ_OPER = 1
PUB_FECHA = LK_FECHA_DIA
pu_codcia = LK_CODCIA
LEER_ALL_LLAVE
Do Until all_llave.EOF
  If PUB_NUM_OPER_EXT = all_llave!ALL_NUMOPER Then
    all_llave.Edit
    all_llave!all_flag_ext = "E"
    all_llave.Update
    Exit Do
  Else
    all_llave.MoveNext
  End If
Loop
If all_llave.EOF Then
   pub_mensaje_err = "Error en Extorno de Allog ..."
   GoTo Error_fatal
End If

If all_llave!ALL_CODTRA = 1455 Then pub_signo_ccm = 0
If pub_signo_ccm <> 0 Then
   SQ_OPER = 2
   PUB_FECHA = LK_FECHA_DIA
   LEER_CHE_LLAVE
   If che_oper.EOF = True Then
      pub_mensaje_err = "Error en Bancos..no existe tal cuenta..."
      GoTo Error_fatal
   End If
   che_oper.Edit
   che_oper!che_estado = "E"
   che_oper.Update
   PUB_CHESEC = che_oper!che_chesec + 50
   PUB_CHENUM_EXT = PUB_CHENUM
   PUB_CHENUM = 0
   PUB_ABREVIADO = "EXT"
End If

Return

ACT19:
   
   If PUB_TIPMOV = 97 Then
   If Nulo_Valors(par_llave!PAR_FLAG_NCRE) = "A" Then
      par_llave.Edit
      par_llave!PAR_FLAG_NCRE = ""
      par_llave.Update
   End If
   End If
   
   If PUB_TIPMOV = 98 Then
   If Nulo_Valors(par_llave!PAR_FLAG_NDEB) = "A" Then
      par_llave.Edit
      par_llave!PAR_FLAG_NDEB = ""
      par_llave.Update
   End If
   End If

   
   If i_fbg.Visible = False Then
      Return
   End If
   
   
   If LK_FLAG_FACTURACION = "A" And PUB_TIPMOV = 10 Then
   If par_llave!PAR_FLAG_G = "A" Or par_llave!PAR_FLAG_B = "A" Or par_llave!PAR_FLAG_F = "A" Then
      par_llave.Edit
      If PUB_FBG = "G" And par_llave!PAR_FLAG_G = "A" Then
         par_llave!PAR_FLAG_G = ""
      ElseIf PUB_FBG = "B" And par_llave!PAR_FLAG_B = "A" Then
         par_llave!PAR_FLAG_B = ""
      ElseIf PUB_FBG = "F" And par_llave!PAR_FLAG_F = "A" Then
         par_llave!PAR_FLAG_F = ""
      ElseIf PUB_FBG = "R" And par_llave!PAR_FLAG_R = "A" Then
         par_llave!PAR_FLAG_R = ""
      End If
      par_llave.Update
   End If
   End If
   
   
   
   
   
If LK_FLAG_FACTURACION = "V" And PUB_TIPMOV = 10 Then
   If (ven_llave!VEM_FLAG_G = "A" Or ven_llave!VEM_FLAG_B = "A" Or ven_llave!VEM_FLAG_F = "A" Or ven_llave!VEM_FLAG_P = "A" Or ven_llave!VEM_FLAG_R = "A") Then
        ven_llave.Edit
        If PUB_FBG = "G" And ven_llave!VEM_FLAG_G = "A" Then
           ven_llave!VEM_FLAG_G = ""
        ElseIf PUB_FBG = "B" And ven_llave!VEM_FLAG_B = "A" Then
           ven_llave!VEM_FLAG_B = ""
        ElseIf PUB_FBG = "F" And ven_llave!VEM_FLAG_F = "A" Then
           ven_llave!VEM_FLAG_F = ""
        ElseIf PUB_FBG = "P" And ven_llave!VEM_FLAG_P = "A" Then
           ven_llave!VEM_FLAG_P = ""
        End If
        ven_llave.Update
    End If
End If

Return

ACT20:
If PUB_TIPMOV = 0 Then
   Return
End If
If i_cambio.Visible = True And i_cambio.Value = 1 Then
If ped_llave.RowCount > 0 Then ped_llave.MoveFirst
Do Until ped_llave.EOF
   ped_llave.Edit
   ped_llave.Delete
   ped_llave.MoveNext
Loop
End If

ww_numsec = 0
fila = 2
FLAG = False
Do While FLAG = False
                                '  CAMTEX QUITE AQUI
   If Val(GridT.TextMatrix(fila, 5)) = 0 Then GoTo OTROMAS2
   If Trim(GridT.TextMatrix(fila, 0)) = "" Then GoTo OTROMAS2
   WS_CANTIDAD = Val(GridT.TextMatrix(fila, 3))
   PUB_CODART = Val(GridT.TextMatrix(fila, 5))
   ped_llave.AddNew
   ped_llave!PED_TIPMOV = PUB_TIPMOV
   ped_llave!PED_CODCIA = LK_CODCIA
   ped_llave!PED_NUMSER = Val(i_ser_lote.Text)
   ped_llave!PED_NUMFAC = Val(i_num_lote.Text)
   ww_numsec = ww_numsec + 1
   ped_llave!PED_NUMSEC = ww_numsec
   ped_llave!PED_CODART = PUB_CODART
   ped_llave!PED_CANTIDAD = Val(WS_CANTIDAD)
   ped_llave!PED_ROLLOS = Val(GridT.TextMatrix(fila, 2))
   ped_llave!PED_CODCLIE = PUB_CODCLIE
   WS_MONEDA_CLI = i_ds.Text
   ped_llave!PED_MONEDA = WS_MONEDA_CLI
'   ped_llave!ped_cp = PUB_CP
   ped_llave!PED_FECHA = LK_FECHA_DIA
   ped_llave!PED_PRECIO = 0
   ped_llave!PED_EQUIV = 0
'   ped_llave!PED_CIERRE = "X"
   ped_llave!PED_CODUSU = LK_CODUSU
     ' AGREGE CAMTEX
   ped_llave!PED_CONTACTO = Trim(GridT.TextMatrix(fila, 6))
   ped_llave!PED_numfac2 = Val(GridT.TextMatrix(fila, 7))
   ' CREE EN LA B/D CAMTEX
   ped_llave!PED_SERGUIA = Val(i_serguia.Text)
   ped_llave!PED_NUMGUIA = Val(i_numguia.Text)
   ped_llave!PED_CODVEN = Val(i_codven.Text)
   ped_llave!PED_ESTADO = " "
   If LK_CODTRA = 2103 And PUB_SECUENCIA = 10 Then ped_llave!PED_ESTADO = "P"
   
   If i_fecha_compra.Visible Then
     ped_llave!PED_FECHA_EMISION = i_fecha_compra.Text
   Else
     ped_llave!PED_FECHA_EMISION = LK_FECHA_DIA
   End If
   For i = 1 To 14
      ped_llave(30 + i) = Text1(i).Text
   Next i
   
   ped_llave.Update
OTROMAS2:
   fila = fila + 1
   If fila > GridT.Rows - 1 Then
      FLAG = True
   End If
Loop
Return

ACT21:

Return


'ACT20:
'If PUB_TIPMOV = 0 Then
'   Return
'End If

'If i_numfac.Text = "" Then llena_PEDIDO

'fila = 2
'FLAG = False
'Do While FLAG = False
'   If Val(grid_fac.TextMatrix(fila, 16)) = 0 Or (Val(grid_fac.TextMatrix(fila, 3)) = 0 And Val(grid_fac.TextMatrix(fila, 4)) = 0) Then
'      GoTo OTROMAS
'   End If
'
'   PUB_CODART = Val(grid_fac.TextMatrix(fila, 16))
'   PUB_JABAS = Val(grid_fac.TextMatrix(fila, 2))
'   PUB_UNIDAD = Val(grid_fac.TextMatrix(fila, 3))
'   WS_CANTIDAD = Val(grid_fac.TextMatrix(fila, 4))
'
'   PUB_PRECIO2 = Val(grid_fac.TextMatrix(fila, 6))
'   PUB_COSPRO = Val(grid_fac.TextMatrix(fila, 11))
'      ped_llave.AddNew
'      ped_llave!PED_TIPMOV = PUB_TIPMOV
'      ped_llave!PED_CODCIA = LK_CODCIA
'      ped_llave!PED_NUMSER = PUB_NUMSER
''      ped_llave!ped_codven = Val(i_codven.text)
'      ped_llave!PED_NUMFAC = PUB_NUMFAC
'      WS_NUMSEC = WS_NUMSEC + 1
'      ped_llave!PED_NUMSEC = WS_NUMSEC
'      ped_llave!PED_CODART = PUB_CODART
'      ped_llave!PED_CANTIDAD = Val(WS_CANTIDAD)
'      ped_llave!PED_CODCLIE = PUB_CODCLIE
'      WS_MONEDA_CLI = i_ds.Text
'      ped_llave!PED_MONEDA = WS_MONEDA_CLI
'      ped_llave!ped_cp = PUB_CP
'      ped_llave!PED_FECHA = LK_FECHA_DIA
'      ped_llave!PED_PRECIO = PUB_PRECIO2
'      ped_llave!PED_EQUIV = Val(Nulo_Valor0(grid_fac.TextMatrix(fila, 14)))
'      ped_llave!ped_fbg = PUB_FBG
'      ped_llave!ped_NUMOPER = PUB_NUM_OPER_XXX
'      ped_llave!PED_UNIDAD = PUB_UNIDAD
'      ped_llave!PED_CODUSU = LK_CODUSU
'      ped_llave.Update
'OTROMAS2:
'   fila = fila + 1
'   If fila > grid_fac.Rows - 1 Then
'      FLAG = True
'   End If
'Loop
'Return

'ACT21:
'      FILAX = gridl.Row
'      PUB_NUMSER_C = Val(gridl.TextMatrix(FILAX, 2))
'      PUB_NUMFAC_C = Val(gridl.TextMatrix(FILAX, 3))
'      PUB_CODCLIE = Val(gridl.TextMatrix(FILAX, 15))
'      PUB_CP = gridl.TextMatrix(FILAX, 10)
'      PUB_NUMDOC = Val(gridl.TextMatrix(FILAX, 12))
'      PUB_SERDOC = Val(gridl.TextMatrix(FILAX, 11))
'      PUB_TIPDOC = gridl.TextMatrix(FILAX, 13)
'      PUB_IMPORTE_AMORT = Val(grid_liq.TextMatrix(1, 1))
'      PUB_FECHA_VCTO = gridl.TextMatrix(FILAX, 6)
'      ww_concepto = Trim(gridl.TextMatrix(FILAX, 1)) & "./ " & gridl.TextMatrix(FILAX, 2) & "-" & gridl.TextMatrix(FILAX, 3) & " : " & ww_concepto
'      PUB_CONCEPTO = i_concepto.Text
'      SQ_OPER = 1
'      pu_cp = PUB_CP
'      pu_codclie = PUB_CODCLIE
'      pu_codcia = LK_CODCIA
'      LEER_CAR_LLAVE
'      If car_llave.EOF Then
'         pub_mensaje = "!!! ERROR EN ACTUALIZACION de CANJE..."
'         exito = False
'         GoTo Error_fatal
'      End If
'      PUB_CODCLIE = car_llave!CAR_CODCLIE
'      WS_MONEDA_CLI = car_llave!CAR_MONEDA
'      GoSub CON1
'      GoSub ACT8
'      GoSub ACT7
'Return



REP1:

'Close #1
'*** Imprime la Transacción a un Archivo
'RUTA = PUB_RUTA_REPORTE & LK_CODUSU & ".txt"
'Open RUTA For Output As #1
'Close #1
'** CARGA EL EDITOR **
Screen.MousePointer = 0
'Load frmeditor
'frmeditor.Show 1



Return

REP2:

Return

REP3:

'PUB_TIPREG = 20
'PUB_NUMTAB = cli_llave!CLI_CASA_ZONA
'SQ_OPER = 1
'LEER_TAB_LLAVE
'If tab_llave.EOF = True Then
'   pub_mensaje_err = "ERROR EN ZONA.."
'End If
'zona_nombre = tab_llave!TAB_NOMLARGO

'PUB_TIPREG = 30
'PUB_NUMTAB = cli_llave!CLI_CASA_SUBZONA
'SQ_OPER = 1
'LEER_TAB_LLAVE
'subzona_nombre = tab_llave!TAB_NOMLARGO


'*** REPORTE DE LETRA ***
'IMP_LETRA CStr(WS_IMPRESION_LET), CStr(PUB_FECHA_VCTO), CStr(PUB_NETO), CStr(PUB_FECHA), CStr(PUB_FECHA_VCTO), CStr(PUB_NETO), cli_llave!cli_nombre, Nulo_Valors(cli_llave!CLI_CASA_DIREC), cli_llave!CLI_CASA_NUM, zona_nombre, subzona_nombre

Return
   

'Screen.MousePointer = 1






GoTo fin


Error_fatal:
    'MsgBox pub_mensaje
    pub_mensaje = "Se ha producido un error " & "al abrir la conexión:" & Err & " - " & Error & vbCr
    For Each er In rdoErrors
        pub_mensaje = pub_mensaje & er.Description & ":" & er.Number & vbCr
        MsgBox pub_mensaje, 48, Pub_Titulo
    Next er

    CN.Execute "Rollback Transaction", rdExecDirect

End
'    Resume AbandonCn
errorr:

fin:
Screen.MousePointer = 0
grabar.Enabled = True
PUB_CODBAN = 0
PUB_CODCLIE = 0
pasa_def
'i_def.SetFocus
Exit Sub

registra_caa:


If pub_signo_car <> 0 Then
cli_llave.Edit
WS_CANTIDAD = WS_IMPORTE_AMORT * pub_signo_car
cli_llave!cli_SALDO = Nulo_Valor0(cli_llave!cli_SALDO) + WS_CANTIDAD
'If pub_signo_car = -1 Then cli_llave!cli_fecha_fac = PUB_FECHA_VCTO
' AGREGE IF PARA NO LIMPIAR EL CLI_MONEDA POR QUE WS_MONEDA_CLI = ""
If Trim(WS_MONEDA_CLI) = "S" Or Trim(WS_MONEDA_CLI) = "D" Then cli_llave!CLI_MONEDA = WS_MONEDA_CLI
cli_llave.Update
End If

caa_histo.AddNew
caa_histo!CAA_CODCLIE = PUB_CODCLIE
caa_histo!CAA_CODCIA = LK_CODCIA
caa_histo!CAA_TIPDOC = PUB_TIPDOC
caa_histo!CAA_CP = PUB_CP
caa_histo!CAA_NUM_OPER = PUB_NUM_OPER_XXX
If grid_liq.Visible = True Then
   caa_histo!caa_INTVEN = grid_liq.TextMatrix(1, 2)
   caa_histo!caa_DIASV = grid_liq.TextMatrix(1, 6)
   caa_histo!caa_DIASA = grid_liq.TextMatrix(1, 7)
   caa_histo!caa_tasav = 0
Else
   caa_histo!caa_INTVEN = 0
   caa_histo!caa_DIASV = 0
   caa_histo!caa_DIASA = 0
   caa_histo!caa_tasav = 0
End If

caa_histo!caa_serdoc = PUB_SERDOC
caa_histo!CAA_NUMDOC = PUB_NUMDOC
caa_histo!CAA_FECHA = LK_FECHA_DIA
caa_histo!CAA_FECHA_VCTO = PUB_FECHA_VCTO
caa_histo!caa_situacion = PUB_SITUACION_ACT
caa_histo!caa_concepto = PUB_CONCEPTO
caa_histo!CAA_IMPORTE = WS_IMPORTE_AMORT * pub_signo_car
caa_histo!CAA_FECHA_COBRO = LK_FECHA_DIA
caa_histo!CAA_TOTAL = Abs(WS_TOT) * pub_signo_car
caa_histo!CAA_SALDO = Nulo_Valor0(cli_llave!cli_SALDO)
caa_histo!caa_SALDO_car = wS_saldo_car
If pub_signo_car <> 0 Then
   caa_histo!CAA_SIGNO_CAJA = PUB_SECUENCIA
Else
   caa_histo!CAA_SIGNO_CAJA = 0
End If
caa_histo!CAA_SIGNO_CAJA_REAL = pub_signo_caja
caa_histo!CAA_SIGNO_CAR = pub_signo_car
caa_histo!CAA_TIPMOV = PUB_TIPMOV
caa_histo!CAA_hora = Now
caa_histo!CAA_CODUSU = LK_CODUSU
caa_histo!CAA_ESTADO = "N"
If Not cli_llave.EOF And cli_llave.RowCount > 0 Then
   If cli_llave!cli_codclie = Val(i_codcli.Text) Then
      caa_histo!CAa_NOMBRE = cli_llave!cli_nombre
   End If
End If

caa_histo!CAA_NUMPLAN = pub_numplan
If LK_CODTRA = 1111 Then caa_histo!CAA_ESTADO = "E"
If LK_CODTRA = 1122 Then caa_histo!CAA_ESTADO = "E"
If LK_CODTRA = 1133 Then caa_histo!CAA_ESTADO = "E"
If fx = 32000 Then caa_histo!CAA_ESTADO = "E"
If ws_flag_car = ingre Then
   caa_histo!CAa_NUM_CHEQUE = PUB_NUM_CHEQUE
   caa_histo!CAa_numser = PUB_NUMSER
   caa_histo!CAa_numfac = PUB_NUMFAC
   caa_histo!CAa_numser_C = PUB_NUMSER_C
   caa_histo!CAa_numfac_C = PUB_NUMFAC_C
   caa_histo!CAa_numGUIA = PUB_NUMGUIA
   caa_histo!CAa_SERGUIA = Val(i_serguia.Text)
   caa_histo!CAA_FBG = PUB_FBG
   caa_histo!CAA_CODVEN = PUB_CODVEN
   caa_histo!caa_situacion = " "
Else
   caa_histo!CAa_NUM_CHEQUE = Nulo_Valors(car_llave!CAR_NUM_CHEQUE)
   caa_histo!CAa_numser = Nulo_Valor0(car_llave!CAR_numSER)
   caa_histo!CAa_numfac = Nulo_Valor0(car_llave!CAR_numfac)
   If LK_CODTRA = 2412 Then
     caa_histo!CAa_numser_C = PUB_NUMSER
     caa_histo!CAa_numfac_C = PUB_NUMFAC
   Else
     caa_histo!CAa_numser_C = PUB_NUMSER_C
     caa_histo!CAa_numfac_C = PUB_NUMFAC_C
   End If
   If LK_CODTRA = 1122 Or LK_CODTRA = 1111 Then
      If ww_codtra_ext = 2412 Or ww_codtra_ext = 2410 Then
         caa_histo!CAa_numfac_C = ww_numdoc
         caa_histo!CAa_numser_C = ww_numser
      End If
   End If
   
   caa_histo!CAA_NOTA = PUB_FBG
   caa_histo!CAA_FBG = Nulo_Valors(car_llave!CAR_FBG)
   If LK_CODTRA = 2770 And i_codven.Visible = True Then
      caa_histo!CAA_CODVEN = PUB_CODVEN
   Else
      caa_histo!CAA_CODVEN = Nulo_Valor0(car_llave!CAR_codven)
   End If
   caa_histo!CAa_numGUIA = Nulo_Valor0(car_llave!car_numguia)
   caa_histo!CAa_SERGUIA = Nulo_Valor0(car_llave!car_SERguia)
   
   caa_histo!caa_situacion = Nulo_Valors(car_llave!CAR_SITUACION)
   caa_histo!caa_signo_ccm = pub_signo_ccm
   caa_histo!caa_codban = PUB_CODBAN
End If
caa_histo!CAA_RECIBO = PUB_CODART
If LK_CODTRA = 2728 And LK_EMP = "PLA" Then
  caa_histo!caa_situacion = PLAZA_FLAG_MANUAL
End If

caa_histo.Update
Return


manda_numero:
SQ_OPER = 2
PUB_FECHA = LK_FECHA_DIA
pu_codcia = LK_CODCIA
LEER_ALL_LLAVE
If all_menor.EOF = False Then
   PUB_NUM_OPER_XXX = all_menor!ALL_NUMOPER
Else
   PUB_NUM_OPER_XXX = 0
End If
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
WS_NUM_OPER2 = PUB_NUM_OPER_XXX
Return

End Sub

Private Sub Grid_all_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Grid_all.Visible = False

End Sub

Private Sub Grid_all_LostFocus()
Grid_all.Visible = False

End Sub

Private Sub Grid_all_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If LK_CODTRA <> 1111 Then
   GoTo fin
End If
If Button <> 2 Then
   GoTo fin
End If

If Grid_all.TextMatrix(Grid_all.Row, 11) <> LK_FECHA_DIA Then
   MsgBox "No puede Extornar Operaciones de otra fecha ..."
   GoTo fin
End If



   If Trim(Grid_all.TextMatrix(Grid_all.Row, 12)) = "E" Then
      MsgBox "Ya Extornado..."
      GoTo fin
   End If
i_num_oper.Text = Grid_all.TextMatrix(Grid_all.Row, 1)
REGRESA:
ext_conta = 0
fila = 1
Do Until fila = Grid_all.Rows
   Grid_all.COL = 1
   Grid_all.Row = fila
   If Val(i_num_oper.Text) = Grid_all.Text And Trim(Grid_all.TextMatrix(Grid_all.Row, 12)) <> "E" Then
      Grid_all.CellBackColor = vbRed
      ext_conta = ext_conta + 1
      If ext_conta = 1 Then PUB_NUM_OPER = Grid_all.TextMatrix(Grid_all.Row, 0)
   Else
      Grid_all.RowHeight(Grid_all.Row) = 0
   End If
   fila = fila + 1

 Loop
If ext_conta > 1 Then
   pub_mensaje = " estas " & ext_conta & " operaciones"
Else
   pub_mensaje = " esta Operacion"
End If
pub_mensaje = MsgBox("Esta seguro de Extornar" & pub_mensaje, 36, Pub_Titulo)

If pub_mensaje = vbNo Then
   i_num_oper.Text = 0
Else
   i_num_oper.Text = PUB_NUM_OPER
   MsgBox "Dar click en Grabar ..."
End If

NUMERO = Diario.WhatsThisHelpID
avanza_grifo

fin:

End Sub

Private Sub grid_autorizacion_GotFocus()
Dim cambiar_color As Boolean
Dim Tit As String
grid_autorizacion.Clear
grid_autorizacion.Cols = 4
grid_autorizacion.ColWidth(0) = 600
grid_autorizacion.ColWidth(1) = 3000
grid_autorizacion.ColWidth(2) = 2700
grid_autorizacion.ColWidth(3) = 1500

grid_autorizacion.TextMatrix(0, 0) = "N.Autoriz."
grid_autorizacion.TextMatrix(0, 1) = "Nombre del Cliente"
grid_autorizacion.TextMatrix(0, 2) = "Articulos ..."
grid_autorizacion.TextMatrix(0, 3) = "Concepto"

grid_autorizacion.Rows = 2
pub_autkey = 30000
pu_codcia = LK_CODCIA
PUB_FECHA = LK_FECHA_DIA
SQ_OPER = 3
LEER_AUT_LLAVE
If aut_menor.EOF = True Then
   MsgBox "Lo siento ... No tiene autorizaciones ", 48, Pub_Titulo
   grid_autorizacion.Visible = False
   GoTo fin
End If
fila = 1
Do Until aut_menor.EOF

If Trim(LK_CODUSU) = Trim(aut_menor!aut_codusu_final) And aut_menor!AUT_SECUENCIA = 1 Then

cambiar_color = False
If Nulo_Valors(aut_menor!aut_estado) = "A" Or Nulo_Valors(aut_menor!aut_estado) = "X" Then
   cambiar_color = True
End If
grid_autorizacion.Rows = grid_autorizacion.Rows + 1

grid_autorizacion.Row = fila
grid_autorizacion.COL = 0
If cambiar_color = True Then
   grid_autorizacion.CellBackColor = vb3DLight
End If
grid_autorizacion.Text = aut_menor!aut_key


grid_autorizacion.COL = 1
If cambiar_color = True Then
   grid_autorizacion.CellBackColor = vb3DLight
End If
SQ_OPER = 1
pu_cp = "C"
pu_codcia = LK_CODCIA
pu_codclie = aut_menor!aut_codclie
LEER_CLI_LLAVE

grid_autorizacion.Text = cli_llave!cli_nombre



SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
PUB_KEY = aut_menor!aut_codart
LEER_ART_LLAVE
grid_autorizacion.COL = 2
If cambiar_color = True Then
   grid_autorizacion.CellBackColor = vb3DLight
End If
grid_autorizacion.Text = Left(art_LLAVE!ART_NOMBRE, 25) & "..."


grid_autorizacion.COL = 3
If cambiar_color = True Then
   grid_autorizacion.CellBackColor = vb3DLight
End If


grid_autorizacion.Text = Nulo_Valors(aut_menor!AUT_concepto)
fila = fila + 1

End If
aut_menor.MoveNext

Loop
fin:

End Sub

Private Sub grid_autorizacion_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If grid_autorizacion.Cols <> 4 Then Exit Sub
pub_autkey = grid_autorizacion.TextMatrix(grid_autorizacion.Row, 0)

grid_autorizacion.Clear
grid_autorizacion.Cols = 7
grid_autorizacion.ColWidth(0) = 600
grid_autorizacion.ColWidth(1) = 2400
grid_autorizacion.ColWidth(2) = 0
grid_autorizacion.ColWidth(3) = 1900
grid_autorizacion.ColWidth(4) = 1200
grid_autorizacion.ColWidth(5) = 1200
grid_autorizacion.ColWidth(6) = 0

grid_autorizacion.TextMatrix(0, 0) = "N.Autoriz."
grid_autorizacion.TextMatrix(0, 1) = "Nombre del Cliente"
grid_autorizacion.TextMatrix(0, 3) = "Articulos "
grid_autorizacion.TextMatrix(0, 4) = "Cantidad"
grid_autorizacion.TextMatrix(0, 5) = "Precio"


pu_codcia = LK_CODCIA
SQ_OPER = 1
LEER_AUT_LLAVE
If aut_llave.EOF = True Then
   MsgBox "Lo siento ... no hay autorizaciones"
   grid_autorizacion.Visible = False
   GoTo fin
End If
fila = 1
Do Until aut_llave.EOF
If Nulo_Valors(aut_llave!aut_estado) = "A" Then
   MsgBox "Ya procesada... "
   grid_autorizacion.Visible = False
   Exit Do
End If
If Nulo_Valors(aut_llave!aut_estado) = "X" Then
   MsgBox "Anulada ... "
   grid_autorizacion.Visible = False
   Exit Do
End If

grid_autorizacion.Row = fila
If aut_llave!AUT_SECUENCIA = 1 Then
grid_autorizacion.COL = 0
grid_autorizacion.Text = aut_llave!aut_key
grid_autorizacion.COL = 1
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
PUB_CODCLIE = aut_llave!aut_codclie
LEER_CLI_LLAVE
grid_autorizacion.Text = cli_llave!cli_nombre
End If

grid_autorizacion.COL = 2
grid_autorizacion.Text = aut_llave!aut_codart



SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
PUB_KEY = aut_llave!aut_codart
LEER_ART_LLAVE
grid_autorizacion.COL = 3
grid_autorizacion.Text = Left(art_LLAVE!ART_NOMBRE, 35)


grid_autorizacion.COL = 4
grid_autorizacion.Text = aut_llave!aut_CANTIDAD

grid_autorizacion.COL = 5
grid_autorizacion.Text = aut_llave!aut_PRECIO

grid_autorizacion.COL = 6
grid_autorizacion.Text = aut_llave!aut_codclie


aut_llave.MoveNext

fila = fila + 1
Loop
fin:

End Sub

Private Sub grid_autorizacion_LostFocus()
grid_autorizacion.Visible = False
End Sub

Private Sub grid_fac_Click()
If LK_CODTRA = 2105 Then Frame4.Visible = True

End Sub

Private Sub grid_fac_EnterCell()
If flag_salto = 1000 Then Exit Sub
'If flag_salto = 1 Then Exit Sub
If SUT_LLAVE.EOF Then Exit Sub

If LK_CODTRA = 2210 And Val(grid_fac.TextMatrix(grid_fac.Row, 10)) = 9999 And grid_fac.COL = 1 Then Exit Sub

If LK_CODTRA = 2210 And grid_fac.COL = 3 Then Exit Sub
If LK_CODTRA = 2406 And grid_fac.COL = 1 Then Exit Sub
If LK_CODTRA = 2406 And grid_fac.COL >= 5 Then Exit Sub

If LK_CODTRA = 2212 And grid_fac.COL = 1 Then Exit Sub
If LK_CODTRA = 2125 And grid_fac.COL = 1 Then Exit Sub


If Val(grid_fac.TextMatrix(grid_fac.Row, 22)) <> 0 And LK_CODTRA = 2101 And LK_CODTRA <> 2107 Then Exit Sub

If Trim(SUT_LLAVE!SUT_abreviado) <> "CONS" Then
   If grid_fac.COL = 4 And LK_CODTRA <> 2101 And LK_CODTRA <> 2103 And LK_CODTRA <> 2211 And LK_CODTRA <> 1402 And LK_CODTRA <> 2406 Then Exit Sub
End If

If grid_fac.COL = 6 Then Exit Sub
'If grid_fac.COL = 10 Then Exit Sub

If LK_CODTRA = 2105 And Val(gridl.TextMatrix(1, 5)) <> 0 Then Exit Sub


If grid_canje.Visible = True Then
   grid_canje.Visible = False
   textovar_canje.Visible = False
End If

i_unidades.Visible = False
i_precios.Visible = False
i_mortal.Visible = False
i_is.Visible = False
If LK_CODTRA <> 2415 Then i_is.Visible = False

textovar.Visible = False
If Trim(grid_fac.TextMatrix(grid_fac.Row, 0)) <> "" Then
   label_nomart.Caption = grid_fac.TextMatrix(grid_fac.Row, 0)
   If LK_ACTIVA <> "A" Then
      label_precio.Caption = grid_fac.TextMatrix(grid_fac.Row, 13)
   Else
      If i_fbg.Text = "F" Or i_fbg.Text = "B" Then
         label_precio.Caption = grid_fac.TextMatrix(grid_fac.Row, 39)
      Else
         label_precio.Caption = grid_fac.TextMatrix(grid_fac.Row, 38)
      End If
   End If
   label_nomart.Visible = True
   label_precio.Visible = True
Else
   label_nomart.Visible = False
   label_precio.Visible = False
End If
If grid_fac.Text = "" Or (grid_fac.COL = 7 And grid_fac.Text = "") Or (grid_fac.COL = 15) Or (SUT_LLAVE!SUT_PRECIO = "2" And PUB_TIPMOV = 10 And grid_fac.COL = 6) Or grid_fac.COL = 10 Then
   procesa_celdas
Else
   grid_fac.SetFocus
End If
End Sub
Private Sub grid_fac_KeyPress(KeyAscii As Integer)
'flag_salto = 1

If KeyAscii <> 13 Then Exit Sub
If Val(grid_fac.TextMatrix(grid_fac.Row, 22)) <> 0 And LK_CODTRA = 2101 And LK_CODTRA <> 2107 Then Exit Sub

'flag_salto = 0
If LK_CODTRA = 2211 And Val(grid_fac.TextMatrix(grid_fac.Row, 10)) = 9999 And grid_fac.COL = 1 Then Exit Sub
'If LK_CODTRA = 2211 And grid_fac.COL = 3 Then Exit Sub

If LK_CODTRA = 2212 And grid_fac.COL = 1 Then Exit Sub
If LK_CODTRA = 2125 And grid_fac.COL = 1 Then Exit Sub

If LK_CODTRA = 2210 And i_origen.ListIndex = 0 Then Exit Sub

If grid_fac.COL = 4 And LK_CODTRA = 1402 And gridC.TextMatrix(gridC.Row, 7) <> "" Then Exit Sub


If Trim(SUT_LLAVE!SUT_abreviado) <> "CONS" Then
   If grid_fac.COL = 4 And LK_CODTRA <> 2101 And LK_CODTRA <> 2103 And LK_CODTRA <> 2211 And LK_CODTRA <> 2212 And LK_CODTRA <> 1402 And LK_CODTRA <> 2406 And LK_CODTRA <> 2408 Then Exit Sub
End If
'If grid_fac.COL = 10 Then Exit Sub

If LK_CODTRA = 2105 And Val(gridl.TextMatrix(1, 5)) <> 0 Then Exit Sub

If grid_fac.COL = 5 Then
Else
procesa_celdas
Exit Sub
End If

SQ_OPER = 1
PUB_KEY = Nulo_Valor0(grid_fac.TextMatrix(grid_fac.Row, 16))
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
If art_LLAVE.EOF Then Exit Sub

If grid_fac.COL = 5 And LK_CODTRA = 1402 Then
   grid_fac.TextMatrix(grid_fac.Row, 6) = ""
   verifica_unidades
   procesa_celdas
Else
   verifica_precios
   procesa_celdas
End If


End Sub

Private Sub grid_fac_KeyUp(KeyCode As Integer, Shift As Integer)
Dim fx As Integer
Dim dflag As Integer

If KeyCode = 113 Then
 grid_fac_KeyPress 13
 Exit Sub
End If

If KeyCode = 45 And grid_fac.COL = 1 Then
   grid_fac.Rows = grid_fac.Rows + 1
   Exit Sub
End If
If KeyCode = 32 And grid_fac.TextMatrix(grid_fac.Row, 1) <> "" Then
  If grid_fac.Rows = 3 Then Exit Sub
   grid_fac.COL = 1
   If grid_fac.CellBackColor = vbBlue Then
      grid_fac.CellBackColor = vbWhite
   Else
      grid_fac.CellBackColor = vbBlue
   End If
End If


If KeyCode <> 46 Then Exit Sub
   pub_mensaje = " ¿Desea Eliminar las filas seleccionadas ... ?"
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   If Pub_Respuesta = vbNo Then Exit Sub
   If grid_fac.Rows >= 3 Then
     grid_fac.RemoveItem (grid_fac.Row)
     label_nomart.Caption = ""
     label_precio.Caption = ""
   End If
   grid_fac.Refresh
   calcula_totales
   grid_fac.SetFocus
   Exit Sub
dflag = 0
OTRO:
fx = 0
FILAX = 2
grid_fac.COL = 1
Do Until FILAX > grid_fac.Rows - 1
   grid_fac.Row = FILAX
   If grid_fac.CellBackColor = vbBlue And FILAX >= 2 And grid_fac.Rows <> 3 Then
      grid_fac.RemoveItem (grid_fac.Row)
      dflag = 1
      fx = 1
      grid_fac.Refresh
      Exit Do
   End If
   FILAX = FILAX + 1
Loop

If fx = 1 Then GoTo OTRO

If dflag = 0 Then
 MsgBox "Seleccione el que desea Quitar. [con barra de espacio]", 48, Pub_Titulo
End If
If grid_fac.Rows = 3 Then
 label_nomart.Caption = ""
 label_precio.Caption = ""
End If


grid_fac.Refresh
calcula_totales
grid_fac.SetFocus

End Sub

Private Sub grid_fac_Scroll()
textovar.Visible = False
i_mortal.Visible = False
i_is.Visible = False
i_precios.Visible = False
i_unidades.Visible = False
End Sub
Private Sub grid_trans_KeyPress(KeyAscii As Integer)
Dim WW_MERMA As Currency
Dim J As Integer
Dim K As Integer
If KeyAscii <> 13 Then Exit Sub
i_ser_lote.Text = grid_trans.TextMatrix(grid_trans.Row, 10)
i_num_lote.Text = grid_trans.TextMatrix(grid_trans.Row, 1)
If LK_CODTRA = 2211 Then GoTo TR_2211
If LK_CODTRA = 2212 Then GoTo TR_2211

If Val(Right(i_origen.Text, 5)) = 1 And i_origen.Visible = True Then GoTo sal
PSFAR_TRANS.rdoParameters(0) = grid_trans.TextMatrix(grid_trans.Row, 10)
PSFAR_TRANS.rdoParameters(1) = grid_trans.TextMatrix(grid_trans.Row, 1)
If LK_CODTRA = 2212 Then
   PSFAR_TRANS.rdoParameters(2) = 4
Else
   PSFAR_TRANS.rdoParameters(2) = Val(Right(i_origen.Text, 5))
End If
   
PSFAR_TRANS.rdoParameters(3) = LK_CODCIA
FAR_TRANS.Requery
If FAR_TRANS.EOF Then
   MsgBox "No hay datos en el Origen: " & i_origen.Text
   Exit Sub
End If
fila = 0
Do Until FAR_TRANS.EOF
'    WW_MERMA = (FAR_TRANS!FAR_CANTIDAD - FAR_TRANS!FAR_CONSIG) / FAR_TRANS!FAR_CANTIDAD
'    If WW_MERMA < 0.05 Then GoTo OTRO
   
   grid_fac.Rows = fila + 2
   PUB_KEY = FAR_TRANS!far_codart
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ART_LLAVE
   
   
   SQ_OPER = 1
   PUB_CODART = PUB_KEY
   pu_codcia = LK_CODCIA
   LEER_ARM_LLAVE
   If arm_llave.EOF Then
      MsgBox "Datos de Articulos ...errados..Revisar"
      Exit Sub
   End If
   
         
   grid_fac.TextMatrix(fila, 0) = art_LLAVE!ART_NOMBRE
   grid_fac.TextMatrix(fila, 1) = art_LLAVE!ART_ALTERNO
   grid_fac.TextMatrix(fila, 16) = art_LLAVE!ART_KEY
   grid_fac.TextMatrix(fila, 33) = art_LLAVE!ART_KEY
   
   grid_fac.TextMatrix(fila, 2) = 0
   
   grid_fac.TextMatrix(fila, 3) = 0

   grid_fac.TextMatrix(fila, 11) = arm_llave!arm_cospro
   grid_fac.TextMatrix(fila, 4) = ""
   If LK_CODTRA = 2212 Then
      grid_fac.TextMatrix(fila, 2) = FAR_TRANS(2)
      grid_fac.TextMatrix(fila, 4) = FAR_TRANS(1)
   Else
      grid_fac.TextMatrix(fila, 3) = FAR_TRANS(1)
   End If
   
   If LK_CODTRA = 2212 Then
      pu_codcia = LK_CODCIA
      PUB_CODART = art_LLAVE!ART_KEY
      LEER_PRE_LLAVE
      grid_fac.TextMatrix(fila, 6) = pre_llave!PRE_PRE1
      grid_fac.TextMatrix(fila, 36) = pre_llave!PRE_PRE1
      grid_fac.TextMatrix(fila, 35) = pre_llave!PRE_PRE1
      grid_fac.TextMatrix(fila, 33) = PUB_CODART
   Else
      grid_fac.TextMatrix(fila, 6) = arm_llave!arm_cospro
   End If
      
   grid_fac.TextMatrix(fila, 14) = 1
   grid_fac.TextMatrix(fila, 5) = " "
   grid_fac.TextMatrix(fila, 8) = 0
   grid_fac.TextMatrix(fila, 10) = 9999
   If LK_CODTRA = 2210 Then
      flag_salto = 1000
      grid_fac.COL = 4
      grid_fac.Row = fila
      flag_salto = 0
      grid_fac.CellBackColor = vbRed
      grid_fac.TextMatrix(fila, 12) = -1
   Else
      grid_fac.TextMatrix(fila, 12) = pub_signo_arm
   End If
   
   grid_fac.TextMatrix(fila, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
   grid_fac.TextMatrix(fila, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
   grid_fac.TextMatrix(fila, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
   
'   grid_fac.TextMatrix(fila, 28) = ped_llave!PED_NUMSER
'   grid_fac.TextMatrix(fila, 30) = ped_llave!PED_NUMFAC
'   grid_fac.TextMatrix(fila, 31) = ped_llave!ped_numsec
   
   fila = fila + 1
OTRO:
   FAR_TRANS.MoveNext
Loop
calcula_totales
sal:
grid_trans.Visible = False
Frame4.Visible = True
label_grid.Visible = False
grid_fac.SetFocus
Exit Sub

TR_2211:

If grid_trans.TextMatrix(grid_trans.Row, 1) = "" Then
   Exit Sub
End If
If LK_CODTRA = 2211 And PUB_SECUENCIA = 1 Then
   grid_fac.ColWidth(5) = 0
End If
fila = 1
 For J = 1 To grid_trans.Rows - 1
   If Val(grid_trans.TextMatrix(J, 10)) = Val(i_ser_lote.Text) And Val(grid_trans.TextMatrix(J, 1)) = Val(i_num_lote.Text) Then
   Else
      GoTo PASE
   End If
 
   If LK_CODTRA = 2212 Then
         PSFAR_TRANS2.rdoParameters(0) = grid_trans.TextMatrix(grid_trans.Row, 1)
         PSFAR_TRANS2.rdoParameters(1) = SUT_LLAVE!sub_art_gru
         PSFAR_TRANS2.rdoParameters(2) = LK_CODCIA
         FAR_TRANS2.Requery
         PUB_CANTIDAD = 0
         PUB_UNIDAD = 0
        
         If FAR_TRANS2.RowCount > 0 Then FAR_TRANS2.MoveFirst
         Do Until FAR_TRANS2.EOF
            If Val(FAR_TRANS2!far_codart) <> Val(grid_trans.TextMatrix(J, 5)) Then GoTo NEXTO2
           If LK_CODTRA = 2211 Then
           If FAR_TRANS2(2) > 0 Then
              PUB_CANTIDAD = FAR_TRANS2(2)
              PUB_UNIDAD = FAR_TRANS2(1)
           End If
           End If
           If LK_CODTRA = 2212 Then
              PUB_CANTIDAD = FAR_TRANS2(2) + PUB_CANTIDAD
              PUB_UNIDAD = FAR_TRANS2(1) + PUB_UNIDAD
           End If
           
NEXTO2:
           FAR_TRANS2.MoveNext
        Loop
Else
        PSFAR_TRANS.rdoParameters(0) = grid_trans.TextMatrix(grid_trans.Row, 10)
        PSFAR_TRANS.rdoParameters(1) = grid_trans.TextMatrix(grid_trans.Row, 1)
        If PUB_TIPMOV = 181 Then
         PSFAR_TRANS.rdoParameters(2) = Nulo_Valor0(SUT_LLAVE!SUT_art_gru)
        Else
         PSFAR_TRANS.rdoParameters(2) = 4
        End If
        PSFAR_TRANS.rdoParameters(3) = LK_CODCIA
        FAR_TRANS.Requery

        PUB_CANTIDAD = 0
        PUB_UNIDAD = 0
        
        If FAR_TRANS.RowCount > 0 Then FAR_TRANS.MoveFirst
        Do Until FAR_TRANS.EOF
           If Val(FAR_TRANS!far_codart) <> Val(grid_trans.TextMatrix(J, 5)) Then GoTo NEXTO
           If LK_CODTRA = 2211 Then
           If FAR_TRANS(2) > 0 Then
              PUB_CANTIDAD = FAR_TRANS(2)
              PUB_UNIDAD = FAR_TRANS(1)
           End If
           End If
           If LK_CODTRA = 2212 Then
              PUB_CANTIDAD = FAR_TRANS(2) + PUB_CANTIDAD
              PUB_UNIDAD = FAR_TRANS(1) + PUB_UNIDAD
           End If
           
NEXTO:
           FAR_TRANS.MoveNext
        Loop

End If

   If PUB_CANTIDAD = 0 And PUB_TIPMOV <> 180 And PUB_TIPMOV <> 181 Then GoTo PASE
   fila = fila + 1
   If i_fbg.Text = "F" Then
     If fila = par_llave!par_fac_lines + 2 Then
         MsgBox "Llego al tope de las Lineas de Facturas." & Chr(13) & "Max. de : " & par_llave!par_fac_lines, 48, Pub_Titulo
         GoTo PASE
     End If
   ElseIf i_fbg.Text = "B" Then
     If fila = par_llave!par_BOL_lines + 2 Then
         MsgBox "Llego al tope de las Lineas de Boletas." & Chr(13) & "Max. de : " & par_llave!par_BOL_lines, 48, Pub_Titulo
         GoTo PASE
     End If
   End If
   grid_fac.Rows = fila + 2
   PUB_KEY = Val(grid_trans.TextMatrix(J, 5))
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ART_LLAVE
   
   
   SQ_OPER = 1
   PUB_CODART = PUB_KEY
   pu_codcia = LK_CODCIA
   LEER_ARM_LLAVE
   If arm_llave.EOF Then
      MsgBox "Datos de Articulos ...errados..Revisar"
      Exit Sub
   End If
   If PUB_TIPMOV = 181 Then
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
     PUB_CODVEN = grid_trans.TextMatrix(J, 13)
     LEER_VEN_LLAVE
     grid_fac.TextMatrix(1, 0) = Format(ven_llave!vem_codven, "00") & " - " & Trim(ven_llave!VEM_NOMBRE)
   End If
   grid_fac.TextMatrix(fila, 0) = art_LLAVE!ART_NOMBRE
   grid_fac.TextMatrix(fila, 1) = art_LLAVE!ART_ALTERNO
   grid_fac.TextMatrix(fila, 16) = art_LLAVE!ART_KEY
   grid_fac.TextMatrix(fila, 33) = art_LLAVE!ART_KEY
   
   grid_fac.TextMatrix(fila, 2) = 0
   grid_fac.TextMatrix(fila, 3) = 0

   grid_fac.TextMatrix(fila, 11) = arm_llave!arm_cospro
   grid_fac.TextMatrix(fila, 4) = ""
   pu_codcia = LK_CODCIA
   PUB_CODART = art_LLAVE!ART_KEY
   LEER_PRE_LLAVE
   If LK_CODTRA = 2211 Then
     If PUB_SECUENCIA <> 0 Then
       grid_fac.TextMatrix(fila, 2) = grid_trans.TextMatrix(J, 3) - PUB_UNIDAD
       grid_fac.TextMatrix(fila, 4) = grid_trans.TextMatrix(J, 4) - PUB_CANTIDAD
       grid_fac.TextMatrix(fila, 6) = grid_trans.TextMatrix(J, 4) - PUB_CANTIDAD
      Else
                                    ' CAMTEX
       grid_fac.TextMatrix(fila, 2) = 0 'grid_trans.TextMatrix(J, 3) - PUB_UNIDAD
                                        ' CAMTEX
       grid_fac.TextMatrix(fila, 4) = 0 'grid_trans.TextMatrix(J, 4) - PUB_CANTIDAD
       grid_fac.TextMatrix(fila, 6) = arm_llave!arm_cospro
      End If
   ElseIf LK_CODTRA = 2212 Then
      grid_fac.TextMatrix(fila, 2) = PUB_UNIDAD
      grid_fac.TextMatrix(fila, 4) = PUB_CANTIDAD
      grid_fac.TextMatrix(fila, 6) = pre_llave!PRE_PRE1
   End If
   grid_fac.TextMatrix(fila, 5) = pre_llave!pre_UNIDAD
   grid_fac.TextMatrix(fila, 14) = 1
   grid_fac.TextMatrix(fila, 8) = 0
   grid_fac.TextMatrix(fila, 10) = 9999
   grid_fac.TextMatrix(fila, 12) = pub_signo_arm
   grid_fac.TextMatrix(fila, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
   grid_fac.TextMatrix(fila, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
   grid_fac.TextMatrix(fila, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
   
   grid_fac.TextMatrix(fila, 28) = grid_trans.TextMatrix(J, 10)
   grid_fac.TextMatrix(fila, 30) = grid_trans.TextMatrix(J, 1)
   grid_fac.TextMatrix(fila, 31) = grid_trans.TextMatrix(J, 6)
   
   grid_fac.TextMatrix(fila, 39) = grid_trans.TextMatrix(J, 12)

PASE:
Next J
calcula_totales
If LK_CODTRA = 2211 And PUB_SECUENCIA = 1 Then grid_fac.TextMatrix(0, 6) = "BRUTO"
grid_trans.Visible = False
Frame4.Visible = True
label_grid.Visible = False
grid_fac.SetFocus







End Sub

Private Sub gridl_Click()
If LK_CODTRA = 2105 Then Frame4.Visible = False
End Sub


Private Sub gridl_KeyUp(KeyCode As Integer, Shift As Integer)
Dim FILAX, fx, FILA_aCT As Integer





If KeyCode = 46 Then
If gridl.Rows <= 3 Then
Else
   gridl.RemoveItem (gridl.Row)
   gridl.Refresh
   calcula_totales2
   Exit Sub
 End If
End If


   


    
End Sub

Private Sub gridl_Scroll()

textovarl.Height = 30
textovarl.Width = 30
End Sub

Private Sub GridT_EnterCell()
If SUT_LLAVE.EOF Then Exit Sub

textovar2.Visible = False
If GridT.Text = "" Then
   PROCESA_CELDAS2
End If

End Sub

Private Sub GridT_KeyPress(KeyAscii As Integer)
If GridT.COL = 6 Then
   If GridT.Text = "OK" Then
      GridT.Text = " "
   Else
      GridT.Text = "OK"
      flag_textil = 8
   End If
End If


'flag_salto = 1

If KeyAscii <> 13 Then Exit Sub

'flag_salto = 0


If grid_fac.COL = 1 Then
Else
PROCESA_CELDAS2
Exit Sub
End If

'SQ_OPER = 1
'PUB_KEY = Nulo_Valor0(GridT.TextMatrix(GridT.Row, 4))
'pu_codcia = LK_CODCIA
'LEER_ART_LLAVE
'If art_LLAVE.EOF Then Exit Sub

PROCESA_CELDAS2

End Sub

Private Sub GridT_KeyUp(KeyCode As Integer, Shift As Integer)
Dim fx As Integer
Dim dflag As Integer

If KeyCode = 113 Then
 GridT_KeyPress 13
 Exit Sub
End If

If KeyCode = 45 And GridT.COL = 1 Then
   GridT.Rows = GridT.Rows + 1
   Exit Sub
End If
If KeyCode = 32 And GridT.TextMatrix(GridT.Row, 1) <> "" Then
  If GridT.Rows = 3 Then Exit Sub
   GridT.COL = 1
   If GridT.CellBackColor = vbBlue Then
      GridT.CellBackColor = vbWhite
   Else
      GridT.CellBackColor = vbBlue
   End If
End If


If KeyCode <> 46 Then Exit Sub

   pub_mensaje = " ¿Desea Eliminar las filas seleccionadas ... ?"
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   If Pub_Respuesta = vbNo Then Exit Sub
   If GridT.Rows >= 3 Then
     GridT.RemoveItem (GridT.Row)
   End If
   GridT.Refresh
   calcula_totales22
   GridT.SetFocus
   Exit Sub
dflag = 0
OTRO:
fx = 0
FILAX = 2
GridT.COL = 1
Do Until FILAX > GridT.Rows - 1
   GridT.Row = FILAX
   If GridT.CellBackColor = vbBlue And FILAX >= 2 And GridT.Rows <> 3 Then
      GridT.RemoveItem (GridT.Row)
      dflag = 1
      fx = 1
      GridT.Refresh
      Exit Do
   End If
   FILAX = FILAX + 1
Loop

If fx = 1 Then GoTo OTRO

If dflag = 0 Then
 MsgBox "Seleccione el que desea Quitar. [con barra de espacio]", 48, Pub_Titulo
End If

GridT.Refresh
calcula_totales22
GridT.SetFocus

End Sub

Private Sub GridT_Scroll()
textovar2.Visible = False
End Sub

Private Sub i_cambio_Click()
If i_cambio.Value = 1 Then
   i_num_lote.Locked = False
   i_ser_lote.Locked = False
   i_numfac.Locked = False
   i_numser.Locked = False
   i_ser_lote.BackColor = QBColor(15)
   i_num_lote.BackColor = QBColor(15)

   i_numfac.BackColor = QBColor(15)
   i_numser.BackColor = QBColor(15)
   
   If i_numfac.Visible Then
    Azul i_numfac, i_numfac
   Else
    Azul i_num_lote, i_num_lote
   End If
   i_fecha_compra.Visible = True
   
Else
 If LK_CODTRA = 2103 Then llena_pedido
   'BolFac_Click
End If

End Sub

Private Sub i_codcli_Change()
If i_codcli.Text = "" Then
  i_nomCLI.Caption = ""
  LV_CLI.Visible = False
  BLOQ(0).Value = 0
  BLOQ(1).Value = 0
  BLOQ(2).Value = 0
  BLOQ(3).Value = 0
End If

End Sub

Private Sub i_codcli_GotFocus()
If LK_CODTRA = 2101 Then i_placas.Visible = True
loc_key = 0

End Sub

Private Sub i_codcli_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not LV_CLI.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And i_codcli.Text = "" Then
  loc_key = 1
  Set LV_CLI.SelectedItem = LV_CLI.ListItems(loc_key)
'  LV_CLI.Visible = False
  LV_CLI.ListItems.Item(loc_key).Selected = True
  LV_CLI.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_CLI.ListItems.count Then loc_key = LV_CLI.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_CLI.ListItems.count Then loc_key = LV_CLI.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
'  KeyCode = 0
  LV_CLI.ListItems.Item(loc_key).Selected = True
  LV_CLI.ListItems.Item(loc_key).EnsureVisible
  i_codcli.Text = Trim(LV_CLI.ListItems.Item(loc_key).Text) & " "
  DoEvents
  i_codcli.SelStart = Len(i_codcli.Text)
  DoEvents
fin:

End Sub

Private Sub i_codcli_KeyUp(KeyCode As Integer, Shift As Integer)
Dim VAR

If Len(i_codcli.Text) = 0 Or IsNumeric(i_codcli.Text) Then
   LV_CLI.Visible = False
   Exit Sub
End If
If LV_CLI.Visible = False Or Len(Trim(i_codcli.Text)) = 1 Then
   loc_key = 0
    VAR = Asc(i_codcli.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    numarchi = 1
    archi = "SELECT CLI_CODCLIE , CLI_CODCIA, CLI_CP, CLI_NOMBRE, CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM   FROM CLIENTES WHERE CLI_CP = '" & PUB_CP & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & i_codcli.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
    PROC_LISVIEW LV_CLI
    loc_key = 0
    If LV_CLI.Visible Then
     loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If LV_CLI.Visible Then
  Set itmFound = LV_CLI.FindItem(LTrim(i_codcli.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > LV_CLI.ListItems.count Then
      LV_CLI.ListItems.Item(LV_CLI.ListItems.count).EnsureVisible
   Else
     LV_CLI.ListItems.Item(loc_key + 8).EnsureVisible
   End If
  End If
  Exit Sub
End If

End Sub
Private Sub i_codcli_LostFocus()
If Not IsNumeric(i_codcli.Text) Then
 Exit Sub
End If
Dim ww_num As Integer

If cli_llave.EOF = False And Val(i_codcli.Text) > 0 Then
   If cli_llave!cli_codclie <> Val(i_codcli.Text) Then
      i_codcli.Text = ""
      Exit Sub
   End If
End If

If cli_llave.EOF = False And Val(i_codcli.Text) > 0 Then
   ww_num = Val(Nulo_Valors(cli_llave!CLI_nucleo))
   If PUB_TIPMOV = 10 And ww_num <> 0 Then
      ' LIMPIA PRECIOS.
    For fila = 2 To grid_fac.Rows - 1
       grid_fac.TextMatrix(fila, 6) = "0.00"
    Next fila
    End If
End If

If cli_llave.EOF = False And Val(i_codcli.Text) > 0 Then
   If cli_llave!cli_codclie = Val(i_codcli.Text) Then
      If LK_CODTRA = 2105 Then
          i_dias.Text = Nulo_Valor0(cli_llave!cli_DIAS_CRED)
          If IsDate(i_fecha_compra.Text) Then i_fecha_vcto.Text = Str(DateAdd("d", Val(i_dias.Text), i_fecha_compra.Text))
      End If
      If Left(SUT_LLAVE!sut_descto, 1) = "P" And LK_CODTRA = 2101 And Nulo_Valor0(cli_llave!CLI_PORDESCTO) <> 0 Then MsgBox "Cliente tiene Descto de : " & Format(cli_llave!CLI_PORDESCTO, "0.00") & "%", 48, Pub_Titulo
'      i_limcre_ant.Text = Nulo_Valor0(cli_llave!cli_limcre)
   End If
End If

If PUB_TIPMOV = 10 And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 0 Then
If Not cli_llave.EOF Then
   If SUT_LLAVE!SUT_SIGNO_CAR = 1 Then
      pu_codcia = LK_CODCIA
      pub_deuda = CAR_TOT_CPX2("C", pu_codcia, cli_llave!cli_codclie)
      If PUB_FLAG_VENCIDO = 1 Then
         MsgBox "CLIENTE TIENE OBLIGACIONES VENCIDAS "
         MsgBox pub_mensaje, 48, Pub_Titulo
      End If
      If Nulo_Valors(cli_llave!CLI_TIPO_BLOQ1) = "1" Then
         MsgBox "Credito Bloqueado ..."
         Exit Sub
      End If
   End If
End If
End If

'grid_fac.Clear
'grid_fac.Rows = 3
'pasa_cabeza
LV_CLI.Visible = False
If LK_CODTRA = 2105 Then
   If Nulo_Valor0(cli_llave!cli_DIAS_CRED) = 0 Then MsgBox "Dias de Credito es 0 ", 48, Pub_Titulo
'   i_fbg.ListIndex = 1
   Option4.Visible = True
   Option4.Enabled = True
   'Option4.SetFocus
End If

If LK_CODTRA = 2101 Then
i_placas.Clear
PS_PLACAS(0) = Val(i_codcli.Text)
PS_PLACAS(1) = LK_CODCIA
TAB_PLACAS.Requery
Do Until TAB_PLACAS.EOF
   i_placas.AddItem TAB_PLACAS!TAB_NOMLARGO
   DoEvents
   TAB_PLACAS.MoveNext
Loop
'If TAB_PLACAS.RowCount = 0 Then
'   i_placas.Visible = False
'   grid_fac.SetFocus
'End If
End If

End Sub
Private Sub i_codven_Change()

If i_codven.Text = "" Then i_nomven.Caption = ""
If i_cambio.Value = 1 Then Exit Sub
If LK_CODTRA = 2406 Then Exit Sub
i_numguia.Text = ""
i_numser.Text = ""
i_numfac.Text = ""

End Sub

Private Sub i_codven_GotFocus()
i_nomven2.Caption = ""
i_turno.Text = ""
i_nomturno.Caption = ""

Azul i_codven, i_codven
End Sub

Private Sub i_codven_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.

If Not LV_VEN.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And i_codven.Text = "" Then
  loc_key = 1
  Set LV_VEN.SelectedItem = LV_VEN.ListItems(loc_key)
  LV_VEN.ListItems.Item(loc_key).Selected = True
  LV_VEN.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_VEN.ListItems.count Then loc_key = LV_VEN.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_VEN.ListItems.count Then loc_key = LV_VEN.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
'  KeyCode = 0
  LV_VEN.ListItems.Item(loc_key).Selected = True
  LV_VEN.ListItems.Item(loc_key).EnsureVisible
  i_codven.Text = Trim(LV_VEN.ListItems.Item(loc_key).Text) & " "
  DoEvents
  i_codven.SelStart = Len(i_codven.Text)
  DoEvents
fin:

End Sub
Private Sub i_codven_KeyUp(KeyCode As Integer, Shift As Integer)
Dim VAR

If Len(i_codven.Text) = 0 Or IsNumeric(i_codven.Text) Then
   LV_VEN.Visible = False
   Exit Sub
End If
If LV_VEN.Visible = False Or Len(i_codven.Text) = 1 Then
    loc_key = 0
    VAR = Asc(i_codven.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    numarchi = 2
    archi = "SELECT VEM_CODVEN , VEM_CODCIA, VEM_NOMBRE  FROM VEMAEST WHERE  VEM_CODCIA = '" & LK_CODCIA & "' AND VEM_NOMBRE BETWEEN '" & i_codven.Text & "' AND  '" & VAR & "' ORDER BY VEM_NOMBRE"
    PROC_LISVIEW LV_VEN
    loc_key = 0
    If LV_VEN.Visible Then
     loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If LV_VEN.Visible Then
  Set itmFound = LV_VEN.FindItem(LTrim(i_codven.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > LV_VEN.ListItems.count Then
      LV_VEN.ListItems.Item(LV_VEN.ListItems.count).EnsureVisible
   Else
     LV_VEN.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If


End Sub

Private Sub i_codven_LostFocus()
Dim i
If i_codven.Text <> "" And ven_llave.EOF Then
   MsgBox "Indicar Vendedor o Isla..."
   i_codven.Text = ""
   i_codven.SetFocus
   Exit Sub
End If

If Not ven_llave.EOF Then
If ven_llave!vem_codven <> Val(i_codven.Text) And Val(i_codven.Text) <> 0 Then
   MsgBox "Falta indicar la Isla..."
   i_codven.Text = ""
   i_codven.SetFocus
   Exit Sub
  End If
End If

i_responsable.ListIndex = -1

End Sub

Private Sub i_def_Click()
i_numfac.Text = ""
i_numser.Text = ""
End Sub

Private Sub i_descto_Click()
If PUB_TIPMOV = 20 Then i_descto.Locked = True
End Sub

Private Sub i_descto_GotFocus()
Azul i_descto, i_descto
i_descto.MaxLength = 13
End Sub

Private Sub i_descto_LostFocus()
'calcula_totales
End Sub

Private Sub i_dias_Change()
If grabar.Enabled = False Then Exit Sub
If IsDate(i_fecha_compra) = False Then Exit Sub
FORM_GRIFO.i_fecha_vcto.Text = Str(DateAdd("d", Val(FORM_GRIFO.i_dias.Text), FORM_GRIFO.i_fecha_compra.Text))
On Error GoTo SALE

Exit Sub
SALE:

End Sub

Private Sub i_dias_GotFocus()
If PUB_TIPMOV = 10 And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 1 And grid_fac.Visible Then
   i_dias.Locked = True
   Exit Sub
ElseIf PUB_TIPMOV = 10 And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) <> 1 And Nulo_Valor0(SUT_LLAVE!SUT_SIGNO_CAR) = 0 And grid_fac.Visible Then
    i_dias.Locked = True
    Exit Sub
End If
i_dias.Locked = False
i_dias.MaxLength = 3
Azul i_dias, i_dias
End Sub
Private Sub i_ds_Click()
If i_ds.Text = "S" Then
   i_moneda.Text = "S/."
   WS_MONEDA_CLI = "S"
Else
   WS_MONEDA_CLI = "D"
   i_moneda.Text = " $"
End If

End Sub

Private Sub i_ds_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   Exit Sub
End If

NUMERO = i_ds.WhatsThisHelpID
avanza_grifo

End Sub
Private Sub i_fbg_Click()
   
If i_numfac.Visible = True And (LK_CODTRA = 2105 Or LK_CODTRA = 2107 Or LK_CODTRA = 2101) Then
Screen.MousePointer = 11
BolFac_Click
Screen.MousePointer = 0
End If
End Sub

Private Sub i_fbg_GotFocus()
textovarl.Height = 30
textovarl.Width = 30
If gridl.Top = grid_canje.Top Then grid_canje.Visible = False
End Sub

Private Sub i_fecha_compra_Change()
If IsDate(i_fecha_compra.Text) Then
  If i_dias.Visible Then
    i_dias_Change
   End If
End If

End Sub

Private Sub i_fecha_compra_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If LK_CODTRA = 2101 Or LK_CODTRA = 2105 Then
 If i_codcli.Visible Then i_codcli.SetFocus
End If
If LK_CODTRA = 2103 Then
    If Text1(1).Visible = True Then Text1(1).SetFocus
End If

End Sub

Private Sub i_fecha_vcto_GotFocus()
Azul i_fecha_vcto, i_fecha_vcto
End Sub

Private Sub i_flete_GotFocus()
i_flete.MaxLength = 13
End Sub

Private Sub i_flete_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_flete, KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If
PUB_FLETE = Val(i_flete.Text)

If LK_EMP = "HER" Then ' QUITAR DESPUES
   calcula_totales
End If

If LK_EMP = "HER" Then Exit Sub 'QUITAR
fin:

End Sub

Private Sub i_flete_LostFocus()
If LK_EMP = "HER" Then ' QUITAR DESPUES
   calcula_totales
End If

End Sub

Private Sub i_gastos_Change()
PUB_GASTOS = Val(i_gastos.Text)
calcula_totales
End Sub

Private Sub i_gastos_GotFocus()
i_gastos.MaxLength = 13
Azul i_gastos, i_gastos
End Sub

Private Sub i_importe_amort_GotFocus()
Azul i_importe_amort, i_importe_amort
End Sub

Private Sub i_impto_Change()
   i_neto.Text = Val(i_subtotal.Text) + Val(i_gastos.Text) - Val(i_descto.Text) + Val(i_impto.Text)

End Sub

Private Sub i_impto_GotFocus()
PUB_IMPTO = i_impto.Text
End Sub

Private Sub i_impto_LostFocus()
If (Abs(Val(i_impto.Text) - PUB_IMPTO) * 100) / (PUB_IMPTO) > 2 Then
   MsgBox "No Procede...demasiada diferencia"
   i_impto.Text = PUB_IMPTO
End If
End Sub


Private Sub i_is_Click()
If grid_fac.COL = 12 Then grid_fac.Text = i_is.Text


If Val(i_is.Text) = -1 Then
   If Val(grid_fac.TextMatrix(grid_fac.Row, 4)) > Val(grid_fac.TextMatrix(grid_fac.Row, 13)) Then
      MsgBox "Ojo stock insuficiente ..."
      If LK_USU_STOCK <> "A" Then
         grid_fac.TextMatrix(grid_fac.Row, 4) = ""
         grid_fac.TextMatrix(grid_fac.Row, 5) = ""
         grid_fac.TextMatrix(grid_fac.Row, 6) = ""
         grid_fac.TextMatrix(grid_fac.Row, 7) = ""
         i_is.ListIndex = -1
         Exit Sub
      End If
   End If
End If

End Sub

Private Sub i_is_GotFocus()

'Exit Sub
textovar.Visible = False
'i_is.Visible = True
If Val(grid_fac.Text) <> 0 Then
   If Val(grid_fac.Text) = 1 Then
      i_is.ListIndex = 0
   Else
      i_is.ListIndex = 1
   End If
Else
   grid_fac.Text = 1
   i_is.ListIndex = 0
End If

End Sub

Private Sub i_is_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
   
If grid_fac.TextMatrix(grid_fac.Row, 16) = "" Then
   grid_fac.COL = tab_derecha(grid_fac.COL)
   Exit Sub
End If
   
flag_salto = 1
If grid_fac.Row = grid_fac.Rows - 1 Then
   grid_fac.Rows = grid_fac.Rows + 1
End If
grid_fac.SetFocus
grid_fac.Row = grid_fac.Row + 1
grid_fac.RowHeight(grid_fac.Row) = 315
flag_salto = 0

If grid_fac.COL = 12 Then i_is.Visible = False
grid_fac.COL = tab_derecha(grid_fac.COL)

fin:


End Sub

Private Sub i_is_LostFocus()
i_is.Visible = False

End Sub

Private Sub i_limcre_GotFocus()
Azul i_limcre, i_limcre
End Sub

Private Sub i_mortal_Click()
grid_fac.Text = i_mortal.Text & "     " & i_mortal.ListIndex
End Sub

Private Sub i_mortal_GotFocus()
textovar.Visible = False
i_mortal.Visible = True
If grid_fac.COL <> 9 Then Exit Sub
If Len(grid_fac.Text) <> 0 Then
   i_mortal.ListIndex = Right(grid_fac.Text, 5)
Else
   grid_fac.Text = i_mortal.Text & "     " & i_mortal.ListIndex
   i_mortal.ListIndex = 0
End If

End Sub
Private Sub i_mortal_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
   
If grid_fac.TextMatrix(grid_fac.Row, 16) = "" Then
   grid_fac.COL = tab_derecha(grid_fac.COL)
   Exit Sub
End If
   
flag_salto = 1
If grid_fac.Row = grid_fac.Rows - 1 Then
   grid_fac.Rows = grid_fac.Rows + 1
End If
grid_fac.SetFocus
grid_fac.Row = grid_fac.Row + 1
grid_fac.RowHeight(grid_fac.Row) = 315
flag_salto = 0

grid_fac.COL = tab_derecha(grid_fac.COL)

fin:

End Sub

Private Sub i_mortal_LostFocus()
i_mortal.Visible = False
End Sub

Private Sub i_neto_Change()
If PUB_TIPMOV = 20 And Check1.Value = 1 And Val(i_neto.Text) <> 0 And pub_flag_cambio <> 99 Then
   pub_flag_cambio = 1
End If

End Sub

Private Sub i_neto_GotFocus()
i_neto.MaxLength = 17
Azul i_neto, i_neto
End Sub



Private Sub i_numfac_c_GotFocus()
textovarl.Height = 30
textovarl.Width = 30
Azul i_numfac_c, i_numfac_c
End Sub

Private Sub i_numfac_Change()
'If Val(i_numfac.Text) < pub_numfac_ult And i_cambio.Value = 1 And i_cambio.Visible Then
'    i_fecha_compra.Visible = True
'Else
'   i_fecha_compra.Visible = False
'End If

End Sub

Private Sub i_numfac_GotFocus()
If i_cambio.Value = 1 Then
    i_numfac.Locked = False
    i_numser.Locked = False
    i_numfac.BackColor = QBColor(15)
    i_numser.BackColor = QBColor(15)
    Azul i_numfac, i_numfac
Else
   BolFac_Click
End If

End Sub
Private Sub i_numfac_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   If i_cambio.Value = 1 Then
      i_cambio.Value = 0
      
   Else
      i_cambio.Value = 1
      i_numser.SetFocus
   End If
   If i_cambio.Value = 1 Then
   i_numfac.Locked = False
   i_numser.Locked = False
   i_numfac.BackColor = QBColor(15)
   i_numser.BackColor = QBColor(15)
'   Azul i_numfac, i_numfac
   i_fecha_compra.Visible = True
   Else
      i_numfac.Locked = True
      i_numser.Locked = True
      i_numfac.BackColor = QBColor(7)
      i_numser.BackColor = QBColor(7)
      llena_numfac_grifo
    End If

End If

End Sub
Private Sub i_numfac_LostFocus()

Dim WF As String * 1
Dim ww_flag As String * 1
Dim WS_FECHITA As Date
If i_cambio.Value = 0 Then Exit Sub
WF = ""
ww_flag = "N"
'If Not cli_llave.EOF Then If Trim(Nulo_Valors(cli_llave!cli_cia_ref)) <> "" Then WF = "A"
pub_flag_cambio = 99
If Val(i_numfac.Text) = 0 Then Exit Sub
If LK_CODTRA = 1122 Then Exit Sub
Dim fin As Integer
Dim WS_NUMFAC, WS_NUMOPER

If PUB_TIPMOV = 0 And LK_CODTRA <> 1403 Then Exit Sub

If PUB_TIPMOV = 20 And i_cambio.Value = 0 Then Exit Sub

If Not i_cambio.Visible Then Exit Sub

PU_NUMFAC = Val(i_numfac.Text)
pu_codcia = LK_CODCIA
PU_NUMSER = Val(i_numser.Text)
If i_fbg.Visible = False Or PUB_TIPMOV = 20 Then
   PUB_FBG = ""
Else
   PUB_FBG = i_fbg.Text
End If
If PUB_TIPMOV = 97 Then PUB_FBG = "N"
If PUB_TIPMOV = 98 Then PUB_FBG = "D"
PU_TIPMOV = PUB_TIPMOV
PU_FBG = PUB_FBG
If LK_EMP = "PIU" And i_cambio.Value = 1 Then
  If PUB_TIPMOV = 5 Or PUB_TIPMOV = 100 Or PUB_TIPMOV = 101 Then
    PU_TIPMOV = 20
  End If
End If


'If WF = "A" Then
'   pu_codcia = cli_llave!cli_cia_ref
'   PU_TIPMOV = 10
'   PU_FBG = i_fbg.text
'End If

SQ_OPER = 1
LEER_FAR_LLAVE
If far_llave.EOF = True Then
pub_mensaje = "NO Existe datos en este Documento ...Desea saltar a este numero? "
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   BolFac_Click
End If
'grid_fac.SetFocus
Exit Sub
End If


far_llave.MoveLast
'If far_llave!far_estado <> "E" And LK_CODTRA <> 1401 And LK_CODTRA <> 2414 Then
'  MsgBox "Nª. Documento: " + i_numser.Text + " - " + i_numfac.Text + "  - está  Emitido...no procede", 48, Pub_Titulo
'  i_numfac.Text = ""
'  i_numfac.SetFocus
'  i_cambio.Value = 0
'  BolFac_Click
'  Exit Sub
'End If


pub_mensaje = "Existe datos en este Documento : " & Trim(i_numser.Text) & " - " & i_numfac.Text & " ¿Desea Reemplazar sus datos por el actual ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
'   BolFac_Click
    If i_fecha_compra.Visible Then Azul i_fecha_compra, i_fecha_compra
   Exit Sub
End If

grid_fac.Clear
grid_fac.Rows = 3
pasa_cabeza
fin = 0
far_llave.MoveLast
WS_NUMOPER = far_llave!FAR_NUMOPER
far_llave.MoveFirst

Do Until fin = 1
'   If LK_CODTRA = 1401 And far_llave!far_ESTADO = "E" Then GoTo pasa
   If far_llave!far_estado = "E" Then GoTo pasa
   
   If far_llave!FAR_CODCIA <> LK_CODCIA And WF <> "A" Then
      MsgBox "!!! NO TE CORRESPONDE..."
      GoTo fin
   End If
   WS_NUMOPER = far_llave!FAR_NUMOPER
   ' AGREGE UNA OPCION MAS ACV
   WS_FECHITA = far_llave!FAR_fecha
   If LK_CODTRA = 1405 Then GoTo SALTAX
   
pasa:
   far_llave.MoveNext
   If far_llave.EOF Then fin = 1
Loop

SALTAX:
fila = 2
far_llave.MoveFirst


Do Until far_llave.EOF
   grid_fac.Rows = fila + 2
   
   If WS_FECHITA = far_llave!FAR_fecha And WS_NUMOPER = far_llave!FAR_NUMOPER Then
      PUB_KEY = far_llave!far_codart
      pu_codcia = LK_CODCIA
      SQ_OPER = 1
      LEER_ART_LLAVE
      If art_LLAVE.EOF And PUB_KEY <> 0 Then
         MsgBox "Error Grave en arti..."
         Exit Sub
      End If
      PUB_CODART = far_llave!far_codart
      pu_codcia = LK_CODCIA
      SQ_OPER = 1
      LEER_ARM_LLAVE
      If arm_llave.EOF Then
         MsgBox "Error Grave en arti..."
         Exit Sub
      End If
      
   If WF <> "A" Then i_codcli.Text = far_llave!far_codclie
   
   i_dias.Text = far_llave!far_DIAS
   i_codven.Text = Nulo_Valor0(far_llave!FAR_codven)
   i_concepto.Text = far_llave!far_concepto
   If Nulo_Valor0(far_llave!far_NUMGUIA) <= 0 Then
      i_numguia.Text = ""
   Else
    i_numguia.Text = far_llave!far_NUMGUIA
   End If
   
   i_numfac_c.Text = far_llave!far_numfac_c
   i_numser_c.Text = far_llave!far_numser_c
   i_fecha_compra.Text = far_llave!far_FECHA_COMPRA
   i_fecha_vcto.Text = Str(DateAdd("d", Val(i_dias.Text), i_fecha_compra.Text))
   If LK_FLAG_GRIFO = "A" Then
    i_turno.Text = far_llave!far_turno
    i_TEXTONCRE.Text = far_llave!FAR_PEDFAC
   End If
   i_serguia.Text = far_llave!far_serguia
   
   If PUB_KEY <> 0 Then grid_fac.TextMatrix(fila, 0) = art_LLAVE!ART_NOMBRE
   If LK_FLAG_ORIGINAL = "A" Then
     grid_fac.TextMatrix(fila, 1) = art_LLAVE!ART_KEY
   Else
     grid_fac.TextMatrix(fila, 1) = art_LLAVE!ART_ALTERNO
   End If
   grid_fac.TextMatrix(fila, 16) = far_llave!far_codart
   grid_fac.TextMatrix(fila, 33) = far_llave!far_codart
   grid_fac.TextMatrix(fila, 2) = far_llave!far_JABAS
   grid_fac.TextMatrix(fila, 3) = far_llave!far_UNIDADES
   grid_fac.TextMatrix(fila, 11) = far_llave!FAR_COSPRO
   If Nulo_Valor0(far_llave!FAR_equiv) > 0 Then
   grid_fac.TextMatrix(fila, 4) = far_llave!FAR_CANTIDAD / Nulo_Valor0(far_llave!FAR_equiv)
   Else
   grid_fac.TextMatrix(fila, 4) = far_llave!FAR_CANTIDAD
   End If
   grid_fac.TextMatrix(fila, 6) = far_llave!FAR_PRECIO
   grid_fac.TextMatrix(fila, 14) = Nulo_Valor0(far_llave!FAR_equiv)
   grid_fac.TextMatrix(fila, 5) = Nulo_Valors(far_llave!FAR_descri)
   grid_fac.TextMatrix(fila, 8) = Nulo_Valor0(far_llave!FAR_FLETE)
   If LK_EMP = "3AA" Or LK_EMP = "HER" Then
    grid_fac.TextMatrix(fila, 10) = Trim(far_llave!FAR_PORDESCTOS)
   Else
     grid_fac.TextMatrix(fila, 10) = far_llave!FAR_DESCTO
   End If
   ' ICA
   grid_fac.TextMatrix(fila, 32) = far_llave!far_precio_neto
   grid_fac.TextMatrix(fila, 35) = far_llave!far_precio_neto
   
   grid_fac.TextMatrix(fila, 11) = far_llave!FAR_COSPRO
   grid_fac.TextMatrix(fila, 12) = far_llave!far_SIGNO_aRM
   If LK_EMP = "3AA" Then
     grid_fac.TextMatrix(fila, 34) = Format(far_llave!far_PESO * far_llave!FAR_CANTIDAD, "0.00")
     grid_fac.TextMatrix(fila, 37) = Format(far_llave!far_PESO, "0.00")
   End If
   If i_fbg.Visible And LK_ACTIVA = "A" Then
    If PUB_TIPMOV <> 0 And (i_fbg.Text = "F" Or i_fbg.Text = "B") Then
        grid_fac.TextMatrix(fila, 13) = arm_llave!ARM_saldo_s
    Else
        grid_fac.TextMatrix(fila, 13) = arm_llave!ARM_Saldo_n
    End If
   Else
      grid_fac.TextMatrix(fila, 13) = arm_llave!ARM_STOCK
   End If
'   If NJ.Visible And LK_ACTIVA = "A" Then
'     If NJ.text = "S" Then
'       grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_saldo_s
'     Else
'       grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_Saldo_n
'     End If
'  End If
'   If WF = "A" Then grid_fac.TextMatrix(fila, 12) = pub_signo_arm
   grid_fac.TextMatrix(fila, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
   grid_fac.TextMatrix(fila, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
   grid_fac.TextMatrix(fila, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
   If PUB_TIPMOV = 10 Then
   ' revisar mejor esto ..pendiente
'      grid_fac.Row = fila ' id_fac.Rows - 1 'fila
'      verifica_precios
   End If
   i_ser_lote.Text = far_llave!FAR_PEDSER
   i_num_lote.Text = far_llave!FAR_PEDFAC
   
   If far_llave!far_estado = "E" Then ww_flag = "S"
   fila = fila + 1
   End If
   WS_MONEDA_CLI = far_llave!far_MONEDA
   i_fecha_compra.Text = far_llave!far_FECHA_COMPRA
   far_llave.MoveNext
Loop
MUE:
calcula_totales
'If i_codcli.Visible Then
'  i_codcli.SetFocus
'  i_codcli_KeyPress 13
'End If
   If WS_MONEDA_CLI = "S" Then
      i_ds.ListIndex = 0
   ElseIf WS_MONEDA_CLI = "D" Then
      If Val(i_ds.ListCount) = 1 Then
        i_ds.ListIndex = 0
      Else
        i_ds.ListIndex = 1
      End If
   Else
      i_ds.ListIndex = -1
   End If
textovar.Visible = False
If Frame4.Visible = False Then Frame4.Visible = True
If i_cambio.Value = 1 And i_cambio.Visible And i_fecha_compra.Visible = True And LK_CODTRA <> 2107 Then
  i_fecha_compra.SetFocus
  Exit Sub
End If
grid_fac.SetFocus


pub_flag_cambio = 0
Exit Sub
' RUTINA PARA JALAR UNA COMPRA EN LA 2403  VENUS


fin:
End Sub
Private Sub i_numguia_GotFocus()
If PUB_TIPMOV = 10 Then
   PUB_SERGUIA = Val(i_serguia.Text)
   llena_numGUIA
End If
PUB_NUMGUIA = PU_NUMFAC
End Sub

Private Sub i_numguia_LostFocus()
If PUB_TIPMOV = 10 And i_numguia.Text <> "" And PUB_NUMGUIA <> Val(i_numguia.Text) Then
   pub_mensaje = "Confirma el numero de Guia de remision ? " & i_numguia.Text
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo)
   If Pub_Respuesta = vbNo Then llena_numGUIA
End If

End Sub

Private Sub i_numplan_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   Exit Sub
End If

NUMERO = i_numplan.WhatsThisHelpID
avanza_grifo

End Sub

Private Sub i_numser_GotFocus()
Azul i_numser, i_numser
End Sub

Private Sub i_numser_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   If i_cambio.Value = 1 Then
      i_cambio.Value = 0
      llena_numfac_grifo
      i_numfac.Locked = True
      i_numser.Locked = True
      i_numfac.BackColor = QBColor(7)
      i_numser.BackColor = QBColor(7)
   Else
      i_cambio.Value = 1
   End If
   If i_cambio.Value = 1 Then
   i_numfac.Locked = False
   i_numser.Locked = False
   i_numfac.BackColor = QBColor(15)
   i_numser.BackColor = QBColor(15)
   Azul i_numfac, i_numfac
   i_fecha_compra.Visible = True
   End If
   
End If

End Sub

Private Sub i_placa_GotFocus()
Azul i_placa, i_placa
End Sub

Private Sub i_placa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
 If i_placa.Visible Then
   i_placa.Visible = False
   i_placas.Visible = True
   i_placas.SetFocus
 End If
End If

End Sub

Private Sub i_placa_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 32 Then
   Exit Sub
End If
If LK_CODTRA = 2101 And PUB_SECUENCIA = 5 Then
 If i_concepto.Visible Then
  i_concepto.SetFocus
  Exit Sub
 End If
End If

If KeyAscii = 32 Then
   If Trim(i_placa.Text) = "" Then
      i_placas.Visible = True
      i_placas.SetFocus
      Exit Sub
   Else
      Exit Sub
   End If
End If
If i_fbg.ListCount > 1 Then
   NUMERO = i_placa.WhatsThisHelpID
   avanza_grifo
Else
   grid_fac.SetFocus
End If
End Sub

Private Sub i_placa_LostFocus()
'If PUB_SECUENCIA = 0 Then


End Sub

Private Sub i_placas_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wvalor
Dim ultimo As Integer
Dim WPLA
If KeyCode = 45 Then
pu_codclie = Val(i_codcli.Text)
If pu_codclie = 0 Then Exit Sub
PUB_TIPREG = 2101
wvalor = InputBox("Ingrese Placa  :", "Datos Necesario...", " ")
If wvalor = "" Then GoTo PA
If Trim(wvalor) = "" Then GoTo PA

PSPLAC_LLAVE(0) = PUB_CODCIA
PSPLAC_LLAVE(1) = PUB_TIPREG
PSPLAC_LLAVE(2) = pu_codclie
PSPLAC_LLAVE(3) = wvalor ' FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 4)
cliplac_llave.Requery
If Not cliplac_llave.EOF Then
    WPLA = Trim(cliplac_llave!TAB_NOMLARGO)
    MsgBox "Descripción Existe. Intente Nuevamente ..", 48, Pub_Titulo
    'If Gridtablas.Visible Then Gridtablas.SetFocus
    For fila = 0 To i_placas.ListCount - 1
      i_placas.ListIndex = fila
      If Trim(WPLA) = Trim(i_placas.Text) Then
         Exit For
      End If
    Next fila
    Exit Sub
End If
PSPLAC_SECU(0) = LK_CODCIA
PSPLAC_SECU(1) = PUB_TIPREG
cliplac_secu.Requery
If cliplac_secu.EOF Then
  ultimo = 0
Else
  ultimo = cliplac_secu!TAB_NUMTAB
End If
ultimo = ultimo + 1
'On Error GoTo SALE
PUB_TIPREG = 2101
cliplac_llave.AddNew
cliplac_llave!TAB_CODCIA = LK_CODCIA
cliplac_llave!TAB_TIPREG = PUB_TIPREG
cliplac_llave!TAB_NUMTAB = ultimo
cliplac_llave!TAB_NOMLARGO = Left(wvalor, 40)
cliplac_llave!tab_nomcorto = Left(wvalor, 10)
cliplac_llave!tab_codclie = pu_codclie
cliplac_llave.Update
'Item = Val(FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 0))

PA:
   
  
'  Exit Sub
'  PUB_TIPREG = 2101
'  PUB_CODCIA = LK_CODCIA
'  Load FrmDatplac
'  FrmDatplac.Caption = "Placas de Clientes : " & PUB_TIPREG
'  FrmDatplac.Show 1
  ultimo = -1
  i_placas.Clear
  PS_PLACAS(0) = Val(i_codcli.Text)
  PS_PLACAS(1) = LK_CODCIA
  TAB_PLACAS.Requery
  Do Until TAB_PLACAS.EOF
    i_placas.AddItem TAB_PLACAS!TAB_NOMLARGO
    If Trim(wvalor) = Trim(TAB_PLACAS!TAB_NOMLARGO) Then
       ultimo = TAB_PLACAS.AbsolutePosition - 1
    End If
    TAB_PLACAS.MoveNext
  Loop
  i_placas.ListIndex = ultimo
  i_placas.SetFocus
  'SendKeys "%{DOWN}"
End If

End Sub

Private Sub i_placas_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 32 Then
   Exit Sub
End If
If LK_CODTRA = 2101 And PUB_SECUENCIA = 5 Then
 If i_concepto.Visible Then
  i_concepto.SetFocus
  Exit Sub
 End If
End If
If KeyAscii = 32 Then
   i_placas.Visible = False
   i_placa.Text = ""
   i_placa.SetFocus
   Exit Sub
End If
''If i_placas.ListIndex = -1 Then
'   i_placas.Visible = False
'   i_placa.SetFocus
'Else
grid_fac.SetFocus
'End If

End Sub

Private Sub i_placas_LostFocus()
i_placa.Text = i_placas.Text
End Sub
Private Sub i_precios_GotFocus()
If i_precios.ListCount <> 0 Then
 i_precios.ListIndex = 0
 SendKeys "%{DOWN}"
End If
'textovar.Visible = False
'i_unidades.Visible = False
'i_mortal.Visible = False
'i_precios.Visible = True
'If Len(grid_fac.text) <> 0 Then
'   i_precios.ListIndex = Right(grid_fac.text, 5)
'Else
'   grid_fac.text = i_precios.text & "     " & i_precios.ListIndex
'   i_precios.ListIndex = 0
'End If

End Sub
Private Sub i_precios_KeyDown(KeyCode As Integer, Shift As Integer)
Dim WTE
If KeyCode <> 45 Then Exit Sub
If LK_PRECIO <> "B" Then Exit Sub
WTE = InputBox("Digite Precio:")
If WTE = "" Then
  Exit Sub
End If
If Not IsNumeric(WTE) Then
  MsgBox "Valor Incorrecto ..., Intente nuevamente ", 48, Pub_Titulo
  Exit Sub
End If
grid_fac.Text = WTE
grid_fac.TextMatrix(grid_fac.Row, 35) = grid_fac.Text

flag_salto = 1
If grid_fac.Row = grid_fac.Rows - 1 Then
   grid_fac.Rows = grid_fac.Rows + 1
End If
grid_fac.SetFocus
grid_fac.Row = grid_fac.Row + 1
grid_fac.RowHeight(grid_fac.Row) = 315
flag_salto = 0

grid_fac.COL = tab_derecha(grid_fac.COL)

End Sub
Private Sub i_precios_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 i_precios.Visible = False
 grid_fac.SetFocus
 Exit Sub
End If
If KeyAscii <> 13 Then Exit Sub
   
grid_fac.Text = Mid(i_precios.Text, 10, 10)
grid_fac.TextMatrix(grid_fac.Row, 35) = grid_fac.Text

grid_fac.TextMatrix(grid_fac.Row, 17) = Right(i_precios.Text, 1)
   
   
flag_salto = 1
If grid_fac.Row = grid_fac.Rows - 1 Then
   grid_fac.Rows = grid_fac.Rows + 1
End If
grid_fac.SetFocus
If tab_derecha(grid_fac.COL) = 1 Then grid_fac.Row = grid_fac.Row + 1
grid_fac.RowHeight(grid_fac.Row) = 315
flag_salto = 0

grid_fac.COL = tab_derecha(grid_fac.COL)

fin:



End Sub

Private Sub i_precios_LostFocus()
'grid_fac.text = i_precios.text & "     " & i_precios.ListIndex
calcula_totales
End Sub


Private Sub i_responsable_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   Exit Sub
End If

NUMERO = i_responsable.WhatsThisHelpID
avanza_grifo

End Sub

Private Sub i_subtotal_Change()
   i_neto.Text = Val(i_subtotal.Text) + Val(i_gastos.Text) - Val(i_descto.Text) + Val(i_impto.Text)
   
End Sub

Private Sub i_TEXTONCRE_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
NUMERO = i_TEXTONCRE.WhatsThisHelpID

avanza_grifo

End Sub


Private Sub i_turno_Change()
i_responsable.ListIndex = -1
End Sub

Private Sub i_turno_GotFocus()
Azul i_turno, i_turno
End Sub

Private Sub i_turno_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If i_turno.Text = "" Then Exit Sub

NUMERO = i_turno.WhatsThisHelpID
avanza_grifo

End Sub

Private Sub i_turno_LostFocus()
If Val(i_turno.Text) > 4 Then
   MsgBox "Turno Invalido..."
   i_turno.SetFocus
   Exit Sub
End If

If Trim(tab_turnos(Val(i_turno.Text))) = "" Then
   MsgBox "Turno Invalido..."
   i_turno.SetFocus
   Exit Sub
End If
If Val(i_turno.Text) > 0 And Val(i_turno.Text) < 5 Then
   i_nomturno.Caption = tab_turnos(Val(i_turno.Text))
   If Not ven_llave.EOF Then
      If ven_llave!vem_codven = Val(i_codven.Text) Then BUSCA_RESP
   End If
End If

End Sub

Private Sub i_unidades_GotFocus()
If LK_EMP = "3AA" Then
   i_unidades_KeyPress 13
   Exit Sub
End If

SendKeys "%{DOWN}"
'i_unidades.ListIndex = 0

'textovar.Visible = False
'i_precios.Visible = False
'i_mortal.Visible = False
'i_unidades.Visible = True
'If Len(grid_fac.text) <> 0 Then
'   i_unidades.ListIndex = Right(grid_fac.text, 5)
'Else
 '  grid_fac.text = i_unidades.text & "     " & i_unidades.ListIndex
 '  i_unidades.ListIndex = 0
'End If

End Sub
Private Sub i_unidades_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 i_unidades.Visible = False
 grid_fac.SetFocus
 Exit Sub
End If
If KeyAscii <> 13 Then Exit Sub
If i_unidades.Text = "" Then Exit Sub
If Trim(Mid(i_unidades.Text, 23, 3)) = "" Then
   MsgBox "Reiniciar Equipo...- Unidades"
   End
End If

PUB_SECUEN = Val(Mid(i_unidades.Text, 23, 3))

verifica_precios
grid_fac.Text = i_unidades.Text
grid_fac.TextMatrix(grid_fac.Row, 14) = Right(i_unidades.Text, 6)
grid_fac.TextMatrix(grid_fac.Row, 19) = i_unidades.ListIndex

If PUB_TIPMOV = 20 Then
   grid_fac.TextMatrix(grid_fac.Row, 6) = Nulo_Valor0(arm_llave!ARM_COSTO_ULT)
End If
If LK_CODTRA = 2403 And PUB_SECUENCIA = 10 Then
   grid_fac.TextMatrix(grid_fac.Row, 6) = Nulo_Valor0(arm_llave!arm_cospro)
End If


If Not arm_llave.EOF And pub_signo_arm = -1 Then
   PUB_CANTIDAD = Val(grid_fac.TextMatrix(grid_fac.Row, 4)) * Val(grid_fac.TextMatrix(grid_fac.Row, 14))
   If PUB_CANTIDAD > Val(grid_fac.TextMatrix(grid_fac.Row, 13)) And Nulo_Valors(art_LLAVE!art_flag_stock) <> "A" Then
      MsgBox "Ojo Cantidad Mayor que Stock ..." & "Existencia : " & Val(grid_fac.TextMatrix(grid_fac.Row, 13))
      If LK_USU_STOCK <> "A" Then
         grid_fac.TextMatrix(grid_fac.Row, 4) = ""
         grid_fac.TextMatrix(grid_fac.Row, 5) = ""
         grid_fac.TextMatrix(grid_fac.Row, 6) = ""
         grid_fac.TextMatrix(grid_fac.Row, 7) = ""
         Exit Sub
      End If

   End If
End If


If (LK_CODTRA = 2403 Or LK_CODTRA = 2425 Or tab_derecha(grid_fac.COL) = 1) Then
   flag_salto = 1
   If grid_fac.Row = grid_fac.Rows - 1 Then grid_fac.Rows = grid_fac.Rows + 1
      If LK_CODTRA = 2414 Then
      i_unidades.Visible = False
      End If
      grid_fac.SetFocus
      grid_fac.Row = grid_fac.Row + 1
      grid_fac.RowHeight(grid_fac.Row) = 315
End If
   
flag_salto = 0



If i_precios.ListCount = 1 And PUB_TIPMOV = 10 Then
   flag_salto = 1
   i_precios.ListIndex = 0
   grid_fac.TextMatrix(grid_fac.Row, 6) = Mid(i_precios.Text, 10, 10)
   grid_fac.TextMatrix(grid_fac.Row, 35) = Mid(i_precios.Text, 10, 10)
   
   grid_fac.TextMatrix(grid_fac.Row, 17) = Right(i_precios.Text, 1)

   If Trim(SUT_LLAVE!sut_descto) <> "" And grid_fac.COL = 5 And i_precios.ListCount = 1 Then
      grid_fac.COL = 10
   Else
      If grid_fac.Row = grid_fac.Rows - 1 Then grid_fac.Rows = grid_fac.Rows + 1
      grid_fac.Row = grid_fac.Row + 1
      grid_fac.RowHeight(grid_fac.Row) = 315
      flag_salto = 0
      grid_fac.COL = tab_derecha(6)
      If textovar.Visible = True Then
         textovar.SetFocus
      Else
         grid_fac.SetFocus
      End If
  End If
Else
   grid_fac.COL = tab_derecha(grid_fac.COL)
End If
End Sub

Private Sub i_unidades_LostFocus()
calcula_totales
End Sub


Private Sub LisTransa_DblClick()
LisTransa_KeyPress 13
End Sub


Private Sub i_codcli_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem    ' Variable FoundItem.
If KeyAscii = 27 Then
 i_codcli.Text = ""
 Exit Sub
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo OJO
pu_codclie = Val(i_codcli.Text)
On Error GoTo 0
If Len(i_codcli.Text) = 0 Then
   Exit Sub
End If
If pu_codclie <> 0 And IsNumeric(i_codcli.Text) = True Then
   If Len(Trim(i_codcli.Text)) = LK_DIG_RUC Then ' LONG DEL RUC
        'pu_cp = "" ' Left(CmbCGP.Text, 1)
        PUB_RUC = Trim(i_codcli.Text)
        SQ_OPER = 4
        pu_codcia = LK_CODCIA
        LEER_CLI_LLAVE
        If cli_ruc.EOF Then
           MsgBox "R.U.C. No Existe ", 48, Pub_Titulo
           Exit Sub
        End If
        i_codcli.Text = cli_ruc!cli_codclie
   End If
   On Error GoTo OJO
   SQ_OPER = 1
   pu_codclie = Val(i_codcli.Text)
   pu_codcia = LK_CODCIA
   pu_cp = PUB_CP
   LEER_CLI_LLAVE
   On Error GoTo 0
   If cli_llave.EOF Then
    Azul i_codcli, i_codcli
    MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
    i_codcli.SetFocus
    GoTo fin
   Else
      i_nomCLI.Caption = Trim(cli_llave!cli_nombre) & " - " & Trim(cli_llave!cli_ruc_esposo)
      NUMERO = i_codcli.WhatsThisHelpID
      avanza_grifo
   End If
Else
On Error GoTo sigue
   If loc_key <> 0 Then valor = UCase(LV_CLI.ListItems.Item(loc_key).Text)
   If Trim(UCase(i_codcli.Text)) = Left(valor, Len(Trim(i_codcli.Text))) Then
   Else
      Exit Sub
   End If
   If loc_key = 0 Then Exit Sub
   i_codcli.Text = Trim(LV_CLI.ListItems.Item(loc_key).SubItems(1))
   pu_codclie = Val(i_codcli.Text)
   pu_cp = PUB_CP
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_CLI_LLAVE
   i_nomCLI.Caption = Trim(cli_llave!cli_nombre) & " - " & Trim(cli_llave!cli_ruc_esposo)
   NUMERO = i_codcli.WhatsThisHelpID
   avanza_grifo
End If
LV_CLI.Visible = False
If LK_FLAG_EXED = "A" Then
If Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) <> 1 And PUB_TIPMOV = 10 And pub_signo_car <> 0 And Nulo_Valor0(cli_llave!cli_limcre) = 0 Then
   MsgBox "Cliente No tiene Limite de Credito.", 48, Pub_Titulo
End If
End If
If Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) <> 1 And PUB_TIPMOV = 10 And pub_signo_car <> 0 And Nulo_Valors(cli_llave!CLI_estado) <> "A" Then
   MsgBox "              !!! O J O !!! " + Chr(13) + "Cliente No está ACTIVO en el Sistema.", 48, Pub_Titulo
End If
Exit Sub
sigue:
If Err.Number = 35600 Then
  Exit Sub
End If
fin:
OJO:
End Sub



Private Sub i_codven_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem    ' Variable FoundItem.

If KeyAscii = 27 Then
  i_codven.Text = ""
  LV_VEN.Visible = False
  Exit Sub
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo OJO
PUB_CODVEN = Val(i_codven.Text)
On Error GoTo 0
If Len(i_codven.Text) = 0 Then
   Exit Sub
End If
If PUB_CODVEN <> 0 And IsNumeric(i_codven.Text) = True Then
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_VEN_LLAVE
   If ven_llave.EOF Then
    Azul i_codven, i_codven
    MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
    i_codven.SetFocus
    GoTo fin
   Else
'      i_vendedor.Text = Nulo_Valor0(ven_llave!VEM_VENDEDOR)
      i_nomven.Caption = ven_llave(2)
      NUMERO = i_codven.WhatsThisHelpID
      avanza_grifo

   End If
Else
On Error GoTo sigue
   If loc_key > 0 Then valor = UCase(LV_VEN.ListItems.Item(loc_key).Text)
   If Trim(UCase(i_codven.Text)) = Left(valor, Len(Trim(i_codven.Text))) Then
   Else
      Exit Sub
   End If
   
   If loc_key = 0 Then Exit Sub
   
   i_codven.Text = Trim(LV_VEN.ListItems.Item(loc_key).SubItems(1))
   PUB_CODVEN = Val(i_codven.Text)
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_VEN_LLAVE
   i_nomven.Caption = ven_llave(2)
   NUMERO = i_codven.WhatsThisHelpID
'   i_vendedor.Text = ven_llave!VEM_VENDEDOR
   
   avanza_grifo
End If
LV_VEN.Visible = False

fin:
sigue:
OJO:
End Sub



Private Sub i_concepto_GotFocus()
Azul i_concepto, i_concepto



End Sub

Private Sub i_concepto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = i_concepto.WhatsThisHelpID

avanza_grifo
fin:

End Sub
Private Sub i_def_GotFocus()
Dim pos1 As Integer
Dim CARAC As String
If TRANS.Text = "" Then
   GoTo fin
End If
'SendKeys "%{DOWN}"
fin:
End Sub

Private Sub i_DEF_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
NUMERO = TABLA_TAG(tra_llave(2))
WS_INDICE_RETORNO = NUMERO
Controls(NUMERO).SetFocus

fin:
End Sub
Private Sub i_def_LostFocus()
Dim pos1, CARAC, i, izq, der
If i_def.ListIndex = -1 Then Exit Sub
'Aqui se lee la contabilidad
pos1 = InStr(1, FORM_GRIFO.i_def.List(i_def.ListIndex), ".", 1)
pos1 = pos1 - 1
CARAC = Mid(i_def.List(i_def.ListIndex), 1, pos1)
PUB_SECUENCIA = Val(CARAC)
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_SUT_LLAVE
If SUT_LLAVE.EOF Then
   MsgBox "No existe Definicion en Sub_Transacciones... "
   GoTo fin
End If
Check1.Value = 0
i_fbg.Clear
If LK_CODTRA = 2103 Then
   i_fbg.AddItem "O"
   Frame4.Visible = True
ElseIf LK_CODTRA = 1402 Then
   gridC.Visible = False
   i_fbg.AddItem "G"
   Frame4.Visible = True
Else
If PUB_SECUENCIA = 1 Or PUB_SECUENCIA = 0 Or PUB_SECUENCIA = 5 Then
   i_fbg.AddItem "F"
   i_fbg.AddItem "B"
ElseIf PUB_SECUENCIA = 15 Then
   i_fbg.AddItem "G"
ElseIf PUB_SECUENCIA = 20 Then
   i_fbg.AddItem "C"
ElseIf PUB_SECUENCIA = 25 Then
   i_fbg.AddItem "P"
End If
End If
If LK_CODTRA = 2212 Then i_fbg.AddItem "G"
If LK_CODTRA = 2103 Then
 GridT.TextMatrix(0, 0) = "Producto"
 GridT.TextMatrix(0, 1) = "Codigo"
 GridT.TextMatrix(0, 2) = "Cantidad"
 GridT.TextMatrix(0, 3) = "Peso"
 If PUB_TIPMOV = 177 Then
   GridT.ColWidth(6) = 1
   GridT.ColWidth(7) = 1
 Else
   GridT.ColWidth(6) = 500
   GridT.ColWidth(7) = 800
 End If
 If PUB_SECUENCIA = 1 Then
   GridT.TextMatrix(0, 6) = "Card."
   GridT.TextMatrix(0, 7) = "NroO/P"
 End If
 
 Option5.Visible = True
 Option6.Visible = True
 If PUB_SECUENCIA = 1 Or PUB_SECUENCIA = 0 Then
  Option5.Visible = False
  Option6.Visible = False
 End If
 
 GridT.RowHeight(GridT.Rows - 1) = 285
End If
If i_fbg.ListCount > 0 Then i_fbg.ListIndex = 0  ' OJO CORREGIR EN EL DE ALAN 05/12/00

grid_fac.Clear
pasa_cabeza
pasa_def
   
If Frame4.Visible = True Or LK_CODTRA = 2211 Or LK_CODTRA = 2212 Or LK_CODTRA = 1402 Then
   Frame2.Enabled = False
   grid_fac.ColWidth(2) = 0
   grid_fac.ColWidth(3) = 0
   grid_fac.ColWidth(5) = 0
   grid_fac.ColWidth(6) = 0
   grid_fac.ColWidth(7) = 0
   grid_fac.ColWidth(4) = 0
   grid_fac.ColWidth(8) = 0
   grid_fac.ColWidth(9) = 0
   grid_fac.ColWidth(10) = 0
   grid_fac.ColWidth(11) = 0
   grid_fac.ColWidth(12) = 0
   
   If PUB_TIPMOV = 20 Then Frame2.Enabled = True
   
   If Len(Trim(SUT_LLAVE!SUT_JABAS)) <> 0 Then
      grid_fac.TextMatrix(0, 2) = SUT_LLAVE!SUT_JABAS
      grid_fac.ColWidth(2) = 600
   End If
   If Len(Trim(SUT_LLAVE!SUT_POLLOS)) <> 0 Then
      grid_fac.TextMatrix(0, 3) = SUT_LLAVE!SUT_POLLOS
      grid_fac.ColWidth(3) = 600
   End If
   
   
   If Nulo_Valors(SUT_LLAVE!sut_is) = "A" And LK_CODTRA <> 2414 And LK_CODTRA <> 2210 Then
      grid_fac.ColWidth(12) = 1200
   End If
   
   If Nulo_Valors(SUT_LLAVE!SUT_UNIDADES) = "A" Then
      grid_fac.ColWidth(5) = 1000
      grid_fac.ColWidth(6) = 1000
   End If
   
   If Nulo_Valors(SUT_LLAVE!SUT_PRECIO) = "1" Then
      grid_fac.ColWidth(6) = 0
   End If

   grid_fac.ColWidth(7) = 900

   
   If PUB_TIPMOV = 20 Then
      grid_fac.ColWidth(6) = 1000
      grid_fac.ColWidth(8) = 600
      grid_fac.ColWidth(10) = 600
   End If
   
   If LK_CODTRA = 2211 Then
      grid_fac.ColWidth(0) = 4000
      grid_fac.ColWidth(1) = 900
      grid_fac.ColWidth(2) = 1000
      grid_fac.ColWidth(3) = 0
      grid_fac.ColWidth(4) = 1200
      grid_fac.ColWidth(5) = 1000
      grid_fac.ColWidth(6) = 1000
      grid_fac.ColWidth(7) = 0
      grid_fac.ColWidth(10) = 0
      grid_fac.ColWidth(13) = 0
      grid_fac.ColWidth(15) = 0
      grid_fac.ColWidth(38) = 0
   End If
   If LK_CODTRA = 2103 Then
      grid_fac.ColWidth(0) = 3500
      grid_fac.ColWidth(2) = 0
      grid_fac.ColWidth(3) = 0
      grid_fac.ColWidth(1) = 600
      grid_fac.ColWidth(5) = 0
      grid_fac.ColWidth(6) = 0
      grid_fac.ColWidth(7) = 0
      grid_fac.ColWidth(13) = 1100
      grid_fac.ColWidth(38) = 800
      grid_fac.ColWidth(10) = 0
      grid_fac.ColWidth(15) = 0
      
   End If
   If LK_CODTRA = 2212 Then
      grid_fac.ColWidth(0) = 3500
      grid_fac.ColWidth(1) = 600
      grid_fac.ColWidth(3) = 0
      grid_fac.ColWidth(2) = 600
      grid_fac.ColWidth(5) = 500
      grid_fac.ColWidth(6) = 500
      grid_fac.ColWidth(7) = 900
      grid_fac.ColWidth(13) = 0
      grid_fac.ColWidth(38) = 0
      grid_fac.ColWidth(10) = 0
      grid_fac.ColWidth(15) = 0
   End If
   
   
   i = 1
   GoSub destinos
   If grid_fac.ColWidth(2) <> 0 Then
      i = 2
      GoSub destinos
   End If
   
   If grid_fac.ColWidth(3) <> 0 Then
      i = 3
      GoSub destinos
   End If
   
   If grid_fac.ColWidth(5) <> 0 Then
      i = 5
      GoSub destinos
   End If
   
   If grid_fac.ColWidth(6) <> 0 Then
      i = 6
      GoSub destinos
   End If
   If grid_fac.ColWidth(7) <> 0 Then
      i = 7
      GoSub destinos
   End If
   
   If grid_fac.ColWidth(8) <> 0 Then
      i = 8
      GoSub destinos
   End If
   If grid_fac.ColWidth(9) <> 0 Then
      i = 9
      GoSub destinos
   End If
   If grid_fac.ColWidth(10) <> 0 Then
      i = 10
      GoSub destinos
   End If
   If grid_fac.ColWidth(12) <> 0 Then
      i = 12
      GoSub destinos
   End If

   
   If LK_CODTRA <> 2103 And LK_CODTRA <> 2211 Then grid_fac.ColWidth(6) = 1000
   
   'If Nulo_Valors(SUT_LLAVE!SUT_mortal) = "1" Then
   '   grid_fac.ColWidth(9) = 2000
   'End If

   If LK_CODTRA = 2414 Then
      grid_fac.ColWidth(12) = 1200
   End If
   If Trim(SUT_LLAVE!sut_descto) <> "" Then
      grid_fac.ColWidth(10) = 800
   End If
   grid_fac.ColWidth(4) = 900
   tab_derecha(15) = 1
   tab_izquierda(15) = 10
   tab_derecha(10) = 15
   tab_derecha(7) = 1
   
   If LK_CODTRA = 2103 Or LK_CODTRA = 2211 Or LK_CODTRA = 2212 Then
      grid_fac.ColWidth(10) = 0
      grid_fac.ColWidth(15) = 0
      tab_derecha(1) = 4
      tab_derecha(2) = 4
      tab_derecha(3) = 4
      tab_derecha(4) = 1
      tab_derecha(38) = 4
      tab_izquierda(38) = 4
      tab_izquierda(4) = 1
      Frame4.Visible = False
   End If
   If LK_CODTRA = 1402 Then
      grid_fac.ColWidth(10) = 0
      grid_fac.ColWidth(15) = 0
      tab_derecha(1) = 4
      tab_derecha(4) = 5
      tab_derecha(5) = 6
   End If
   
   If SUT_LLAVE!SUT_SECUENCIA = 2 And LK_CODTRA <> 2211 Then i_fbg.ListIndex = 2
   
   If LK_CODTRA = 2211 Then
      tab_derecha(1) = 2
      tab_derecha(2) = 4
  End If

End If


'BolFac_Click
   If PUB_TIPMOV = 93 Then
      grid_fac.ColWidth(10) = 400
   End If
   If PUB_TIPMOV <> 0 And LK_CODTRA = 2103 Then
   PSCNT_LLAVE.rdoParameters(0) = LK_CODCIA
   PSCNT_LLAVE.rdoParameters(1) = LK_CODTRA
   PSCNT_LLAVE.rdoParameters(2) = SUT_LLAVE!SUT_SECUENCIA
   cnt_llave.Requery
   If cnt_llave.EOF Then
      'MsgBox "No ha definido Serie y Correlativo ..."
      GoTo SALT
   End If
   i_numser.Text = Nulo_Valor0(cnt_llave!cnt_serie)
   i_ser_lote.Text = Nulo_Valor0(cnt_llave!cnt_serie)
'   If Nulo_Valor0(cnt_llave!cnt_serie) > 0 Then WS_ESTADO = 1
   End If
SALT:
   If PUB_TIPMOV <> 0 And LK_CODTRA <> 2412 And LK_CODTRA <> 2410 And LK_CODTRA <> 2580 And LK_CODTRA <> 1122 And LK_CODTRA <> 2748 Then
        i_moneda.Visible = True
   Else
        i_moneda.Visible = False
   End If
   
Exit Sub

destinos:
   izq = i - 1
   der = i + 1
   tab_derecha(i) = 1
   tab_izquierda(i) = 1
   Do Until der > 12 And izq < 1
      If der < 13 Then
      If grid_fac.ColWidth(der) <> 0 Then
         tab_derecha(i) = der
         der = 13
      End If
      End If
      
      If izq > 0 Then
         If grid_fac.ColWidth(izq) <> 0 Then
            tab_izquierda(i) = izq
            izq = 0
         End If
      End If
       
      der = der + 1
      izq = izq - 1
   Loop

  Return
fin:


End Sub

Private Sub i_descto_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_descto, KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If
PUB_DESCTO = Val(i_descto.Text)
grid_fac.Row = 2
grid_fac.COL = 1
grid_fac.SetFocus


fin:

End Sub
Private Sub i_dias_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 i_dias.Text = ""
 Exit Sub
End If
If KeyAscii = 13 Then GoTo OTRO
SOLO_ENTERO KeyAscii

Dim FECHA_DIA As Date
Dim fecha_vcto As Variant
Dim pub_mensaje As String
If KeyAscii <> 13 Then
   GoTo fin
End If
'If ENTERO(Val(i_dias.text)) = False Then
'   Azul i_dias, i_dias
'   GoTo fin
'End If

If LK_CODTRA = 2101 Then
If Val(i_dias.Text) > 30 Then
   pub_mensaje = " ¿Esta seguro del numero de Dias... ?"
     Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then   ' El usuario eligió
     i_dias.SetFocus
     Exit Sub
  End If
End If
End If
OTRO:
NUMERO = i_dias.WhatsThisHelpID
avanza_grifo
fin:

End Sub
Private Sub i_dias_LostFocus()
wtemporal = ""
i_fecha_vcto.Text = Str(DateAdd("d", Val(i_dias.Text), LK_FECHA_DIA))
End Sub

Private Sub i_fbg_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
NUMERO = i_fbg.WhatsThisHelpID
If i_cambio.Value = 0 Then
   avanza_grifo
Else
   i_numser.Text = ""
   i_numser.SetFocus
End If

fin:

End Sub

Private Sub i_fecha_vcto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
If IsDate(i_fecha_vcto) = False Then
   MsgBox "FECHA NO VALIDA ...", 48, Pub_Titulo
   GoTo fin
End If

NUMERO = i_fecha_vcto.WhatsThisHelpID
avanza_grifo
fin:

End Sub
Private Sub i_gastos_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_gastos, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo SALE
PUB_GASTOS = Val(i_gastos.Text)
grid_fac.SetFocus
On Error GoTo 0


fin:
Exit Sub
SALE:
Azul i_gastos, i_gastos
End Sub


Private Sub i_importe_amort_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_importe_amort, KeyAscii

If KeyAscii <> 13 Then Exit Sub
If LK_CODTRA <> 2770 Then
   NUMERO = i_importe_amort.WhatsThisHelpID
   avanza_grifo
End If




fin:

End Sub

Private Sub i_importe_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_importe, KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = i_importe.WhatsThisHelpID
avanza_grifo
 
 
 
fin:

End Sub


Private Sub i_gastos_fijos_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_gastos_fijos, KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = i_gastos_fijos.WhatsThisHelpID
avanza_grifo
fin:

End Sub




Private Sub i_limcre_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_limcre, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If

If i_limcre.Text = "" Then Exit Sub
NUMERO = i_limcre.WhatsThisHelpID
avanza_grifo
fin:
'End Sub



'salta:

'i_subtotal.Enabled = False
'i_gastos.Enabled = False
'i_descto.Enabled = False
'i_impto.Enabled = False
'i_neto.Enabled = False
'i_dias.Enabled = False
'i_fecha_vcto.Enabled = False
'grid_fac.SetFocus
'grid_fac.Row = 2
'grid_fac.Col = 1

'fin:

End Sub


Private Sub i_num_oper_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = i_num_oper.WhatsThisHelpID
avanza_grifo
fin:

End Sub

Private Sub i_numdoc_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = i_numdoc.WhatsThisHelpID
avanza_grifo
fin:

End Sub

Private Sub I_NUMDOC_R_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If

'NUMERO = i_numdoc_r.WhatsThisHelpID
'avanza_grifo


fin:

End Sub
Private Sub i_numfac_c_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
If LK_CODTRA = 2408 Then GoTo SALTA
If LK_CODTRA = 2406 Then GoTo SALTA
If LK_CODTRA = 2748 Then GoTo SALTA
If LK_CODTRA = 2410 Then GoTo SALTA
If LK_CODTRA = 1122 Then If gridl.Visible = False Then GoTo SALTA

If PUB_TIPMOV = 20 Then GoTo SALTA
   
   SQ_OPER = 4
   pu_codcia = LK_CODCIA
   PU_FBG = i_fbg.Text
   PUB_TIPDOC = SUT_LLAVE!SUT_tipdoc
   PUB_NUMSER = Val(i_numser_c.Text)
   PUB_NUMFAC = Val(i_numfac_c.Text)
   
   LEER_CAR_LLAVE
   Do Until car_far.EOF
      If car_far!CAR_IMPORTE <> 0 And car_far!CAR_cp = "C" Then Exit Do
      car_far.MoveNext
   Loop
   If car_far.EOF Then
      MsgBox "Documento no Existe.", 48, Pub_Titulo
      i_numser_c.SetFocus
      Exit Sub
   End If
   
   If Val(i_codven.Text) = 0 And LK_CODTRA = 2770 And LK_EMP = "HER" Then
      MsgBox "Falta Cod. Vendedor / Cobrador ", 48, Pub_Titulo
      If i_codven.Visible Then i_codven.SetFocus
      Exit Sub
   End If
   
   If gridl.Visible = False Then gridl.Visible = True
   If textovarl.Visible = False Then textovarl.Visible = True
   If textovar_canje.Visible = False Then textovar_canje.Visible = True
   textovar_canje.Width = 30
   textovar_canje.Height = 30
   
   pasa_cabeza_canje
   pasa_cabeza_L

   pu_codclie = car_far!CAR_CODCLIE
   pu_codcia = car_far!CAR_CODCIA
   pu_cp = car_far!CAR_cp
   SQ_OPER = 1
   LEER_CLI_LLAVE
      
      FILAX = gridl.Rows - 1
      
      If gridl.TextMatrix(FILAX, 0) = "" Then
      Else
         gridl.Rows = gridl.Rows + 1
         FILAX = gridl.Rows - 1
      End If
      
      gridl.RowHeight(FILAX) = 350
      gridl.TextMatrix(FILAX, 15) = cli_llave!cli_codclie
      gridl.TextMatrix(FILAX, 0) = cli_llave!cli_nombre
      gridl.TextMatrix(FILAX, 13) = car_far!CAR_TIPDOC
      gridl.TextMatrix(FILAX, 1) = Nulo_Valors(car_far!CAR_FBG)
      If car_far!CAR_cp = "P" Then
          gridl.TextMatrix(FILAX, 2) = Nulo_Valor0(car_far!CAR_NUMSER_C)
          gridl.TextMatrix(FILAX, 3) = Nulo_Valor0(car_far!CAR_NUMFAC_C)
       Else
          gridl.TextMatrix(FILAX, 2) = Nulo_Valor0(car_far!CAR_numSER)
          gridl.TextMatrix(FILAX, 3) = Nulo_Valor0(car_far!CAR_numfac)
       End If
       
       If car_far!CAR_MONEDA = "S" Then
          gridl.TextMatrix(FILAX, 4) = "S/." & car_far!CAR_IMPORTE
       Else
          gridl.TextMatrix(FILAX, 4) = "  $" & car_far!CAR_IMPORTE
       End If

       
       gridl.TextMatrix(FILAX, 5) = ""
       gridl.TextMatrix(FILAX, 6) = Format(car_far!car_fecha_vcto, "dd/mm/yy")
       gridl.TextMatrix(FILAX, 7) = Val(i_codven.Text)
       gridl.TextMatrix(FILAX, 9) = car_far!CAR_IMP_INI
       gridl.TextMatrix(FILAX, 8) = car_far!CAR_FECHA_INGR
       gridl.TextMatrix(FILAX, 10) = car_far!CAR_cp
       gridl.TextMatrix(FILAX, 11) = car_far!car_SERDOC
       gridl.TextMatrix(FILAX, 12) = car_far!car_NUMDOC
       gridl.TextMatrix(FILAX, 14) = car_far!CAR_CODCIA
       gridl.TextMatrix(FILAX, 18) = car_far!CAR_codven
       gridl.TextMatrix(FILAX, 24) = car_far!CAR_MONEDA
       gridl.TextMatrix(FILAX, 25) = car_far!CAR_IMPORTE
       gridl.TextMatrix(FILAX, 26) = car_far!car_fecha_vcto
       
       If gridl.Visible = False Then gridl.Visible = True
       gridl.SetFocus
       flag_salto = 1
       gridl.Row = FILAX
       flag_salto = 0
       gridl.COL = 5


fin:
Exit Sub
SALTA:
   NUMERO = i_numfac_c.WhatsThisHelpID
   avanza_grifo

End Sub

Private Sub i_numfac_KeyPress(KeyAscii As Integer)

SOLO_ENTERO KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = i_numfac.WhatsThisHelpID
If i_fecha_compra.Visible Then Azul i_fecha_compra, i_fecha_compra

avanza_grifo
fin:

End Sub



Private Sub i_numguia_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
NUMERO = i_numguia.WhatsThisHelpID
avanza_grifo
fin:

End Sub

Private Sub i_numser_c_GotFocus()
'   Azul i_numser_c, i_numser_c
End Sub

Private Sub i_numser_c_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
NUMERO = i_numser_c.WhatsThisHelpID
avanza_grifo
fin:

End Sub

Private Sub i_numser_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = i_numser.WhatsThisHelpID

avanza_grifo
fin:

End Sub
Private Sub i_numser_r_KeyPress(KeyAscii As Integer)
Dim respuesta
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If

   
NUMERO = i_numser_r.WhatsThisHelpID
avanza_grifo
   

fin:

End Sub




Private Sub i_subtotal_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_subtotal, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If

fin:

End Sub
Private Sub i_subtotal_LostFocus()
PUB_SUBTOTAL_BAK = Val(i_subtotal.Text)
End Sub
Private Sub i_tipdoc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = i_tipdoc.WhatsThisHelpID

avanza_grifo
fin:

End Sub

Private Sub LisTransa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  TRANS.Text = Trim(Left(LisTransa.Text, 7))
  LisTransa.Visible = False
  TRANS.SetFocus
ElseIf KeyAscii = 27 Then
  KeyAscii = 0
  LisTransa.Visible = False
  TRANS.SetFocus
End If

End Sub

Private Sub LisTransa_LostFocus()
  LisTransa.Visible = False
  TRANS.SetFocus
End Sub

Private Sub Lmoneda_Click()
'If LK_CODTRA = 2412 And Val(grid_fac.TextMatrix(1, 7)) <> 0 Then
'if Val(grid_fac.TextMatrix(1, 7)) <> 0 Then
'Else
   If Trim(lmoneda.Caption) = "$" Then
     lmoneda.Caption = "S/."
     WS_MONEDA_CLI = "S"
   Else
     lmoneda.Caption = "  $"
     WS_MONEDA_CLI = "D"
   End If
'End If

End Sub

Private Sub LV_ART_DblClick()

If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
 loc_key = LV_ART.SelectedItem.Index
 textovar.Text = Trim(LV_ART.ListItems.Item(loc_key).Text) & " "
 textovar_KeyPress 13
Else
 loc_key = LV_ART.SelectedItem.Index
 textovar.Text = Trim(LV_ART.ListItems.Item(loc_key).Text) & " "
 textovar_KeyPress 13
End If

End Sub

Private Sub LV_ART_GotFocus()
If loc_key <> 0 Then
 Set LV_ART.SelectedItem = LV_ART.ListItems(loc_key)
 LV_ART.ListItems.Item(loc_key).Selected = True
 LV_ART.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub LV_ART_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = LV_ART.SelectedItem.Index
 If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  textovar.Text = Trim(LV_ART.ListItems.Item(loc_key).Text) & " "
 Else
  textovar.Text = Trim(LV_ART.ListItems.Item(loc_key).Text) & " "
 End If
End If

End Sub

Private Sub LV_ART_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 If textovar.Visible Then
   textovar.Text = ""
   textovar.SetFocus
 End If
End If

If KeyAscii = 13 Then
 If textovar.Visible Then
   textovar_KeyPress 13
 End If

End If
End Sub

Private Sub LV_ART_LostFocus()
LV_ART.Visible = False
End Sub
Private Sub LV_ART2_DblClick()

If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
 loc_key = LV_ART2.SelectedItem.Index
 textovar2.Text = Trim(LV_ART2.ListItems.Item(loc_key).Text) & " "
 textovar2_KeyPress 13
Else
 loc_key = LV_ART2.SelectedItem.Index
 textovar2.Text = Trim(LV_ART2.ListItems.Item(loc_key).Text) & " "
 textovar2_KeyPress 13
End If

End Sub

Private Sub LV_ART2_GotFocus()
If loc_key <> 0 Then
 Set LV_ART2.SelectedItem = LV_ART2.ListItems(loc_key)
 LV_ART2.ListItems.Item(loc_key).Selected = True
 LV_ART2.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub LV_ART2_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = LV_ART2.SelectedItem.Index
 If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  textovar2.Text = Trim(LV_ART2.ListItems.Item(loc_key).Text) & " "
 Else
  textovar2.Text = Trim(LV_ART2.ListItems.Item(loc_key).Text) & " "
 End If
End If

End Sub

Private Sub LV_ART2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 If textovar2.Visible Then
   textovar2.Text = ""
   textovar2.SetFocus
 End If
End If

If KeyAscii = 13 Then
 If textovar2.Visible Then
   textovar2_KeyPress 13
 End If

End If

End Sub

Private Sub LV_ART2_LostFocus()
LV_ART2.Visible = False

End Sub

Private Sub LV_CCM_GotFocus()
If loc_key <> 0 Then
 Set LV_CCM.SelectedItem = LV_CCM.ListItems(loc_key)
 LV_CCM.ListItems.Item(loc_key).Selected = True
 LV_CCM.ListItems.Item(loc_key).EnsureVisible
End If

End Sub


Private Sub LV_CLI_DblClick()
loc_key = LV_CLI.SelectedItem.Index
i_codcli.Text = Trim(LV_CLI.ListItems.Item(loc_key).Text) & " "
i_codcli_KeyPress 13

End Sub

Private Sub LV_CLI_GotFocus()
If loc_key <> 0 Then
 Set LV_CLI.SelectedItem = LV_CLI.ListItems(loc_key)
 LV_CLI.ListItems.Item(loc_key).Selected = True
 LV_CLI.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub LV_CLI_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
  loc_key = LV_CLI.SelectedItem.Index
  i_codcli.Text = Trim(LV_CLI.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub LV_CLI_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 If i_codcli.Visible Then
    LV_CLI.Visible = False
    i_codcli.Text = ""
    i_codcli.SetFocus
 End If
End If
If KeyAscii = 13 Then
 If i_codcli.Visible Then
    i_codcli_KeyPress 13
 End If
End If
End Sub

Private Sub LV_CLI_LostFocus()
LV_CLI.Visible = False
End Sub

Private Sub LV_VEN_DblClick()
loc_key = LV_VEN.SelectedItem.Index
i_codven.Text = Trim(LV_VEN.ListItems.Item(loc_key).Text) & " "
i_codven_KeyPress 13
End Sub

Private Sub LV_VEN_GotFocus()
If loc_key <> 0 Then
 Set LV_VEN.SelectedItem = LV_VEN.ListItems(loc_key)
 LV_VEN.ListItems.Item(loc_key).Selected = True
 LV_VEN.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub LV_VEN_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
  loc_key = LV_VEN.SelectedItem.Index
  i_codven.Text = Trim(LV_VEN.ListItems.Item(loc_key).Text) & " "
End If
End Sub

Private Sub LV_VEN_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 If i_codven.Visible Then
    i_codven.Text = ""
    i_codven.SetFocus
 End If
End If
If KeyAscii = 13 Then
 If i_codven.Visible Then
    i_codven_KeyPress 13
 End If
End If

End Sub

Private Sub LV_VEN_LostFocus()
LV_VEN.Visible = False
End Sub


Private Sub Option1_Click()
PROCESA_GRIDL
End Sub

Private Sub Option2_Click()
PROCESA_GRIDL
End Sub

Private Sub Option3_Click()
PROCESA_GRIDL
End Sub

Private Sub Option4_Click()
Frame4.Visible = True
End Sub

Private Sub Option5_Click()
textovar2.Visible = False
textovar3.Visible = False
textovar.Visible = False
Frame5.Visible = True
Frame4.Visible = False
Frame6.Visible = False
GridD.Visible = True
GridT.Visible = True
Dim f As Integer
Dim FF As Integer
Dim ww_soles As Currency
If Val(GridT.TextMatrix(1, 3)) = Val(GridD.TextMatrix(1, 1)) Then GoTo PAS

GridD.TextMatrix(0, 0) = "Prod.Term."
GridD.TextMatrix(0, 1) = "Pedido"
GridD.TextMatrix(0, 2) = "Suma Kgr"
GridD.TextMatrix(0, 3) = "Codigo"
GridD.TextMatrix(0, 4) = "Insumo"
GridD.TextMatrix(0, 5) = "Bultos"
GridD.TextMatrix(0, 6) = "Kilos"
GridD.RowHeight(1) = 560


mas:
      f = 2
      GridD.Refresh
      FF = 0
      Do Until f = GridD.Rows - 1
         FF = 0
         If Trim(GridD.TextMatrix(f, 3)) = "" Then
               GridD.RemoveItem (f)
               FF = 999
               Exit Do
         End If
         f = f + 1
      Loop
      
      If FF = 999 Then GoTo mas

Option5.Value = True
f = 2
GoSub VERIFICA
FF = GridD.Rows - 1
If FF = 2 Then FF = 1






pub_cadena = " SELECT * FROM PROCESOS WHERE PRO_CODCIA = ? AND PRO_CODPRO= ? ORDER BY PRO_ORDEN"
Set PS_FOR = CN.CreateQuery("", pub_cadena)
PS_FOR.rdoParameters(0) = LK_CODCIA
PS_FOR.rdoParameters(1) = 0
Set llave_FOR = PS_FOR.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)


ww_soles = 0
f = 2
Do Until f = GridT.Rows - 1

If Trim(GridT.TextMatrix(f, 0)) = "" Then GoTo PASALO
' camtex
'If Trim(GridT.TextMatrix(f, 6)) = "OK" Then GoTo PASALO
'GridT.TextMatrix(f, 6) = "OK"
flag_textil = 9

PS_FOR.rdoParameters(0) = LK_CODCIA
PS_FOR.rdoParameters(1) = GridT.TextMatrix(f, 5)
llave_FOR.Requery
If llave_FOR.EOF Then
'   MsgBox "No hay Insumos Relacionados..en " & GridT.TextMatrix(f, 0)
   FF = FF + 1
   GridD.Rows = FF + 1
   GridD.RowHeight(FF) = 320
   GridD.TextMatrix(FF, 0) = GridT.TextMatrix(f, 0)
   GridD.TextMatrix(FF, 1) = GridT.TextMatrix(f, 3)
   GridD.TextMatrix(FF, 4) = ""
   GridD.TextMatrix(FF, 3) = ""
   GridD.TextMatrix(FF, 5) = ""
   GridD.TextMatrix(FF, 6) = ""
   GoTo PASALO
End If

Do Until llave_FOR.EOF
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_KEY = llave_FOR!PRO_CODART
   LEER_ART_LLAVE

    FF = FF + 1
    GridD.Rows = FF + 1
    GridD.RowHeight(FF) = 320
    GridD.TextMatrix(FF, 0) = GridT.TextMatrix(f, 0)
    GridD.TextMatrix(FF, 1) = GridT.TextMatrix(f, 3)
    GridD.TextMatrix(FF, 4) = art_LLAVE!ART_NOMBRE
    GridD.TextMatrix(FF, 3) = art_LLAVE!ART_ALTERNO
    GridD.TextMatrix(FF, 5) = redondea(Nulo_Valor0(llave_FOR!PRO_CANTIDAD) * GridT.TextMatrix(f, 3) * 1.1 / 100)
    GridD.TextMatrix(FF, 6) = art_LLAVE!ART_KEY
    ww_soles = ww_soles + Val(GridD.TextMatrix(FF, 5))
    llave_FOR.MoveNext
Loop
PASALO:
f = f + 1
Loop
GridD.TextMatrix(1, 1) = GridT.TextMatrix(1, 3)
PAS:
RECALCULA_TODO
calcula_totales24
GridD.Row = 2
GridD.COL = 4
GridD.SetFocus
Exit Sub

VERIFICA:

FF = 2
Do Until f = GridT.Rows - 1
If Trim(GridT.TextMatrix(f, 6)) = "OK" Then GoTo PASALO2

   Do Until FF = GridD.Rows
      If Trim(GridT.TextMatrix(f, 0)) = Trim(GridD.TextMatrix(FF, 0)) Then
          If FF = 2 And GridD.Rows = 3 Then
          Else
             GridD.RemoveItem (FF)
             GoTo VERIFICA
          End If
      End If
      FF = FF + 1
   Loop
PASALO2:

f = f + 1

If f = GridT.Rows - 1 Then
Else
   GoTo VERIFICA
End If
Loop
GridD.Refresh
GridD.SetFocus
Return
End Sub
Private Sub Option6_Click()
flag_textil = 10

Frame5.Visible = True
Frame4.Visible = False
Frame6.Visible = False
textovar2.Visible = False
textovar3.Visible = False
textovar.Visible = False

Dim fx As Integer
Dim DIF As Currency
Dim WS_CODART As String
Dim FF
Dim ww_cantidad As Currency
Dim ww_req As Currency
Dim FILAX As Integer
Dim subtotal As Currency
subtotal = 0
Screen.MousePointer = 11
grid_fac.Rows = 3
gridl.ColWidth(0) = 0
gridl.ColWidth(1) = 1500
gridl.ColWidth(2) = 1000
gridl.ColWidth(4) = 1000
gridl.ColWidth(5) = 1000

gridl.TextMatrix(0, 1) = "Insumo"
gridl.TextMatrix(0, 2) = "Cantidad"
gridl.TextMatrix(0, 4) = "Stock "
gridl.TextMatrix(0, 5) = "Faltante"
grid_fac.ColWidth(2) = 800
gridl.Rows = GridD.Rows
fx = 2
Do Until fx = GridD.Rows
   gridl.Rows = GridD.Rows
   gridl.TextMatrix(fx, 0) = GridD.TextMatrix(fx, 0)
   gridl.TextMatrix(fx, 1) = GridD.TextMatrix(fx, 1)
   gridl.TextMatrix(fx, 2) = GridD.TextMatrix(fx, 2)
   gridl.TextMatrix(fx, 3) = GridD.TextMatrix(fx, 3)
   gridl.TextMatrix(fx, 4) = GridD.TextMatrix(fx, 4)
   gridl.TextMatrix(fx, 5) = GridD.TextMatrix(fx, 5)
   gridl.TextMatrix(fx, 6) = GridD.TextMatrix(fx, 6)
   gridl.TextMatrix(fx, 7) = GridD.TextMatrix(fx, 7)
   fx = fx + 1
Loop

   fx = 2
   FILAX = 2
   gridl.Rows = gridl.Rows + 1
   Do Until fx = gridl.Rows - 1
      WS_CODART = Trim(gridl.TextMatrix(fx, 7))
      FF = fx + 1
      Do Until FF = gridl.Rows
         If Trim(gridl.TextMatrix(FF, 7)) = Trim(WS_CODART) And Trim(WS_CODART) <> "" Then
            gridl.TextMatrix(fx, 6) = Val(gridl.TextMatrix(fx, 6)) + Val(gridl.TextMatrix(FF, 6))
            gridl.TextMatrix(FF, 6) = ""
            gridl.TextMatrix(FF, 7) = ""
         End If
         FF = FF + 1
      Loop
      fx = fx + 1
   Loop
mas:
      fx = 2
      Do Until fx = gridl.Rows - 1
         If Trim(gridl.TextMatrix(fx, 6)) = "" Then
            gridl.RemoveItem (fx)
            fx = 999
            Exit Do
         End If
         fx = fx + 1
      Loop
      
      If fx = 999 Then GoTo mas
      
      ww_cantidad = 0
      ww_req = 0
      fx = 2
      Do Until fx = gridl.Rows - 1
         If Val(gridl.TextMatrix(fx, 6)) <> 0 Then
            SQ_OPER = 1
            PUB_CODART = Val(gridl.TextMatrix(fx, 7))
            pu_codcia = LK_CODCIA
            LEER_ARM_LLAVE
'            GridD.TextMatrix(fx, 4) = arm_llave!arm_stock
            grid_fac.Rows = grid_fac.Rows + 1
            grid_fac.TextMatrix(fx, 4) = Val(gridl.TextMatrix(fx, 6))
            grid_fac.TextMatrix(fx, 1) = gridl.TextMatrix(fx, 3)
            grid_fac.TextMatrix(fx, 2) = gridl.TextMatrix(fx, 5)
            grid_fac.TextMatrix(fx, 16) = Val(gridl.TextMatrix(fx, 7))
            grid_fac.TextMatrix(fx, 33) = Val(gridl.TextMatrix(fx, 7))
            grid_fac.TextMatrix(fx, 0) = gridl.TextMatrix(fx, 4)
            grid_fac.TextMatrix(fx, 6) = arm_llave!arm_cospro
            grid_fac.TextMatrix(fx, 13) = Format(arm_llave!ARM_STOCK, "0.00")
            grid_fac.TextMatrix(fx, 11) = arm_llave!arm_cospro
            grid_fac.TextMatrix(fx, 12) = pub_signo_arm
            grid_fac.TextMatrix(fx, 14) = 1
            
            If Val(grid_fac.TextMatrix(fx, 13)) > 0 Then
               grid_fac.TextMatrix(fx, 38) = Val(grid_fac.TextMatrix(fx, 4)) - Val(grid_fac.TextMatrix(fx, 13))
            Else
               grid_fac.TextMatrix(fx, 38) = grid_fac.TextMatrix(fx, 4)
            End If
            If grid_fac.TextMatrix(fx, 38) < 0 Then grid_fac.TextMatrix(fx, 38) = ""
            
            
            grid_fac.TextMatrix(fx, 7) = Format(Val(grid_fac.TextMatrix(fx, 4)) * Val(grid_fac.TextMatrix(fx, 6)), "0.00")
            
         End If
         fx = fx + 1
         
      Loop
      
      
         
      
calcula_totales
      
Screen.MousePointer = 0
Frame5.Visible = False
Frame4.Visible = True
textovar2.Visible = False
grid_fac.SetFocus

End Sub

Private Sub Option7_Click()
Frame5.Visible = True
GridD.Visible = False
GridT.Visible = False
Frame4.Visible = False
Frame6.Visible = True
textovar2.Visible = False
textovar3.Visible = False
textovar.Visible = False
If Text1(1).Visible = True Then Text1(1).SetFocus
End Sub

Private Sub Option8_Click()
textovar2.Visible = False
textovar3.Visible = False
textovar.Visible = False

Frame5.Visible = True
Frame4.Visible = False
Frame6.Visible = False
GridD.Visible = False
GridT.Visible = True
GridT.SetFocus
End Sub

Private Sub Picture1_Click()
Dim fx As Integer
Dim DIF As Currency
Dim WS_CODART As String
Dim FF
Dim NETO As Currency
Dim FILAX As Integer
Screen.MousePointer = 11


   NETO = Val(i_neto.Text)
   fx = 2
   FILAX = 2
   grid_fac.Rows = grid_fac.Rows + 1
   Do Until fx = grid_fac.Rows - 1
      WS_CODART = Trim(grid_fac.TextMatrix(fx, 1))
      PUB_PRECIO = Val(grid_fac.TextMatrix(fx, 6))
      FF = fx + 1
      Do Until FF = grid_fac.Rows - 1
         If Trim(grid_fac.TextMatrix(FF, 1)) = Trim(WS_CODART) And Val(grid_fac.TextMatrix(FF, 6)) = PUB_PRECIO Then
            grid_fac.TextMatrix(fx, 4) = Val(grid_fac.TextMatrix(fx, 4)) + Val(grid_fac.TextMatrix(FF, 4))
            grid_fac.TextMatrix(fx, 7) = Val(grid_fac.TextMatrix(fx, 7)) + Val(grid_fac.TextMatrix(FF, 7))
            grid_fac.TextMatrix(FF, 1) = ""
         End If
         FF = FF + 1
      Loop
      fx = fx + 1
   Loop
mas:
      fx = 2
      Do Until fx = grid_fac.Rows - 1
         If Trim(grid_fac.TextMatrix(fx, 1)) = "" Then
            grid_fac.RemoveItem (fx)
            fx = 999
            Exit Do
         End If
         fx = fx + 1
      Loop
      If fx = 999 Then GoTo mas
      calcula_totales
      DIF = NETO - Val(i_neto.Text)
      If DIF <> 0 Then
         MsgBox "REVISAR TOTALES..."
      End If
      
Screen.MousePointer = 0
      

End Sub

Private Sub salir_Click()
Unload FORM_GRIFO
LK_CODTRA = 0
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Index < 15 And Text1(Index + 1).Visible = True Then Text1(Index + 1).SetFocus
End Sub

Private Sub TEXTOVAR_CANJE_Change()
grid_canje.Text = textovar_canje.Text
End Sub

Private Sub textovar_canje_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   textovar_canje.Text = textovar_bak
End If


If (KeyCode > 36 And KeyCode < 41) Or KeyCode = 13 Then
Else
   Exit Sub
End If


   If KeyCode = 38 Then
      If grid_canje.Row = 2 Then Exit Sub
   End If
   
   If KeyCode = 37 Then
      If grid_canje.COL = 0 Then Exit Sub
   End If
   If KeyCode = 40 Then
      If grid_canje.Row = grid_canje.Rows - 1 Then Exit Sub
   End If
   
If KeyCode = 13 Then
   If grid_canje.COL = 1 Or grid_canje.COL = 2 Then
   Else
   Exit Sub
   End If
End If

SALTO:

grid_canje.TextMatrix(grid_canje.Row, grid_canje.COL) = textovar_canje.Text

textovar_canje.Visible = False
grid_canje.SetFocus



If KeyCode = 38 Then
   grid_canje.Row = grid_canje.Row - 1
ElseIf KeyCode = 40 Then
   grid_canje.Row = grid_canje.Row + 1
End If


If KeyCode = 37 Then
   grid_canje.COL = grid_canje.COL - 1
ElseIf KeyCode = 39 Or KeyCode = 13 Then
      If grid_canje.COL = 7 Then grid_canje.COL = 2
      grid_canje.COL = grid_canje.COL + 1
End If



End Sub

Private Sub TEXTOVAR_CANJE_KeyPress(KeyAscii As Integer)

'If Len(Trim(letche.ToolTipText)) <> 2 Then
'   MsgBox "Selecccionar Tipo Doc."
'   KeyAscii = 0
'   Exit Sub
'End If


If KeyAscii <> 13 Then Exit Sub


'If Len(Trim(letche.ToolTipText)) <> 2 Then
'   MsgBox "Selecccionar Tipo Doc."
'   textovar_canje.text = ""
'   Exit Sub
'End If
'If grid_canje.Col = 3 Then CALCULA_TOTALES_CANJE

'If grid_canje.Col = 3 Then
'   grid_canje.Col = 4
'   Exit Sub
'End If
'If grid_canje.Col = 4 Then
'   If IsDate(textovar_canje.text) = False Then
'      MsgBox "Fecha Incorrecta..."
'       Exit Sub
'   End If
'   grid_canje.Col = 5
'   Exit Sub
'End If
'If grid_canje.Col = 5 Then
'   grid_canje.Col = 6
'   Exit Sub
'End If
'If grid_canje.Col = 6 Then
'   grid_canje.Col = 7
'   Exit Sub
'End If


'If grid_canje.Col = 7 Then
'If grid_canje.Row = grid_canje.Rows - 1 And LK_CODTRA <> 2770 Then
'      grid_canje.Rows = grid_canje.Rows + 1
'      grid_canje.Row = grid_canje.Rows - 1
'      grid_canje.RowHeight(grid_canje.Row) = 350
'      grid_canje.Col = 3
'Else
'    grid_canje.Col = 3
'    grid_canje.Row = grid_canje.Row + 1
'End If
'End If

End Sub
Private Sub textovar_Change()

grid_fac.Text = textovar.Text
' ACV
If grid_fac.COL = 4 Then
   grid_fac.TextMatrix(grid_fac.Row, 7) = redondea(Val(grid_fac.TextMatrix(grid_fac.Row, 35)) * Val(textovar.Text))
End If
      

'If grid_fac.COL = 10 And PUB_TIPMOV = 10 Then
'   grid_fac.TextMatrix(grid_fac.Row, 6) = redondea(Val(grid_fac.TextMatrix(grid_fac.Row, 35)) * (100 - Val(textovar.Text)) / 100)
'   grid_fac.TextMatrix(grid_fac.Row, 7) = redondea(Val(grid_fac.TextMatrix(grid_fac.Row, 6)) * Val(grid_fac.TextMatrix(grid_fac.Row, 4)))
'   calcula_totales
'End If


'If grid_fac.COL = 7 And Val(textovar.Text) > 0 Then
'   grid_fac.TextMatrix(grid_fac.Row, 25) = "1"
'End If
'If grid_fac.COL = 6 And Val(textovar.Text) > 0 Then grid_fac.TextMatrix(grid_fac.Row, 25) = ""
If LK_CODTRA = 1402 Then calcula_totales
End Sub
Private Sub textovar_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tf
Dim WSPOR As Currency
Dim fil As Integer
Dim celda_anterior
' CAMTEX
On Error GoTo SAS
If KeyCode = 13 Then
  If grid_fac.COL = 1 And Val(textovar.Text) = 0 Then Exit Sub
End If

flag_salto = 1
If grid_fac.COL = 1 Then
  If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
    If textovar.Text = "" Then
    Else
     GoTo BUSCANDO
    End If
  Else
   If IsNumeric(textovar.Text) = True Or textovar.Text = "" Then
   Else
     GoTo BUSCANDO
  End If
End If
End If

If KeyCode = 27 Then
   textovar.Text = textovar_bak
End If


If (KeyCode > 36 And KeyCode < 41 And textovar.Text = "") Or KeyCode = 13 Then
Else
   Exit Sub
End If


   If KeyCode = 38 Then
      If grid_fac.Row = 2 Then Exit Sub
   End If
   
   If KeyCode = 37 Then
      If grid_fac.COL = 0 Then Exit Sub
   End If
   If KeyCode = 40 Then
      If grid_fac.Row = grid_fac.Rows - 1 Then Exit Sub
   End If
   
If KeyCode = 13 Then
   If grid_fac.COL = 1 Or grid_fac.COL = 2 Or grid_fac.COL = 3 Or grid_fac.COL = 4 Or grid_fac.COL = 5 Or grid_fac.COL = 6 Or grid_fac.COL = 7 Or grid_fac.COL = 8 Or grid_fac.COL = 10 Or grid_fac.COL = 15 Then
   Else
   Exit Sub
   End If
End If


If PUB_TIPMOV = 10 Then
   If Trim(i_fbg.Text) = "F" Then
      If Nulo_Valor0(par_llave!par_fac_lines) < grid_fac.Row - 1 Then
         MsgBox "Tope de lineas alcanzado..."
         grid_fac.TextMatrix(grid_fac.Row, 4) = ""
         grid_fac.TextMatrix(grid_fac.Row, 5) = ""
         grid_fac.TextMatrix(grid_fac.Row, 6) = ""
         grid_fac.TextMatrix(grid_fac.Row, 7) = ""
         Exit Sub
      End If
    ElseIf Trim(i_fbg.Text) = "B" Then
      If Nulo_Valor0(par_llave!par_BOL_lines) < grid_fac.Row - 1 Then
         MsgBox "Tope de lineas alcanzado..."
         grid_fac.TextMatrix(grid_fac.Row, 4) = ""
         grid_fac.TextMatrix(grid_fac.Row, 5) = ""
         grid_fac.TextMatrix(grid_fac.Row, 6) = ""
         grid_fac.TextMatrix(grid_fac.Row, 7) = ""
         Exit Sub
      End If
End If
End If

If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
      If Nulo_Valor0(par_llave!par_NOTAS_lines) < grid_fac.Row - 1 Then
         MsgBox "Tope de lineas alcanzado..."
         grid_fac.TextMatrix(grid_fac.Row, 4) = ""
         grid_fac.TextMatrix(grid_fac.Row, 5) = ""
         grid_fac.TextMatrix(grid_fac.Row, 6) = ""
         grid_fac.TextMatrix(grid_fac.Row, 7) = ""
         Exit Sub
      End If
End If


If grid_fac.COL = 1 Then
   If textovar.Text = "" And KeyCode <> 13 Then
         grid_fac.TextMatrix(grid_fac.Row, 1) = ""
         grid_fac.TextMatrix(grid_fac.Row, 4) = ""
         grid_fac.TextMatrix(grid_fac.Row, 5) = ""
         grid_fac.TextMatrix(grid_fac.Row, 6) = ""
         grid_fac.TextMatrix(grid_fac.Row, 7) = ""
         label_nomart.Caption = ""
         calcula_totales
      GoTo SALTO
   End If
Else
   GoTo SALTO
End If

    SQ_OPER = 1
    On Error GoTo OJO
    PUB_KEY = Val(textovar.Text)
    On Error GoTo 0
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       grid_fac.TextMatrix(grid_fac.Row, 0) = ""
       label_nomart.Caption = ""
       GoTo fin
    Else
       If LK_CODTRA = 2101 Then
        If art_LLAVE!art_familia = 1 Then
          tf = 1
          Do Until tf = 20
             If Val(tab_items(Val(i_codven.Text), tf)) = art_LLAVE(0) Then Exit Do
             tf = tf + 1
          Loop
          If tf = 20 Then
             MsgBox "Producto no pertence a la Isla ... ", 48, Pub_Titulo
             textovar.Text = ""
             Exit Sub
          End If
        End If
       End If

       grid_fac.TextMatrix(grid_fac.Row, 0) = art_LLAVE(2)
    End If
    grid_fac.TextMatrix(grid_fac.Row, 16) = art_LLAVE(0)
    
    SQ_OPER = 1
    PUB_CODART = Val(textovar.Text)
    pu_codcia = LK_CODCIA
    LEER_ARM_LLAVE
    grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_STOCK
    
If Nulo_Valors(SUT_LLAVE!SUT_UNIDADES) = "A" And pub_signo_arm = -1 Then
   If arm_llave!ARM_STOCK <= 0 Then
      MsgBox "Ojo Stock no disponible " & "Existencia : " & Val(grid_fac.TextMatrix(grid_fac.Row, 13))
      If LK_USU_STOCK <> "A" Then
         grid_fac.TextMatrix(grid_fac.Row, 1) = ""
         grid_fac.TextMatrix(grid_fac.Row, 4) = ""
         grid_fac.TextMatrix(grid_fac.Row, 5) = ""
         grid_fac.TextMatrix(grid_fac.Row, 6) = ""
         grid_fac.TextMatrix(grid_fac.Row, 7) = ""
         textovar.Text = ""
         Exit Sub
      End If
   End If
End If

    
    
    
grid_fac.TextMatrix(grid_fac.Row, 16) = art_LLAVE!ART_KEY
grid_fac.TextMatrix(grid_fac.Row, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
grid_fac.TextMatrix(grid_fac.Row, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
grid_fac.TextMatrix(grid_fac.Row, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)

VERIFICA_REPET
grid_fac.TextMatrix(grid_fac.Row, 11) = arm_llave!arm_cospro
grid_fac.TextMatrix(grid_fac.Row, 14) = 1
'REVISAR ....24/11/99
If Val(grid_fac.TextMatrix(grid_fac.Row, 12)) = 0 Then grid_fac.TextMatrix(grid_fac.Row, 12) = pub_signo_arm


grid_fac.CellForeColor = vbBlack
If Nulo_Valors(SUT_LLAVE!SUT_PRECIO) = "1" Then
   grid_fac.TextMatrix(grid_fac.Row, 6) = arm_llave!arm_cospro
End If
 
verifica_unidades

verifica_precios

grid_fac.TextMatrix(grid_fac.Row, 11) = arm_llave!arm_cospro


SALTO:

grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL) = textovar.Text

If grid_fac.COL <> 1 Then
If grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL) <> textovar_bak Then
   calcula_totales
End If
End If
PEPE:
textovar.Visible = False
grid_fac.SetFocus
If grid_fac.COL = 6 And KeyCode = 13 Then Exit Sub

If grid_fac.COL = 10 And PUB_TIPMOV = 10 Then
   If Val(grid_fac.TextMatrix(grid_fac.Row, 6)) < Val(grid_fac.TextMatrix(grid_fac.Row, 32)) Then
      MsgBox "Precio Minimo es  : " & Val(grid_fac.TextMatrix(grid_fac.Row, 32)), 48, Pub_Titulo
      grid_fac.TextMatrix(grid_fac.Row, 10) = ""
      grid_fac.SetFocus
      grid_fac_KeyPress 13
      Exit Sub
   End If
End If


If KeyCode = 38 Then
   grid_fac.Row = grid_fac.Row - 1
ElseIf KeyCode = 40 Then
   grid_fac.Row = grid_fac.Row + 1
End If

If LK_CODTRA = 2107 Then GoTo fin

If KeyCode = 37 Then
   grid_fac.COL = tab_izquierda(grid_fac.COL)
ElseIf KeyCode = 39 Or KeyCode = 13 Then
    If tab_derecha(grid_fac.COL) = 1 And grid_fac.TextMatrix(grid_fac.Row, 16) <> "" Then
      flag_salto = 1
      ' CAMTEX
      If LK_CODTRA = 2211 Then
         If grid_fac.COL < 7 Then
            If grid_fac.COL = 4 Then
              grid_fac.COL = 6
            Else
              grid_fac.COL = grid_fac.COL + 1
            End If
            GoTo fin
         End If
      End If
      If grid_fac.Row = grid_fac.Rows - 1 Then
         grid_fac.Rows = grid_fac.Rows + 1
      End If
      grid_fac.Row = grid_fac.Row + 1
      grid_fac.RowHeight(grid_fac.Row) = 315
      flag_salto = 0
      grid_fac.COL = 1
    Else
      grid_fac.COL = tab_derecha(grid_fac.COL)
    End If
    
End If
GoTo fin

BUSCANDO:
If grid_fac.COL <> 1 Then Exit Sub

Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not LV_ART.Visible Then
 Exit Sub
End If

If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And textovar.Text = "" Then
  loc_key = 1
  Set LV_ART.SelectedItem = LV_ART.ListItems(loc_key)
'  LV_CLI.Visible = False
  LV_ART.ListItems.Item(loc_key).Selected = True
  LV_ART.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_ART.ListItems.count Then loc_key = LV_ART.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_ART.ListItems.count Then loc_key = LV_ART.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
'  KeyCode = 0
  LV_ART.ListItems.Item(loc_key).Selected = True
  LV_ART.ListItems.Item(loc_key).EnsureVisible
  textovar.Text = Trim(LV_ART.ListItems.Item(loc_key).Text) & " "
  DoEvents
  
  textovar.SelStart = Len(textovar.Text)
  DoEvents
fin:
'grid_fac.COL = 6
'grid_fac.SetFocus
OJO:
SAS:

End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i, car
Dim itmFound As ListItem    ' Variable FoundItem.
car = Chr(KeyAscii)
car = UCase(car)
KeyAscii = Asc(car)
If KeyAscii = 27 Then
  textovar.Text = ""
  LV_ART.Visible = False
  Exit Sub
End If


If KeyAscii <> 13 Then pub_ojo = "A"

If grid_fac.COL = 3 Then
 Consistencias grid_fac, textovar, KeyAscii, 1, 3
 If Len(textovar.Text) > 13 Then
    KeyAscii = 0
    Exit Sub
 End If
End If

If grid_fac.COL = 38 And KeyAscii = 13 Then
   textovar.Visible = False
   grid_fac.SetFocus
   Exit Sub
End If



If grid_fac.COL = 2 Then
 If Len(textovar.Text) > 13 Then
    KeyAscii = 0
    Exit Sub
 End If
 Consistencias grid_fac, textovar, KeyAscii, 2, 2
End If
If grid_fac.COL = 4 Then
 If Len(textovar.Text) > 13 Then
    KeyAscii = 0
    Exit Sub
 End If
  Consistencias grid_fac, textovar, KeyAscii, 2, 4
End If
 
'Else
'textovar.MaxLength = 0
'End If
If grid_fac.COL = 5 Then Consistencias grid_fac, textovar, KeyAscii, 2, 5
If grid_fac.COL = 6 Then Consistencias grid_fac, textovar, KeyAscii, 2, 6
If grid_fac.COL = 7 Then Consistencias grid_fac, textovar, KeyAscii, 2, 6
If grid_fac.COL = 8 Then Consistencias grid_fac, textovar, KeyAscii, 2, 8
If grid_fac.COL = 10 Then Consistencias grid_fac, textovar, KeyAscii, 2, 6
If grid_fac.COL = 38 Then Consistencias grid_fac, textovar, KeyAscii, 2, 6


If grid_fac.COL = 1 And KeyAscii <> 13 Then
   label_nomart.Caption = ""
   grid_fac.TextMatrix(grid_fac.Row, 0) = ""
   grid_fac.TextMatrix(grid_fac.Row, 4) = ""
   grid_fac.TextMatrix(grid_fac.Row, 5) = ""
   grid_fac.TextMatrix(grid_fac.Row, 6) = ""
   grid_fac.TextMatrix(grid_fac.Row, 7) = ""
   grid_fac.TextMatrix(grid_fac.Row, 16) = ""
   calcula_totales
End If

If KeyAscii <> 13 Then
   textovar.ForeColor = vbRed
End If

If KeyAscii <> 13 Then
   GoTo fin
End If

flag_salto = 0
If grid_fac.COL <> 1 Then Exit Sub
'solo para codart....
If grid_fac.TextMatrix(grid_fac.Row, 1) <> textovar.Text Then Exit Sub

 If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  PUB_KEY = 0
 Else
 On Error GoTo OJO
 PUB_KEY = Val(textovar.Text)
 On Error GoTo 0
 If Len(textovar.Text) = 0 Then
    Exit Sub
 End If
 If IsNumeric(textovar.Text) = False Then
   PUB_KEY = 0
 End If
End If

If PUB_KEY <> 0 Then
    SQ_OPER = 1
    PUB_KEY = textovar.Text
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul textovar, textovar
       GoTo fin
    Else
      If LK_CODTRA = 2101 Then
        If art_LLAVE!art_familia = 1 Then
          tf = 1
          Do Until tf = 20
             If Val(tab_items(Val(i_codven.Text), tf)) = art_LLAVE(0) Then Exit Do
             tf = tf + 1
          Loop
          If tf = 20 Then
             MsgBox "Producto no pertence a la Isla ... ", 48, Pub_Titulo
             textovar.Text = ""
             Exit Sub
          End If
       End If
      End If
   End If
Else
  If VAR_ACTIVAR <> 99 And textovar <> "" And LK_FLAG_ORIGINAL <> "A" And LK_FLAG_ALTERNO = "A" Then
     SQ_OPER = 3
     pu_alterno = textovar.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul textovar, textovar
       Exit Sub
     Else
       If PUB_TIPMOV = 20 Then
          If SUT_LLAVE!SUT_SECUENCIA = 1 Then
             If art_llave_alt!art_flag_stock <> "M" Then
                MsgBox "No es Mercaderias ... ", 48, Pub_Titulo
                textovar.Text = ""
                Exit Sub
             End If
          ElseIf art_llave_alt!art_flag_stock <> "A" Then
                MsgBox "No es Activo Fijo ... ", 48, Pub_Titulo
                textovar.Text = ""
                Exit Sub
        End If
        
     End If
     End If
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
      PUB_KEY = art_llave_alt!ART_KEY
      LEER_ART_LLAVE
      If PUB_TIPMOV = 75 Then
         If Val(Nulo_Valor0(art_LLAVE!ART_CODART2)) = 0 Then
            MsgBox "Articulo no esta relacionado..."
            Azul textovar, textovar
            Exit Sub
         End If
      End If
      If LK_CODTRA = 2101 Then
        If art_LLAVE!art_familia = 1 Then
           tf = 1
           Do Until tf = 20
              If Val(tab_items(Val(i_codven.Text), tf)) = art_LLAVE(0) Then Exit Do
              tf = tf + 1
           Loop
           If tf = 20 Then
              MsgBox "Producto no pertence a la Isla ... ", 48, Pub_Titulo
              textovar.Text = ""
              Exit Sub
           End If
        End If
       End If
  Else
    If loc_key > LV_ART.ListItems.count Or loc_key = 0 Then
     Exit Sub
    End If
    valor = UCase(LV_ART.ListItems.Item(loc_key).Text)
    If Trim(UCase(textovar.Text)) = Left(valor, Len(Trim(textovar.Text))) And Len(Trim(textovar.Text)) <> 0 Then
      If VAR_ACTIVAR <> 99 Then
      textovar.Text = Trim(LV_ART.ListItems.Item(loc_key).SubItems(1))
      Else
      textovar.Text = Trim(LV_ART.ListItems.Item(loc_key))
      End If
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       PUB_KEY = Val(LV_ART.ListItems.Item(loc_key).SubItems(2))
      Else
       PUB_KEY = textovar.Text
      End If
      LEER_ART_LLAVE
      If LK_CODTRA = 2101 Then
        If art_LLAVE!art_familia = 1 Then
          tf = 1
          Do Until tf = 20
             If Val(tab_items(Val(i_codven.Text), tf)) = art_LLAVE(0) Then Exit Do
             tf = tf + 1
          Loop
          If tf = 20 Then
             MsgBox "Producto no pertence a la Isla ... ", 48, Pub_Titulo
             textovar.Text = ""
             Exit Sub
          End If
        End If
      End If

      VAR_ACTIVAR = 0
    Else
      Exit Sub
    End If
  End If
End If
    SQ_OPER = 1
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       If VAR_ACTIVAR <> 99 Then
          PUB_CODART = PUB_KEY
       Else
          PUB_CODART = Val(LV_ART.ListItems.Item(loc_key).SubItems(2))
       End If
    Else
       PUB_CODART = Val(textovar.Text)
    End If
    pu_codcia = LK_CODCIA
    LEER_ARM_LLAVE
    
   If arm_llave!ARM_STOCK <= 0 And pub_signo_arm = -1 Then
      MsgBox "Ojo Stock no disponible " & "Existencia : " & arm_llave!ARM_STOCK
      If LK_USU_STOCK <> "A" Then
         grid_fac.TextMatrix(grid_fac.Row, 3) = ""
         grid_fac.TextMatrix(grid_fac.Row, 4) = ""
         grid_fac.TextMatrix(grid_fac.Row, 5) = ""
         grid_fac.TextMatrix(grid_fac.Row, 6) = ""
         grid_fac.TextMatrix(grid_fac.Row, 7) = ""
         textovar.Text = ""
         Exit Sub
      End If
  End If
    
    

verifica_precios

If Nulo_Valors(SUT_LLAVE!SUT_PRECIO) = "1" Then
   grid_fac.TextMatrix(grid_fac.Row, 6) = Nulo_Valor0(arm_llave!arm_cospro)
End If
grid_fac.CellForeColor = vbBlack
grid_fac.TextMatrix(grid_fac.Row, 0) = art_LLAVE!ART_NOMBRE
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
   If VAR_ACTIVAR <> 99 Then
   Else
    grid_fac.TextMatrix(grid_fac.Row, 1) = LV_ART.ListItems.Item(loc_key)
   End If
Else
  grid_fac.TextMatrix(grid_fac.Row, 1) = art_LLAVE!ART_KEY
End If
grid_fac.TextMatrix(grid_fac.Row, 11) = Nulo_Valor0(arm_llave!arm_cospro)
grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_STOCK
grid_fac.TextMatrix(grid_fac.Row, 16) = art_LLAVE!ART_KEY
grid_fac.TextMatrix(grid_fac.Row, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
grid_fac.TextMatrix(grid_fac.Row, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
grid_fac.TextMatrix(grid_fac.Row, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
VERIFICA_REPET

If Val(grid_fac.TextMatrix(grid_fac.Row, 12)) = 0 Then grid_fac.TextMatrix(grid_fac.Row, 12) = pub_signo_arm


grid_fac.TextMatrix(grid_fac.Row, 14) = 1


grid_fac.COL = tab_derecha(1)

LV_ART.Visible = False
VAR_ACTIVAR = 0
'grid_fac.SetFocus  2210 no esta textovar
fin:
OJO:

End Sub
Private Sub textovar_KeyUp(KeyCode As Integer, Shift As Integer)
'' CHEQUEAR AQUI...
Dim VAR
Dim WS_TIPO As Integer
Dim WS_CALIDAD As Integer
Dim ws_codcia As String

If grid_fac.COL = 1 And KeyCode = 46 Then
   label_nomart.Caption = ""
   grid_fac.TextMatrix(grid_fac.Row, 0) = ""
   grid_fac.TextMatrix(grid_fac.Row, 5) = ""
   grid_fac.TextMatrix(grid_fac.Row, 6) = ""
End If

If KeyCode > 35 And KeyCode < 41 Then
   Exit Sub
End If
If grid_fac.COL <> 1 Then Exit Sub
' puede estar pruebas
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
 If Len(textovar.Text) = 0 Or textovar.Text = "" Then
   VAR_ACTIVAR = 0
   LV_ART.Visible = False
   Exit Sub
 End If
 If textovar.Text = "*" And KeyCode = 106 Then
   VAR_ACTIVAR = 99
   Exit Sub
 ElseIf textovar.Text = "" Then
   VAR_ACTIVAR = 0
   Exit Sub
 End If
 If VAR_ACTIVAR <> 99 Then
   Exit Sub
 End If
 If Left(textovar.Text, 1) = "*" Then
   textovar.Text = Mid(textovar.Text, 2, Len(textovar.Text))
   textovar.SelStart = Len(textovar.Text)
 End If

Else
 If Len(textovar.Text) = 0 Or IsNumeric(textovar.Text) = True Then
   LV_ART.Visible = False
   Exit Sub
 End If
End If
If LV_ART.Visible = False And KeyCode <> 13 Or Len(textovar.Text) = 1 Then
'If LV_ART.Visible = False And IsNumeric(textovar.text) = False Then
    If textovar.Text = "" Then Exit Sub
    VAR = Asc(textovar.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    WS_TIPO = Val(Nulo_Valor0(SUT_LLAVE!SUT_art_gru))
    If LK_CODTRA = 2103 Then WS_TIPO = 1
    If LK_CODTRA = 2210 Then WS_TIPO = Val(Right(i_destino.Text, 5))
    WS_CALIDAD = Val(Nulo_Valor0(SUT_LLAVE!SUT_CALIDAD))
    ws_codcia = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       ws_codcia = "00"
    End If
    
    
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      'archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO  FROM ARTI WHERE  ART_CODCIA = '" & LK_CODCIA & "' AND ART_ALTERNO BETWEEN '" & textovar.text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
      If WS_TIPO = 0 And WS_CALIDAD = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO  FROM ARTI WHERE  ART_KEY <> 0 AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & textovar.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
      If WS_CALIDAD = 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO  FROM ARTI WHERE  ART_KEY <> 0 AND ART_FAMILIA = " & WS_TIPO & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & textovar.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
      If WS_CALIDAD <> 0 And WS_TIPO = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_KEY <> 0 AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & textovar.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
      If WS_CALIDAD <> 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_KEY <> 0 AND ART_FAMILIA = " & WS_TIPO & " AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & textovar.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      If WS_TIPO = 0 And WS_CALIDAD = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO  FROM ARTI WHERE  ART_KEY <> 0 AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & textovar.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
      If WS_CALIDAD = 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO  FROM ARTI WHERE  ART_KEY <> 0 AND ART_FAMILIA = " & WS_TIPO & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & textovar.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
      If WS_CALIDAD <> 0 And WS_TIPO = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_KEY <> 0 AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & textovar.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
      If WS_CALIDAD <> 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_KEY <> 0 AND ART_FAMILIA = " & WS_TIPO & " AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & textovar.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
    End If
    
    
     PROC_LISVIEW LV_ART
     loc_key = 0
     If LV_ART.Visible Then
      loc_key = 1
     End If
    Exit Sub
End If

'If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
' Exit Sub
'End If

Dim itmFound As ListItem    ' Variable FoundItem.
If LV_ART.Visible Then
  Set itmFound = LV_ART.FindItem(LTrim(textovar.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > LV_ART.ListItems.count Then
      LV_ART.ListItems.Item(LV_ART.ListItems.count).EnsureVisible
   Else
     LV_ART.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If

End Sub
Private Sub textovar_LostFocus()
calcula_totales
End Sub

Private Sub textovar2_Change()
If GridT.Text <> textovar2.Text And LK_CODTRA <> 2103 Then
   GridT.TextMatrix(GridT.Row, 6) = ""
   grid_fac.Clear
   pasa_cabeza
   GridD.Clear
   flag_textil = 8
End If

GridT.Text = textovar2.Text
calcula_totales22
End Sub

Private Sub textovar2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tf
Dim WSPOR As Currency
Dim fil As Integer
Dim celda_anterior

If KeyCode = 13 Then
  If GridT.COL = 1 And Val(textovar2.Text) = 0 Then Exit Sub
End If

flag_salto = 1
If GridT.COL = 1 Then
  If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
    If textovar2.Text = "" Then
    Else
     GoTo BUSCANDO
    End If
  Else
   If IsNumeric(textovar2.Text) = True Or textovar2.Text = "" Then
   Else
     GoTo BUSCANDO
  End If
End If
End If


If (KeyCode > 36 And KeyCode < 41 And textovar2.Text = "") Or KeyCode = 13 Then
Else
   Exit Sub
End If


   If KeyCode = 38 Then
      If GridT.Row = 1 Then Exit Sub
   End If
   
   If KeyCode = 37 Then
      If GridT.COL = 0 Then Exit Sub
   End If
   If KeyCode = 40 Then
      If GridT.Row = GridT.Rows - 1 Then Exit Sub
   End If
   
'If KeyCode = 13 Then
'   If GridT.COL = 1 Or GridT.COL = 2 Then
'   Else
'   Exit Sub
'   End If
'End If


If GridT.COL = 1 Then
   If textovar2.Text = "" And KeyCode <> 13 Then
         GridT.TextMatrix(GridT.Row, 1) = ""
         GridT.TextMatrix(GridT.Row, 5) = ""
         calcula_totales22
      GoTo SALTO
   End If
Else
   GoTo SALTO
End If

    SQ_OPER = 1
    On Error GoTo OJO
    PUB_KEY = Val(textovar2.Text)
    On Error GoTo 0
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       GridT.TextMatrix(GridT.Row, 0) = ""
       GoTo fin
    Else
       GridT.TextMatrix(GridT.Row, 0) = art_LLAVE(2)
    End If
'    GridT.TextMatrix(GridT.Row, 16) = art_LLAVE(0)
    
    SQ_OPER = 1
    PUB_CODART = Val(textovar2.Text)
    pu_codcia = LK_CODCIA
    LEER_ARM_LLAVE
    GridT.TextMatrix(GridT.Row, 4) = arm_llave!ARM_STOCK
    
   
    
GridT.TextMatrix(GridT.Row, 5) = art_LLAVE!ART_KEY

VERIFICA_REPET2
    
GridT.CellForeColor = vbBlack

'verifica_precios

SALTO:
If GridT.TextMatrix(GridT.Row, 3) = "" And GridT.COL = 2 And GridT.TextMatrix(GridT.Row, 2) <> "" Then
   GridT.TextMatrix(GridT.Row, 3) = GridT.TextMatrix(GridT.Row, 2) * 20
End If

GridT.TextMatrix(GridT.Row, GridT.COL) = textovar2.Text

'If GridT.COL <> 1 Then
'If GridT.TextMatrix(GridT.Row, GridT.COL) <> textovar_bak Then
'   calcula_totales
'End If
'End If
PEPE:
textovar2.Visible = False
GridT.SetFocus

'If grid_fac.COL = 10 And PUB_TIPMOV = 10 Then
'   If Val(grid_fac.TextMatrix(grid_fac.Row, 6)) < Val(grid_fac.TextMatrix(grid_fac.Row, 32)) Then
'      MsgBox "Precio Minimo es  : " & Val(grid_fac.TextMatrix(grid_fac.Row, 32)), 48, Pub_Titulo
'      grid_fac.TextMatrix(grid_fac.Row, 10) = ""
'      grid_fac.SetFocus
'      grid_fac_KeyPress 13
'      Exit Sub
'   End If
'End If



If KeyCode = 38 Then
   GridT.Row = GridT.Row - 1
   Exit Sub
ElseIf KeyCode = 40 Then
   GridT.Row = GridT.Row + 1
   Exit Sub
End If


If KeyCode = 37 Then
   If GridT.COL > 0 Then GridT.COL = GridT.COL - 1
   Exit Sub
ElseIf KeyCode = 39 Then
   If GridT.COL < 3 Then GridT.COL = GridT.COL + 1
   Exit Sub
End If


If KeyCode = 13 Then
   If GridT.COL = 1 Or GridT.COL = 2 Then
      GridT.COL = GridT.COL + 1
      Exit Sub
   End If
   If GridT.COL = 3 Then
     If PUB_TIPMOV = 177 Then
     GoTo IR
     Else
      GridT.COL = 7
     End If
      Exit Sub
   End If
   If GridT.COL = 6 Then
      GridT.COL = 7
      Exit Sub
   End If
   If GridT.COL = 7 And GridT.TextMatrix(GridT.Row, 5) <> "" Then
IR:
          GridT.COL = 1
          flag_salto = 1
          If GridT.Row = GridT.Rows - 1 Then GridT.Rows = GridT.Rows + 1
          GridT.Row = GridT.Row + 1
          GridT.RowHeight(GridT.Row) = 315
          flag_salto = 0
          GridT.COL = 1
   End If
End If

GoTo fin

BUSCANDO:
If GridT.COL <> 1 Then Exit Sub

Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not LV_ART2.Visible Then
 Exit Sub
End If

If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And textovar2.Text = "" Then
  loc_key = 1
  Set LV_ART2.SelectedItem = LV_ART2.ListItems(loc_key)
'  LV_CLI.Visible = False
  LV_ART2.ListItems.Item(loc_key).Selected = True
  LV_ART2.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_ART2.ListItems.count Then loc_key = LV_ART2.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_ART2.ListItems.count Then loc_key = LV_ART2.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
'  KeyCode = 0
  LV_ART2.ListItems.Item(loc_key).Selected = True
  LV_ART2.ListItems.Item(loc_key).EnsureVisible
  textovar2.Text = Trim(LV_ART2.ListItems.Item(loc_key).Text) & " "
  DoEvents
  
  textovar2.SelStart = Len(textovar2.Text)
  DoEvents
fin:
'grid_fac.COL = 6
'grid_fac.SetFocus
OJO:

End Sub

Private Sub textovar2_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i, car
Dim itmFound As ListItem    ' Variable FoundItem.
car = Chr(KeyAscii)
car = UCase(car)
KeyAscii = Asc(car)
If KeyAscii = 27 Then
  textovar2.Text = ""
  LV_ART2.Visible = False
  Exit Sub
End If

If KeyAscii <> 13 Then pub_ojo = "A"


If GridT.COL = 3 Then
 Consistencias GridT, textovar2, KeyAscii, 2, 3
 If Len(textovar2.Text) > 13 Then
    KeyAscii = 0
    Exit Sub
 End If
End If


If GridT.COL = 2 Then
 If Len(textovar2.Text) > 13 Then
    KeyAscii = 0
    Exit Sub
 End If
 Consistencias GridT, textovar2, KeyAscii, 2, 2
End If
If GridT.COL = 2 Then
 If Len(textovar2.Text) > 13 Then
    KeyAscii = 0
    Exit Sub
 End If
  Consistencias GridT, textovar2, KeyAscii, 2, 4
End If
 
'Else
'textovar.MaxLength = 0
'End If
If GridT.COL = 5 Then Consistencias GridT, textovar2, KeyAscii, 2, 5
If GridT.COL = 6 Then Consistencias GridT, textovar2, KeyAscii, 2, 6
If GridT.COL = 7 Then Consistencias GridT, textovar2, KeyAscii, 2, 6
If GridT.COL = 8 Then Consistencias GridT, textovar2, KeyAscii, 2, 8
If GridT.COL = 10 Then Consistencias GridT, textovar2, KeyAscii, 2, 6


If GridT.COL = 1 And KeyAscii <> 13 Then
   GridT.TextMatrix(GridT.Row, 0) = ""
   GridT.TextMatrix(GridT.Row, 1) = ""
   GridT.TextMatrix(GridT.Row, 2) = ""
   calcula_totales22
End If

If KeyAscii <> 13 Then
   textovar2.ForeColor = vbRed
End If

If KeyAscii <> 13 Then
   GoTo fin
End If

flag_salto = 0
If GridT.COL <> 1 Then Exit Sub
If GridT.TextMatrix(GridT.Row, 1) <> textovar2.Text Then Exit Sub

 If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  PUB_KEY = 0
 Else
 On Error GoTo OJO
 PUB_KEY = Val(textovar2.Text)
 On Error GoTo 0
 If Len(textovar2.Text) = 0 Then
    Exit Sub
 End If
 If IsNumeric(textovar2.Text) = False Then
   PUB_KEY = 0
 End If
End If

If PUB_KEY <> 0 Then
    SQ_OPER = 1
    PUB_KEY = textovar2.Text
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul textovar2, textovar2
       GoTo fin
    End If
Else
  If VAR_ACTIVAR <> 99 And textovar2 <> "" And LK_FLAG_ORIGINAL <> "A" And LK_FLAG_ALTERNO = "A" Then
     SQ_OPER = 3
     pu_alterno = textovar2.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul textovar2, textovar2
       Exit Sub
     Else
       If PUB_TIPMOV = 20 Then
          If SUT_LLAVE!SUT_SECUENCIA = 1 Then
             If art_llave_alt!art_flag_stock <> "M" Then
                MsgBox "No es Mercaderias ... ", 48, Pub_Titulo
                textovar2.Text = ""
                Exit Sub
             End If
          ElseIf art_llave_alt!art_flag_stock <> "A" Then
                MsgBox "No es Activo Fijo ... ", 48, Pub_Titulo
                textovar2.Text = ""
                Exit Sub
        End If
        
     End If
     End If
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
      PUB_KEY = art_llave_alt!ART_KEY
      LEER_ART_LLAVE
      If PUB_TIPMOV = 75 Then
         If Val(Nulo_Valor0(art_LLAVE!ART_CODART2)) = 0 Then
            MsgBox "Articulo no esta relacionado..."
            Azul textovar2, textovar2
            Exit Sub
         End If
      End If
  Else
    If loc_key > LV_ART2.ListItems.count Or loc_key = 0 Then
     Exit Sub
    End If
    valor = UCase(LV_ART2.ListItems.Item(loc_key).Text)
    If Trim(UCase(textovar2.Text)) = Left(valor, Len(Trim(textovar2.Text))) And Len(Trim(textovar2.Text)) <> 0 Then
      If VAR_ACTIVAR <> 99 Then
      textovar2.Text = Trim(LV_ART2.ListItems.Item(loc_key).SubItems(1))
      Else
      textovar2.Text = Trim(LV_ART2.ListItems.Item(loc_key))
      End If
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       PUB_KEY = Val(LV_ART2.ListItems.Item(loc_key).SubItems(2))
      Else
       PUB_KEY = textovar2.Text
      End If
      LEER_ART_LLAVE
      VAR_ACTIVAR = 0
    Else
      Exit Sub
    End If
  End If
End If
    SQ_OPER = 1
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       If VAR_ACTIVAR <> 99 Then
          PUB_CODART = PUB_KEY
       Else
          PUB_CODART = Val(LV_ART2.ListItems.Item(loc_key).SubItems(2))
       End If
    Else
       PUB_CODART = Val(textovar2.Text)
    End If
    pu_codcia = LK_CODCIA
    LEER_ARM_LLAVE
    
'   If arm_llave!arm_stock <= 0 And pub_signo_arm = -1 Then
'      MsgBox "Ojo Stock no disponible " & "Existencia : " & arm_llave!arm_stock
'      If LK_USU_STOCK <> "A" Then
'         grid_fac.TextMatrix(grid_fac.Row, 3) = ""
'         grid_fac.TextMatrix(grid_fac.Row, 4) = ""
'         grid_fac.TextMatrix(grid_fac.Row, 5) = ""
'         grid_fac.TextMatrix(grid_fac.Row, 6) = ""
'         grid_fac.TextMatrix(grid_fac.Row, 7) = ""
'         textovar.Text = ""
'         Exit Sub
'      End If
'  End If
    
    

verifica_precios

GridT.CellForeColor = vbBlack
GridT.TextMatrix(GridT.Row, 0) = art_LLAVE!ART_NOMBRE

If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
   If VAR_ACTIVAR <> 99 Then
   Else
    GridT.TextMatrix(GridT.Row, 1) = LV_ART2.ListItems.Item(loc_key)
   End If
Else
  GridT.TextMatrix(GridT.Row, 1) = art_LLAVE!ART_KEY
End If
GridT.TextMatrix(GridT.Row, 5) = art_LLAVE!ART_KEY
VERIFICA_REPET2
calcula_totales22
GridT.COL = 2
PROCESA_CELDAS2

LV_ART2.Visible = False
VAR_ACTIVAR = 0
fin:
OJO:

End Sub

Private Sub textovar2_KeyUp(KeyCode As Integer, Shift As Integer)
'' CHEQUEAR AQUI...
Dim VAR
Dim WS_TIPO As Integer
Dim WS_CALIDAD As Integer
Dim ws_codcia As String

If GridT.COL = 1 And KeyCode = 46 Then
   GridT.TextMatrix(GridT.Row, 0) = ""
   GridT.TextMatrix(GridT.Row, 5) = ""
End If

If KeyCode > 35 And KeyCode < 41 Then
   Exit Sub
End If
If GridT.COL <> 1 Then Exit Sub
' puede estar pruebas
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
 If Len(textovar2.Text) = 0 Or textovar2.Text = "" Then
   VAR_ACTIVAR = 0
   LV_ART2.Visible = False
   Exit Sub
 End If
 If textovar2.Text = "*" And KeyCode = 106 Then
   VAR_ACTIVAR = 99
   Exit Sub
 ElseIf textovar2.Text = "" Then
   VAR_ACTIVAR = 0
   Exit Sub
 End If
 If VAR_ACTIVAR <> 99 Then
   Exit Sub
 End If
 If Left(textovar2.Text, 1) = "*" Then
   textovar2.Text = Mid(textovar2.Text, 2, Len(textovar2.Text))
   textovar2.SelStart = Len(textovar2.Text)
 End If

Else
 If Len(textovar2.Text) = 0 Or IsNumeric(textovar2.Text) = True Then
   LV_ART2.Visible = False
   Exit Sub
 End If
End If
If LV_ART2.Visible = False And KeyCode <> 13 Or Len(textovar2.Text) = 1 Then
'If LV_ART.Visible = False And IsNumeric(textovar.text) = False Then
    If textovar2.Text = "" Then Exit Sub
    VAR = Asc(textovar2.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    WS_TIPO = Val(Nulo_Valor0(SUT_LLAVE!SUT_art_gru))
    WS_CALIDAD = Val(Nulo_Valor0(SUT_LLAVE!SUT_CALIDAD))
    ws_codcia = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       ws_codcia = "00"
    End If
    'If LK_CODTRA = 2103 Then WS_TIPO = 4

    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      'archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO  FROM ARTI WHERE  ART_CODCIA = '" & LK_CODCIA & "' AND ART_ALTERNO BETWEEN '" & textovar.text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
      If WS_TIPO = 0 And WS_CALIDAD = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO  FROM ARTI WHERE  ART_KEY <> 0 AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & textovar.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
      If WS_CALIDAD = 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO  FROM ARTI WHERE  ART_KEY <> 0 AND ART_FAMILIA = " & WS_TIPO & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & textovar2.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
      If WS_CALIDAD <> 0 And WS_TIPO = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_KEY <> 0 AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & textovar2.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
      If WS_CALIDAD <> 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_KEY <> 0 AND ART_FAMILIA = " & WS_TIPO & " AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & textovar2.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      If WS_TIPO = 0 And WS_CALIDAD = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO  FROM ARTI WHERE  ART_KEY <> 0 AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & textovar2.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
      If WS_CALIDAD = 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO  FROM ARTI WHERE  ART_KEY <> 0 AND ART_FAMILIA = " & WS_TIPO & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & textovar2.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
      If WS_CALIDAD <> 0 And WS_TIPO = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_KEY <> 0 AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & textovar2.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
      If WS_CALIDAD <> 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_KEY <> 0 AND ART_FAMILIA = " & WS_TIPO & " AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & textovar2.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
    End If
     PROC_LISVIEW LV_ART2
     loc_key = 0
     If LV_ART2.Visible Then
      loc_key = 1
     End If
    Exit Sub
End If

'If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
' Exit Sub
'End If

Dim itmFound As ListItem    ' Variable FoundItem.
If LV_ART2.Visible Then
  Set itmFound = LV_ART2.FindItem(LTrim(textovar2.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > LV_ART2.ListItems.count Then
      LV_ART2.ListItems.Item(LV_ART2.ListItems.count).EnsureVisible
   Else
     LV_ART2.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If

End Sub


Private Sub textovar3_Change()
Dim xx As Integer
Dim WW_TOT As Currency
GridD.Text = textovar3.Text
If GridD.COL = 6 Then RECALCULA
End Sub

Private Sub textovar3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tf
Dim WSPOR As Currency
Dim fil As Integer
Dim celda_anterior
If KeyCode = 13 Then
  If GridD.COL = 3 And Val(textovar3.Text) = 0 Then Exit Sub
End If

flag_salto = 1
If GridD.COL = 3 Then
  If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
    If textovar3.Text = "" Then
    Else
     GoTo BUSCANDO
    End If
  Else
   If IsNumeric(textovar3.Text) = True Or textovar3.Text = "" Then
   Else
     GoTo BUSCANDO
  End If
End If
End If


If (KeyCode > 36 And KeyCode < 41 And textovar3.Text = "") Or KeyCode = 13 Then
Else
   Exit Sub
End If


   If KeyCode = 38 Then
      If GridD.Row = 1 Then Exit Sub
   End If
   
   If KeyCode = 37 Then
      If GridD.COL = 0 Then Exit Sub
   End If
   If KeyCode = 40 Then
      If GridD.Row = GridD.Rows - 1 Then Exit Sub
   End If
   

If GridD.COL = 3 Then
   If textovar3.Text = "" And KeyCode <> 13 Then
         GridD.TextMatrix(GridD.Row, 3) = ""
         GridD.TextMatrix(GridD.Row, 4) = ""
         GridD.TextMatrix(GridD.Row, 5) = ""
         GridD.TextMatrix(GridD.Row, 6) = ""
         calcula_totales24
      GoTo SALTO
   End If
Else
   GoTo SALTO
End If


SALTO:

GridD.TextMatrix(GridD.Row, GridD.COL) = textovar3.Text

textovar3.Visible = False
GridD.SetFocus

If KeyCode = 38 Then
   GridD.Row = GridD.Row - 1
   Exit Sub
ElseIf KeyCode = 40 Then
   GridD.Row = GridD.Row + 1
   Exit Sub
End If


If KeyCode = 37 Then
   If GridD.COL > 0 Then GridD.COL = GridD.COL - 1
   Exit Sub
ElseIf KeyCode = 39 Then
   If GridD.COL < 6 Then GridD.COL = GridD.COL + 1
   Exit Sub
End If


If KeyCode = 13 Then
   If GridD.COL = 3 Then
      GridD.COL = GridD.COL + 1
      Exit Sub
   End If
End If

GoTo fin

BUSCANDO:
If GridD.COL <> 3 Then Exit Sub

Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not LV_ART2.Visible Then
 Exit Sub
End If

If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And textovar2.Text = "" Then
  loc_key = 1
  Set LV_ART2.SelectedItem = LV_ART2.ListItems(loc_key)
'  LV_CLI.Visible = False
  LV_ART2.ListItems.Item(loc_key).Selected = True
  LV_ART2.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_ART2.ListItems.count Then loc_key = LV_ART2.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_ART2.ListItems.count Then loc_key = LV_ART2.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
'  KeyCode = 0
  LV_ART2.ListItems.Item(loc_key).Selected = True
  LV_ART2.ListItems.Item(loc_key).EnsureVisible
  textovar3.Text = Trim(LV_ART2.ListItems.Item(loc_key).Text) & " "
  DoEvents
  
  textovar3.SelStart = Len(textovar3.Text)
  DoEvents
fin:
'grid_fac.COL = 6
'grid_fac.SetFocus
OJO:

End Sub

Private Sub textovar3_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i, car
Dim itmFound As ListItem    ' Variable FoundItem.
car = Chr(KeyAscii)
car = UCase(car)
KeyAscii = Asc(car)
If KeyAscii = 27 Then
  textovar3.Text = ""
  LV_ART2.Visible = False
  Exit Sub
End If


If KeyAscii <> 13 Then pub_ojo = "A"

If GridD.COL = 6 Then
 Consistencias GridD, textovar3, KeyAscii, 2, 3
 If Len(textovar3.Text) > 13 Then
    KeyAscii = 0
    Exit Sub
 End If
End If



If GridD.COL = 3 And KeyAscii <> 13 Then
   GridD.TextMatrix(GridD.Row, 3) = ""
   GridD.TextMatrix(GridD.Row, 4) = ""
   GridD.TextMatrix(GridD.Row, 5) = ""
   GridD.TextMatrix(GridD.Row, 6) = ""
   GridD.TextMatrix(GridD.Row, 7) = ""
   calcula_totales24
End If

If KeyAscii <> 13 Then
   textovar3.ForeColor = vbRed
End If

If KeyAscii <> 13 Then
   GoTo fin
End If

flag_salto = 0
'If GridD.COL <> 3 Then Exit Sub
If GridD.TextMatrix(GridD.Row, 3) <> textovar3.Text Then Exit Sub

 If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  PUB_KEY = 0
 Else
 On Error GoTo OJO
 PUB_KEY = Val(textovar3.Text)
 On Error GoTo 0
 If Len(textovar3.Text) = 0 Then
    Exit Sub
 End If
 If IsNumeric(textovar3.Text) = False Then
   PUB_KEY = 0
 End If
End If

If PUB_KEY <> 0 Then
    SQ_OPER = 1
    PUB_KEY = textovar3.Text
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul textovar3, textovar3
       GoTo fin
    End If
    GridD.TextMatrix(GridD.Row, 7) = PUB_KEY ' art_LLAVE!ART_KEY
Else
  If VAR_ACTIVAR <> 99 And textovar3 <> "" And LK_FLAG_ORIGINAL <> "A" And LK_FLAG_ALTERNO = "A" Then
     SQ_OPER = 3
     pu_alterno = textovar3.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul textovar3, textovar3
       Exit Sub
     Else
       If PUB_TIPMOV = 20 Then
          If SUT_LLAVE!SUT_SECUENCIA = 1 Then
             If art_llave_alt!art_flag_stock <> "M" Then
                MsgBox "No es Mercaderias ... ", 48, Pub_Titulo
                textovar3.Text = ""
                Exit Sub
             End If
          ElseIf art_llave_alt!art_flag_stock <> "A" Then
                MsgBox "No es Activo Fijo ... ", 48, Pub_Titulo
                textovar3.Text = ""
                Exit Sub
        End If
        
     End If
     End If
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
      PUB_KEY = art_llave_alt!ART_KEY
      LEER_ART_LLAVE
      If PUB_TIPMOV = 75 Then
         If Val(Nulo_Valor0(art_LLAVE!ART_CODART2)) = 0 Then
            MsgBox "Articulo no esta relacionado..."
            Azul textovar3, textovar3
            Exit Sub
         End If
      End If
  Else
    If loc_key > LV_ART2.ListItems.count Or loc_key = 0 Then
     Exit Sub
    End If
    valor = UCase(LV_ART2.ListItems.Item(loc_key).Text)
    If Trim(UCase(textovar3.Text)) = Left(valor, Len(Trim(textovar3.Text))) And Len(Trim(textovar3.Text)) <> 0 Then
      If VAR_ACTIVAR <> 99 Then
      textovar3.Text = Trim(LV_ART2.ListItems.Item(loc_key).SubItems(1))
      Else
      textovar3.Text = Trim(LV_ART2.ListItems.Item(loc_key))
      End If
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       PUB_KEY = Val(LV_ART2.ListItems.Item(loc_key).SubItems(2))
      Else
       PUB_KEY = textovar3.Text
      End If
      LEER_ART_LLAVE
      VAR_ACTIVAR = 0
    Else
      Exit Sub
    End If
  End If
End If
    SQ_OPER = 1
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       If VAR_ACTIVAR <> 99 Then
          PUB_CODART = PUB_KEY
       Else
          PUB_CODART = Val(LV_ART2.ListItems.Item(loc_key).SubItems(2))
       End If
    Else
       PUB_CODART = Val(textovar3.Text)
    End If
    pu_codcia = LK_CODCIA
    LEER_ARM_LLAVE
    

verifica_precios

GridD.CellForeColor = vbBlack
GridD.TextMatrix(GridD.Row, 4) = art_LLAVE!ART_NOMBRE
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
   If VAR_ACTIVAR <> 99 Then
   Else
    GridD.TextMatrix(GridD.Row, 3) = LV_ART2.ListItems.Item(loc_key)
   End If
End If

GridD.TextMatrix(GridD.Row, 7) = art_LLAVE!ART_KEY
calcula_totales24
GridD.COL = 5
PROCESA_CELDAS3

LV_ART2.Visible = False
VAR_ACTIVAR = 0
fin:
OJO:

End Sub

Private Sub textovar3_KeyUp(KeyCode As Integer, Shift As Integer)
'' CHEQUEAR AQUI...
Dim VAR
Dim WS_TIPO As Integer
Dim WS_CALIDAD As Integer
Dim ws_codcia As String
If KeyCode = 46 Then
   GridD.CellBackColor = vbBlue
   pub_mensaje = "¿Desea Eliminar la fila seleccionada ... ?"
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   If Pub_Respuesta = vbNo Then
      GridD.CellBackColor = vbWhite
      Exit Sub
   End If
   GridD.RemoveItem (GridD.Row)
   GridD.Refresh
   RECALCULA
   GridD.SetFocus
   Exit Sub
End If
' CAMTEX
If GridD.COL = 6 And KeyCode = 116 And LK_EMP = "CAM" Then
  LOC_BRUTO = 0
  'Load Frmcalculo
  'Frmcalculo.Show 1
  textovar3.Text = LOC_BRUTO
  Exit Sub
End If




If KeyCode > 35 And KeyCode < 41 Then
   Exit Sub
End If
If GridD.COL <> 3 Then Exit Sub
' puede estar pruebas
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
 If Len(textovar3.Text) = 0 Or textovar3.Text = "" Then
   VAR_ACTIVAR = 0
   LV_ART2.Visible = False
   Exit Sub
 End If
 If textovar3.Text = "*" And KeyCode = 106 Then
   VAR_ACTIVAR = 99
   Exit Sub
 ElseIf textovar3.Text = "" Then
   VAR_ACTIVAR = 0
   Exit Sub
 End If
 If VAR_ACTIVAR <> 99 Then
   Exit Sub
 End If
 If Left(textovar3.Text, 1) = "*" Then
   textovar3.Text = Mid(textovar3.Text, 2, Len(textovar3.Text))
   textovar3.SelStart = Len(textovar3.Text)
 End If

Else
 If Len(textovar3.Text) = 0 Or IsNumeric(textovar3.Text) = True Then
   LV_ART2.Visible = False
   Exit Sub
 End If
End If
If LV_ART2.Visible = False And KeyCode <> 13 Or Len(textovar3.Text) = 1 Then
'If LV_ART.Visible = False And IsNumeric(textovar.text) = False Then
    If textovar3.Text = "" Then Exit Sub
    VAR = Asc(textovar3.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    WS_TIPO = Val(Nulo_Valors(SUT_LLAVE!SUT_POLLOS))
    WS_CALIDAD = Val(Nulo_Valor0(SUT_LLAVE!SUT_CALIDAD))
    ws_codcia = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       ws_codcia = "00"
    End If
    'If LK_CODTRA = 2103 Then WS_TIPO = 1

    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      'archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO  FROM ARTI WHERE  ART_CODCIA = '" & LK_CODCIA & "' AND ART_ALTERNO BETWEEN '" & textovar.text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
      If WS_TIPO = 0 And WS_CALIDAD = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO  FROM ARTI WHERE  ART_KEY <> 0 AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & textovar3.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
      If WS_CALIDAD = 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO  FROM ARTI WHERE  ART_KEY <> 0 AND ART_FAMILIA = " & WS_TIPO & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & textovar3.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
      If WS_CALIDAD <> 0 And WS_TIPO = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_KEY <> 0 AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & textovar3.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
      If WS_CALIDAD <> 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_KEY <> 0 AND ART_FAMILIA = " & WS_TIPO & " AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & textovar.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      If WS_TIPO = 0 And WS_CALIDAD = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO  FROM ARTI WHERE  ART_KEY <> 0 AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & textovar.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
      If WS_CALIDAD = 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO  FROM ARTI WHERE  ART_KEY <> 0 AND ART_FAMILIA = " & WS_TIPO & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & textovar3.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
      If WS_CALIDAD <> 0 And WS_TIPO = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_KEY <> 0 AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & textovar3.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
      If WS_CALIDAD <> 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_KEY <> 0 AND ART_FAMILIA = " & WS_TIPO & " AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & textovar3.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
    End If
     PROC_LISVIEW LV_ART2
     loc_key = 0
     If LV_ART2.Visible Then
      loc_key = 1
     End If
    Exit Sub
End If

'If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
' Exit Sub
'End If

Dim itmFound As ListItem    ' Variable FoundItem.
If LV_ART2.Visible Then
  Set itmFound = LV_ART2.FindItem(LTrim(textovar3.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > LV_ART2.ListItems.count Then
      LV_ART2.ListItems.Item(LV_ART2.ListItems.count).EnsureVisible
   Else
     LV_ART2.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If

End Sub

Private Sub textovarL_Change()
gridl.Text = textovarl.Text
End Sub

Private Sub textovarl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   textovarl.Text = textovar_bak
End If


If (KeyCode > 36 And KeyCode < 41) Or KeyCode = 13 Then
Else
   Exit Sub
End If


   If KeyCode = 38 Then
      If gridl.Row = 2 Then Exit Sub
   End If
   
   If KeyCode = 37 Then
      If gridl.COL = 0 Then Exit Sub
   End If
   If KeyCode = 40 Then
      If gridl.Row = gridl.Rows - 1 Then Exit Sub
   End If
   
If KeyCode = 13 Then
   If gridl.COL = 1 Or gridl.COL = 2 Then
   Else
   Exit Sub
   End If
End If

If gridl.COL = 5 Then
If Val(gridl.TextMatrix(gridl.Row, 5)) > Val(gridl.TextMatrix(gridl.Row, 25)) And Val(gridl.TextMatrix(gridl.Row, 5)) <> 0 Then
   MsgBox "Importe mayor ..."
   If LK_CODTRA = 2770 Then textovarl.Text = ""
End If
End If



SALTO:

gridl.TextMatrix(gridl.Row, gridl.COL) = textovarl.Text

textovarl.Visible = False
gridl.SetFocus



If KeyCode = 38 Then
   gridl.Row = gridl.Row - 1
ElseIf KeyCode = 40 Then
   gridl.Row = gridl.Row + 1
End If


If KeyCode = 37 Then
   gridl.COL = gridl.COL - 1
ElseIf KeyCode = 39 Or KeyCode = 13 Then
      If gridl.COL = 23 Then gridl.COL = 1
      gridl.COL = gridl.COL + 1
End If

End Sub

Private Sub textovarL_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i, car
Dim itmFound As ListItem    ' Variable FoundItem.
car = Chr(KeyAscii)
car = UCase(car)
KeyAscii = Asc(car)



If gridl.COL = 5 Then
 Consistencias_rich gridl, textovarl, KeyAscii, 2, 4
 textovarl.MaxLength = 17
Else
 textovarl.MaxLength = 0
End If





If KeyAscii <> 13 Then Exit Sub
PROCESA_CELDASL

'If i_codcli.Visible = True Then
'If Trim(gridl.TextMatrix(gridl.Row, 0)) = "" Then Exit Sub
'End If


'If gridl.Col = 18 Then
'   gridl.Col = 5
'   Exit Sub
'End If
'If gridl.Col = 23 Then
'   gridl.Col = 5
'   Exit Sub
'End If

'If gridl.Col = 19 Then
'   calcula_totales2
 '  gridl.Col = 20
 '  Exit Sub
'End If
'If gridl.Col = 20 Then
'   gridl.Col = 21
'   Exit Sub
'End If
'If gridl.Col = 21 Then
'   gridl.Col = 22
'   Exit Sub
'End If
'If gridl.Col = 22 Then
'   gridl.Col = 19
'   Exit Sub
'End If
'If LK_EMP = "PLA" Then
'If gridl.Col = 5 Then
'   calcula_totales2
'   gridl.Col = 23
'   Exit Sub
'End If
'End If
If gridl.COL = 5 Then
       gridl.COL = 6
ElseIf gridl.COL = 6 Then
       If i_fbg.Visible = True Then
          i_fbg.SetFocus
       Else
       gridl.COL = 5
       End If
End If
        
If gridl.COL = 6 Or gridl.COL = 20 Then calcula_totales2

End Sub

Private Sub textovarL_KeyUp(KeyCode As Integer, Shift As Integer)

'Dim i_tasaa
'Dim plaza_tasa1
'Dim plaza_tasa2
'i_tasaa = 0
If KeyCode = 113 And gridl.COL = 5 And LK_CODTRA <> 2728 Then
   textovarl.Text = Abs(gridl.TextMatrix(gridl.Row, 25))
   calcula_totales2
   If i_fbg.Visible = True Then i_fbg.SetFocus
   Exit Sub
End If

'Dim WS_DIASV, ws_diasa, WS_INTVEN, WS_INTADE
'Dim WK_DIASV As Integer
'Dim WK_DIASA As Integer


'If (KeyCode = 115 Or KeyCode = 119) And (gridl.Col = 5) Then
'Else
'Exit Sub
'End If'


'fila = gridl.Row

'If LK_EMP = "PLA" Then
'   If gridl.TextMatrix(gridl.Row, 26) < LK_FECHA_DIA Then
'      gridl.TextMatrix(gridl.Row, 6) = LK_FECHA_DIA
'   Else
'      gridl.TextMatrix(gridl.Row, 6) = gridl.TextMatrix(gridl.Row, 26)
'   End If
'End If

'If IsDate(gridl.TextMatrix(gridl.Row, 6)) = False Then
'   MsgBox "Fecha esta errada.. Revisar "
'   Exit Sub
'End If

'   gridl.Visible = False
'   textovarl.Visible = False
'   grid_liq.Visible = True
'   grid_liq.Width = 7500
'   grid_liq.Height = 2000
'   grid_liq.Top = gridl.Top
'   grid_liq.Left = 1080
'   grid_liq.Row = 1
'   grid_liq.RowHeight(0) = 500
'   grid_liq.RowHeight(1) = 500
'   grid_liq.Rows = 2
'   grid_liq.Cols = 13
'   grid_liq.ColWidth(0) = 0
'   grid_liq.ColWidth(1) = 1000
'   grid_liq.ColWidth(2) = 1000
'   grid_liq.ColWidth(3) = 0
'   grid_liq.ColWidth(4) = 0
'   grid_liq.ColWidth(5) = 1000
'   grid_liq.ColWidth(6) = 1000
'   grid_liq.ColWidth(7) = 1000
'   grid_liq.TextMatrix(1, 1) = 0
'   grid_liq.TextMatrix(1, 2) = 0
   
'Dim J
'   WS_DIASV = DateDiff("d", gridl.TextMatrix(gridl.Row, 26), LK_FECHA_DIA)
'   If WS_DIASV < 0 Then
'      ws_diasa = DateDiff("d", gridl.TextMatrix(gridl.Row, 26), gridl.TextMatrix(gridl.Row, 6))
'   Else
'      ws_diasa = DateDiff("d", LK_FECHA_DIA, gridl.TextMatrix(gridl.Row, 6))
'   End If
   
'   WK_DIASA = ws_diasa
'   WK_DIASV = WS_DIASV
   
'If WK_DIASV > 0 Then
'   i_tasav.Text = Val(XLL.Cells(WK_DIASV, 1)) + Val(XLL.Cells(WK_DIASV, 2))
   ' PARA PLAZA GRAU
'   If LK_CODTRA = 2728 And LK_EMP = "PLA" Then
'      PLAZA_FLAG_MANUAL = " "
'      plaza_tasa1 = InputBox("Ingrese (%)Tasa  Compensatorio : Días: " & WK_DIASV, "Tasa de Interes", "0")
'      If plaza_tasa1 = "" Then GoTo pasa_tasa
'      plaza_tasa2 = InputBox("Ingrese (%)Tasa  Moratoria     : Días: " & WK_DIASV, "Tasa de Interes", "0")
'      If plaza_tasa2 = "" Then GoTo pasa_tasa
'      If Val(i_tasav.Text) <> (Val(plaza_tasa1) + Val(plaza_tasa2)) Then
'         pub_mensaje = "Importe de Tasas No son Correctas  ...   ¿Desea Continuar de todas maneras ... ?"
'         Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
'         If Pub_Respuesta = vbNo Then
'              GoTo pasa_tasa
'              Exit Sub
'         End If
'         PLAZA_FLAG_MANUAL = "M"
'      End If
'pasa_tasa:
'   End If
'End If

'   WS_INTADE = 0
'   WS_INTVEN = 0
'
'  If KeyCode = 119 Then  ' AMORTIZACION SOLO
'         i_importe.Text = gridl.TextMatrix(fila, 25)
'         i_importe_amort.Text = gridl.TextMatrix(fila, 5)
'         WS_INTVEN = Format((i_importe.Text * Val(i_tasav.Text) / 100), "0.00")
'         WS_INTADE = (i_importe.Text - i_importe_amort.Text) * Val(i_tasaa) / 100
'         WS_INTADE = Format(WS_INTADE, "0.00")
         
'         If Val(WS_INTVEN) < 0 Then WS_INTVEN = 0
'         If Val(WS_INTADE) < 0 Then WS_INTADE = 0
       
'         WS_IMPORTE = Val(i_importe_amort.Text) + Val(WS_INTVEN) + Val(WS_INTADE)
'         GoTo saltito
'  End If
      
      
'      WS_IMPORTE = 0
      
'      WK_DIASV = WS_DIASV
'      WK_DIASA = ws_diasa
'      If WK_DIASV < 0 Then WK_DIASV = 0
'      If WK_DIASA < 0 Then WK_DIASA = 0
         
            
'      i_importe_amort.Text = (Val(gridl.TextMatrix(fila, 5)) - Val(gridl.TextMatrix(fila, 4)) * (WK_DIASV * Val(i_tasav.Text) / 100 + WK_DIASA * i_tasaa / 100)) / (1 - Val(WK_DIASA) * Val(i_tasaa) / 100)
      
'      J = 0
'      Do Until WS_IMPORTE = Val(gridl.TextMatrix(fila, 5)) Or J > 5000
'      i_importe_amort = Format(i_importe_amort.Text, "0.00")
'      WS_INTVEN = Format((Val(gridl.TextMatrix(fila, 25)) * Val(i_tasav.Text) / 100), "0.00")
'      WS_INTADE = 0 ' (gridl.TextMatrix(fila, 4) - i_importe_amort.Text) * i_tasaa * WK_DIASA / 100
'      WS_INTADE = Format(WS_INTADE, "0.00")
      
'      If Val(WS_INTVEN) < 0 Then
'         WS_INTVEN = 0
'      End If
'      If Val(WS_INTADE) < 0 Then
'         WS_INTADE = 0
'      End If
      
      
'      J = J + 1
'      WS_IMPORTE = Val(i_importe_amort.Text) + Val(WS_INTADE) + Val(WS_INTVEN)
'
'      If WS_IMPORTE < Val(gridl.TextMatrix(fila, 5)) Then
'         i_importe_amort.Text = Val(i_importe_amort.Text) + 0.01
'      ElseIf WS_IMPORTE > Val(gridl.TextMatrix(fila, 5)) Then
'         i_importe_amort.Text = Val(i_importe_amort.Text) - 0.01
'      End If
'
'      Loop
      
'      If Val(i_importe_amort.Text) < 0 Or J >= 5000 Then
'         MsgBox "Monto Insuficiente...aumentar..."
'         i_importe_amort.Text = 0
'         WS_INTADE = 0
'         WS_INTVEN = 0
'         gridl.Visible = True
'         Exit Sub
'      End If
'saltito:
'   grid_liq.TextMatrix(1, 1) = i_importe_amort.Text
'   grid_liq.TextMatrix(1, 2) = WS_INTVEN
'   grid_liq.TextMatrix(1, 3) = WS_INTADE
   
'   grid_liq.TextMatrix(1, 5) = Val(grid_liq.TextMatrix(1, 1)) + Val(grid_liq.TextMatrix(1, 2)) + Val(grid_liq.TextMatrix(1, 3))
'   grid_liq.TextMatrix(1, 6) = WS_DIASV
'   grid_liq.TextMatrix(1, 7) = ws_diasa


End Sub

Private Sub textovarl_LostFocus()
calcula_totales2
End Sub


Private Sub TRANS_Click()
TRANS.Locked = False
End Sub

Private Sub TRANS_GotFocus()
TRANS.Locked = False
'Azul TRANS, TRANS
End Sub

Private Sub Trans_KeyPress(KeyAscii As Integer)
Dim car As String
Dim SN As String
Dim Control As Object
Dim PRIMER As Integer
Dim pos1, pos2, DIF As Integer
Dim CARAC, CARAC2 As String
Dim i, J, WS_SECUENCIA As Integer
inicio:
SOLO_DECIMAL TRANS, KeyAscii

If KeyAscii = 27 Then
  Exit Sub
End If
If KeyAscii = 13 Then
  If TRANS.Text = "" Then
    TRANS.SetFocus
    Exit Sub
  End If
End If
If KeyAscii <> 13 Then
   Exit Sub
End If
If Val(TRANS.Text) = 2408 Or Val(TRANS.Text) = 2406 Or Val(TRANS.Text) = 2101 Or Val(TRANS.Text) = 2105 Or Val(TRANS.Text) = 5555 Or Val(TRANS.Text) = 2107 Or Val(TRANS.Text) = 2103 Or Val(TRANS.Text) = 2210 Or Val(TRANS.Text) = 2212 Or Val(TRANS.Text) = 2125 Or Val(TRANS.Text) = 1402 Or Val(TRANS.Text) = 2211 Then
Else
     MsgBox "Transacción NO corresponde a esté Módulo ", 48, Pub_Titulo
     Azul TRANS, TRANS
     Exit Sub
End If



If LK_CODTRA = Val(TRANS.Text) And grabar.Enabled = True Then
   If grid_fac.Visible Then grid_fac.SetFocus
   Exit Sub
End If
If Trim(TRANS.Text) = "5555" And (LK_CODUSU = "ADMIN") Then
  ACT_FORM_GRIFO
  GoTo SALIR
End If
If Trim(TRANS.Text) = "2408" Or Trim(TRANS.Text) = "2406" Or Trim(TRANS.Text) = "2101" Or Trim(TRANS.Text) = "2105" Or Trim(TRANS.Text) = "2107" Or Trim(TRANS.Text) = "2103" Or Trim(TRANS.Text) = "2210" Or Trim(TRANS.Text) = "2212" Or Trim(TRANS.Text) = "2125" Or Trim(TRANS.Text) = "1402" Or Trim(TRANS.Text) = "2211" Then
Else
   MsgBox "Seleccione otro boton ..."
   GoTo SALIR
End If


If tra_llave.EOF Then
   GoTo SALTARIN
End If


cancela_todito
pasa_cabeza
' desaparece a todos los campos actuales
nn = 2
m_ind = 0
Do Until Val(tra_llave(nn)) = 0 Or nn = 62
         m_ind = m_ind + 1
         LABELGEN(m_ind).Visible = False
         NUMERO = TABLA_TAG(tra_llave(nn))
         If TypeOf Controls(NUMERO) Is TextBox Then
            Controls(NUMERO).Text = ""
         End If
         Controls(NUMERO).Visible = False
nn = nn + 4
Loop
SALTARIN:


If Left(TRANS.Text, 4) = "2103" Then
PUB_TIPREG = 600
PUB_CODCIA = "00"
SQ_OPER = 2
LEER_TAB_LLAVE
Do Until tab_mayor.EOF
  If Trim(Nulo_Valors(tab_mayor!TAB_NOMLARGO)) <> "" Then
     Text1(tab_mayor!TAB_NUMTAB).Visible = True
     Text1(tab_mayor!TAB_NUMTAB).Text = ""
     Label1(tab_mayor!TAB_NUMTAB).Visible = True
     Label1(tab_mayor!TAB_NUMTAB).Caption = tab_mayor!TAB_NOMLARGO
  End If
  tab_mayor.MoveNext
Loop

End If




i_def.Visible = True
i_def.Clear

PUB_CODTRA = Val(Left(TRANS.Text, 4))
LK_CODTRA = PUB_CODTRA
If TRANS.Text <> "" Then
   SQ_OPER = 1
   LEER_TRA_LLAVE
   If tra_llave.EOF Then
      MsgBox "TRANSACCION NO EXISTE...", 48, Pub_Titulo
      Azul TRANS, TRANS
      GoTo SALIR
   Else
     If Nulo_Valors(tra_llave!TRA_FLAG_ACTIVO) <> "A" Then
        Azul TRANS, TRANS
        GoTo SALIR
      End If
     LK_NOMTRA = tra_llave(1)
      LK_CODTRA = tra_llave(0)
      nomtra.Caption = LK_NOMTRA
   End If
End If
SN = "N"
i = 0
Do Until SN = "S" Or i = 10
   i = i + 1
   J = 0
   Do Until SN = "S" Or J = 10
     J = J + 1
     If lk_GRUPOS(J) = tra_llave(92 + i) And lk_GRUPOS(J) <> 0 Then
         SN = "S"
     End If
   Loop
Loop

'If SN = "S" Then
'   GoTo CACHETE
'End If


   J = 1
   Do Until lk_CODTRAS(J) = "" Or SN = "Y" Or J = 10
      If Left(lk_CODTRAS(J), 4) = tra_llave(0) Then
         SN = "Y"
         Exit Do
      End If
   J = J + 1
   Loop

If SN = "N" And LK_CODUSU <> "ADMIN" Then
   MsgBox "No Tiene los Derechos Asignados"
   GoTo SALIR
End If
'i_fbg.ListIndex = 0
PS_GEN(0) = 0
GEN.Requery
PASO:

nomtra.Visible = True
CARAC = TRANS.Text
LLENA_CAMPOS
pasa_cabeza
If gridl.Visible = True Then
   gridl.Clear
   gridl.Rows = 3
   pasa_cabeza_L
End If




TRANS.Text = CARAC

NUMERO = TABLA_TAG(tra_llave(2))
WS_INDICE_RETORNO = NUMERO
SN = "S"
If SN = "S" Then
SQ_OPER = 2
PUB_CODTRA = LK_CODTRA
LEER_SUT_LLAVE
Do Until SUT_MAYOR.EOF
   If LK_EMP = "HER" Then
      If SUT_MAYOR!SUT_SECUENCIA = 25 And LK_CODTRA = 2101 Then
         If (LK_CODUSU = "SUPERVISOR" Or LK_CODUSU = "ADMIN") Then
            i_def.AddItem SUT_MAYOR!SUT_SECUENCIA & ".-" & SUT_MAYOR!SUT_DESCRIPCION
         End If
      Else
          i_def.AddItem SUT_MAYOR!SUT_SECUENCIA & ".-" & SUT_MAYOR!SUT_DESCRIPCION
      End If
   Else
     i_def.AddItem SUT_MAYOR!SUT_SECUENCIA & ".-" & SUT_MAYOR!SUT_DESCRIPCION
   End If
SUT_MAYOR.MoveNext
Loop
GoTo todo
End If


CARAC2 = "N"
pos1 = 5
pos2 = 99
Do Until CARAC2 = "S" Or pos2 = 0
   pos1 = pos1 + 1
   pos2 = InStr(pos1, lk_CODTRAS(J), ".", 1)
   If pos2 = 0 Then
      pos2 = pos1 + 3
      CARAC2 = "S"
   End If
   DIF = pos2 - pos1
   CARAC = Mid(lk_CODTRAS(J), pos1, DIF)
   PUB_SECUENCIA = Val(CARAC)
   SQ_OPER = 1
   PUB_CODTRA = LK_CODTRA
   LEER_SUT_LLAVE
   i_def.AddItem SUT_LLAVE!SUT_SECUENCIA & ".-" & SUT_LLAVE!SUT_DESCRIPCION
pos1 = pos2
Loop
   
todo:
pos1 = InStr(1, TRANS.Text, ".", 1)
pos1 = pos1 + 1
If pos1 = 1 Then
   PUB_SECUENCIA = 0
Else
   CARAC = Mid(TRANS.Text, pos1)
   PUB_SECUENCIA = Val(CARAC)
End If
If PUB_SECUENCIA = 0 Then
   i = 0
   GoTo listo
End If
SN = "N"
i = 0

Do Until SN = "S" Or i > i_def.ListCount - 1
   pos1 = InStr(1, i_def.List(i), ".", 1)
   pos1 = pos1 - 1
   CARAC = Mid(i_def.List(i), 1, pos1)

   If PUB_SECUENCIA = Val(CARAC) Then
      SN = "S"
      Exit Do
   End If
i = i + 1
Loop

If SN = "N" And LK_CODUSU <> "ADMIN" Then
   MsgBox "No Tiene los Derechos Asignados "
   GoTo SALIR
End If



listo:

If i_def.ListCount = 0 Then
   MsgBox "Falta Definir Sub TRansacciones"
   Exit Sub
End If

Frame4.Visible = False
BolFac.Caption = ""
grabar.Enabled = True
cancelar.Enabled = True
i_def.ListIndex = 0
i_def.SetFocus
SendKeys "%{UP}"
If LK_EMP = "3AA" And LK_CODTRA = 2101 Then
 i_def.ListIndex = 1
 i_DEF_KeyPress 13
 Exit Sub
End If


Exit Sub
SALIR:

i_numfac.Locked = True
i_numser.Locked = True
i_ser_lote.Locked = True
i_num_lote.Locked = True

i_cambio.Value = 0
i_cambio.Visible = False
i_numfac.BackColor = QBColor(7)
i_numser.BackColor = QBColor(7)
BolFac.Caption = ""

grabar.Enabled = False
cancelar.Enabled = False
i_def.Visible = False
nomtra.Caption = ""
i_numplan.Text = ""
BolFac.Caption = ""
i_def.Clear

End Sub

Private Sub trans_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
 LLENA_LISTRANSA LisTransa, 2
 LisTransa.Left = 980
 LisTransa.Top = 0
 LisTransa.Width = 4500
 LisTransa.Height = 3000
 LisTransa.SetFocus
End If

End Sub

Private Sub TRANS_LostFocus()
TRANS.Locked = True

End Sub

Public Function consis() As Boolean
Dim filafac As Integer

PUB_SO = "X"

If PUB_TIPMOV = 10 Or PUB_TIPMOV = 20 Then
   PUB_SO = ""
   If i_fbg.Text = "F" Or i_fbg.Text = "B" Then PUB_SO = "A"
End If


'If NJ.Visible = True Then
'   If NJ.text = "S" Or NJ.text = "N" Then
'   Else
'      MsgBox "S para confirmar ...", 48, Pub_Titulo
'      GoTo salirf
'   End If
'   If NJ.text = "S" Then PUB_SO = "A"
'   If NJ.text = "N" Then PUB_SO = " "
'End If




If LK_CODTRA = 2728 Then
   If Val(grid_liq.TextMatrix(1, 5)) = 0 Or grid_liq.Visible = False Then
        MsgBox "Faltan Datos ...", 48, Pub_Titulo
        GoTo salirf
   End If
End If

If LK_CODTRA = 2210 Then
   If grid_fac.TextMatrix(2, 0) = "" Or grid_fac.Visible = False Then
      MsgBox "Faltan Datos ...", 48, Pub_Titulo
      GoTo salirf
   End If
End If


If i_fecha_compra.Visible = True Then
   If IsDate(i_fecha_compra.Text) = False Then
      MsgBox "Fecha Errada ...", 48, Pub_Titulo
        GoTo salirf
   End If
End If


If LK_CODTRA = 2105 Then
      If (Val(gridl.TextMatrix(1, 5)) <> Val(i_neto.Text) And Val(gridl.TextMatrix(1, 5))) Or Val(i_neto.Text) = 0 Then
         MsgBox "Importe no corresponde ...", 48, Pub_Titulo
         GoTo salirf
      End If
End If

If LK_CODTRA = 2101 Then
      If Val(i_neto.Text) <> Val(grid_canje.TextMatrix(1, 3)) And Val(grid_canje.TextMatrix(1, 3)) > 0 Then
         MsgBox "Importe no corresponde ...", 48, Pub_Titulo
         GoTo salirf
      End If
End If


If i_ds.Visible = True Then
   If i_ds.Text = "S" Or i_ds.Text = "D" Then
   Else
       MsgBox "Seleccionar Moneda...", 48, Pub_Titulo
       GoTo salirf
   End If
End If


   



If pub_ojo = "A" And LK_CODTRA <> 2103 Then
   MsgBox "Antes de Grabar ... Favor de Revisar Totales ...", 48, Pub_Titulo
   calcula_totales
   GoTo salirf
End If


If grid_fac.Visible And LK_CODTRA <> 2412 And LK_CODTRA <> 2101 And LK_CODTRA <> 2210 Then
 If PUB_TIPMOV <> 0 And Val(grid_fac.TextMatrix(1, 4)) = 0 Then
      MsgBox " Falta Seleccionar datos...", 48, Pub_Titulo
      GoTo salirf
 End If
End If

If LK_CODTRA = 2103 And grid_fac.Visible = False And Option6.Visible = True Then
   MsgBox " Falta Seleccionar Resumen ..", 48, Pub_Titulo
   GoTo salirf
End If
If LK_CODTRA = 2103 And Val(i_num_lote.Text) = 0 Then
   MsgBox " Falta Numero de orden de Pedido..", 48, Pub_Titulo
   GoTo salirf
End If


If grid_fac.Visible And LK_CODTRA = 2101 Then
 If Val(grid_fac.TextMatrix(1, 7)) = 0 Then
    'llena_numfac_grifo
    If PUB_NUMFAC = 0 Then
       MsgBox "Faltan Datos.."
       GoTo salirf
    End If
    pub_mensaje = "Es Anulacion del Documento :" & BolFac.Caption & "-" & i_numser.Text & " " & i_numfac.Text
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo)
    If Pub_Respuesta = vbNo Then
       MsgBox " Falta Seleccionar datos...", 48, Pub_Titulo
       GoTo salirf
    Else
       grid_fac.TextMatrix(2, 4) = "ANULACION"
    End If
 End If
End If


If grid_fac.Visible = True And Trim(SUT_LLAVE!SUT_abreviado) = "CONS" Then
   fila = 2
   Do Until fila > 900
     If Val(grid_fac.TextMatrix(fila, 4)) > Val(grid_fac.TextMatrix(fila, 22)) Then
            FILAX = fila
            fila = 999
            Exit Do
     End If
      fila = fila + 1
      If fila = grid_fac.Rows Then fila = 998
   Loop
   If fila = 999 Then
   ' CAMTEX
'      MsgBox "Cantidad supera lo emitido  en la Guia ... ." & grid_fac.TextMatrix(FILAX, 0), 48, Pub_Titulo
'      GoTo salirf
   End If
End If
If LK_FLAG_GRIFO = "A" Then
 If pub_responsable = 0 Then
    MsgBox "Falta Responsable...", 48, Pub_Titulo
    GoTo salirf
 End If

 PS_VE2.rdoParameters(0) = LK_CODCIA
 PS_VE2.rdoParameters(1) = Val(i_codven.Text)
 PS_VE2.rdoParameters(2) = Val(i_turno.Text)
 VE2_LLAVE.Requery
 If VE2_LLAVE.EOF Then
    MsgBox "Verifique Responsable..."
    GoTo salirf
 Else
    If VE2_LLAVE!VE2_CODRES <> pub_responsable Then
       pub_mensaje = "Ha Cambiado el Responsable ...Desea Continuar ?"
       Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo)
       If Pub_Respuesta = vbNo Then
          GoTo salirf
       Else
          VE2_LLAVE.Edit
          VE2_LLAVE!VE2_CODRES = pub_responsable
          VE2_LLAVE.Update
       End If
    End If
 End If
End If
If PUB_TIPMOV <> 0 And Val(i_numfac.Text) = 0 And LK_CODTRA <> 2103 And LK_CODTRA <> 2211 Then
'QUITE CAMTEX
'If PUB_TIPMOV <> 0 And Val(i_numfac.Text) = 0 And LK_CODTRA <> 2103 Then
   MsgBox "Falta Numero de Doc...."
   GoTo salirf
End If

consis = True

Exit Function
salirf:
consis = False

End Function

Public Sub avanza_grifo()
Dim indice
If tab_avanza(NUMERO) = 0 Then
   grabar.SetFocus
   GoTo fin
End If
If tab_avanza(NUMERO) = 100 Then
      grid_fac.Row = 2
      grid_fac.COL = 1
      If grid_fac.Visible Then grid_fac.SetFocus
      GoTo fin
End If
If tab_avanza(NUMERO) = 111 Then
      GridT.Row = 2
      GridT.COL = 1
      If GridT.Visible Then GridT.SetFocus
      GoTo fin
End If

If LK_CODTRA = 2103 Then
If indice = -2 Then
    If Text1(1).Visible = True Then Text1(1).SetFocus
    GoTo fin
End If
End If


indice = tab_avanza(NUMERO)
If indice < 0 Then GoTo fin
NUMERO = TABLA_TAG(indice)

If FORM_GRIFO.Controls(TABLA_TAG(indice)).Enabled = False Then
   GoTo fin
End If

If FORM_GRIFO.Controls(TABLA_TAG(indice)).Name = "Frame3" Then
   Frame3.Visible = True
   Frame3.Enabled = True
   Option1.Enabled = True
   Option1.Visible = True
   Option1.SetFocus
   Option1_Click
Else
'   MsgBox FORM_GRIFO.Controls(TABLA_TAG(indice)).Name

   FORM_GRIFO.Controls(TABLA_TAG(indice)).SetFocus
End If

If TypeOf FORM_GRIFO.Controls(TABLA_TAG(indice)) Is ComboBox Then
   SendKeys "%{DOWN}"
End If
fin:

End Sub

Private Sub Ver_Click()

End Sub

Public Function ANEXO_1() As Boolean
ANEXO_1 = True
If PUB_CP = "C" Or PUB_CP = "P" Then
Else
  ANEXO_1 = False
  Exit Function
End If

SQ_OPER = 1
If PUB_CODCLIE = 0 Then
   pub_mensaje_err = "Cliente Invalido ..."
   ANEXO_1 = False
   Exit Function
End If
pu_codcia = LK_CODCIA
pu_cp = PUB_CP
pu_codclie = PUB_CODCLIE
LEER_CLI_LLAVE
If cli_llave.EOF = True Then
   pub_mensaje_err = "Cliente/Proveedor no existe...."
   ANEXO_1 = False
   Exit Function
End If
i_nomCLI.Caption = Trim(cli_llave!cli_nombre) & " - " & Trim(cli_llave!cli_ruc_esposo)

If LK_CODTRA = 1111 Or LK_CODTRA = 1122 Then Exit Function

If PUB_TIPMOV = 10 And PUB_FBG = "B" Then
   If Val(i_neto.Text) >= Val(GEN!GEN_UIT) And Trim(cli_llave!cli_ruc_esposo) = "" Then
      pub_mensaje_err = "Falta DNI para la Boleta ...por pasar de la 1/2 UIT que es = S/. " & GEN!GEN_UIT
        ANEXO_1 = False
        Exit Function
   End If
End If


   If PUB_TIPMOV <> 0 And LK_CODTRA <> 2748 And LK_CODTRA <> 2412 And LK_CODTRA <> 2410 And LK_CODTRA <> 1111 And LK_CODTRA <> 1122 Then
      If Trim(i_moneda.Text) = "" Then
         pub_mensaje_err = "Falta LA Moneda ..."
           ANEXO_1 = False
           Exit Function
      End If
   Else
    If PUB_TIPMOV <> 0 And LK_CODTRA <> 2748 And LK_CODTRA <> 1111 And LK_CODTRA <> 1122 Then
      If Trim(lmoneda.Caption) = "" Then
         pub_mensaje_err = "Falta LA Moneda ..."
           ANEXO_1 = False
           Exit Function
      End If
    End If
   End If
     

If PUB_TIPMOV <> 0 And LK_CODTRA <> 1111 And LK_CODTRA <> 1122 Then
If WS_MONEDA_CLI = "S" Or WS_MONEDA_CLI = "D" Then
Else
   pub_mensaje_err = "Falta LA Moneda ..."
   ANEXO_1 = False
   Exit Function
End If
End If

If (PUB_TIPMOV = 10 And PUB_FBG = "F") And LK_CODTRA <> 1111 Then
   If Trim(cli_llave!cli_ruc_esposo) = "" Then
      pub_mensaje_err = "Falta el N. Ruc ..."
      ANEXO_1 = False
      Exit Function
  End If
End If
If LK_CODTRA <> 1122 And LK_CODTRA <> 2412 And LK_CODTRA <> 2410 Then
If (Nulo_Valors(cli_llave!CLI_TIPO_BLOQ1) = "1" Or Nulo_Valors(cli_llave!CLI_TIPO_BLOQ2) = "1" Or Nulo_Valors(cli_llave!CLI_TIPO_BLOQ3) = "1" Or Nulo_Valors(cli_llave!CLI_TIPO_BLOQ4) = "1") And pub_signo_car = 1 And LK_CODTRA <> 1111 And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 0 Then
     pub_mensaje_err = "Credito Bloqueado ..."
     ANEXO_1 = False
     Exit Function
End If
End If


End Function
Public Function ANEXO_2() As Boolean
Dim i
ANEXO_2 = True

If PUB_TIPMOV <> 180 Then Exit Function
fila = 2
i = 0
Do Until fila = GridT.Rows - 1
   If Trim(GridT.TextMatrix(fila, 0)) = "" Then Exit Do
   If Trim(GridT.TextMatrix(fila, 7)) = "0" Then GoTo estabien
    PUB_PEDFAC = Val(GridT.TextMatrix(fila, 7))
    pu_codcia = LK_CODCIA
    PUB_PEDSER = 0
    PUB_TIPMOV = 177
    SQ_OPER = 1
    LEER_PED_LLAVE
    PUB_TIPMOV = 180
    If ped_llave.EOF = True Then
       pub_mensaje_err = "El Nro. de Pedido no Existe"
       ANEXO_2 = False
       Exit Function
    End If
    Do Until ped_llave.EOF
        If Val(ped_llave!PED_CODART) = Val(GridT.TextMatrix(fila, 1)) Then
          GoTo estabien
        End If
     ped_llave.MoveNext
    Loop
       pub_mensaje_err = "El Nro. de Pedido, no Corresponde al Articulo " & Chr(13) & Trim(GridT.TextMatrix(fila, 0)) & " - Nro O/P : " & Trim(GridT.TextMatrix(fila, 7))
       ANEXO_2 = False
       Exit Function
estabien:
fila = fila + 1
Loop

End Function

