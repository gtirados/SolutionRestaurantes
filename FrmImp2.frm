VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmImp2 
   Caption         =   "Emitir Reportes"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11550
   ControlBox      =   0   'False
   Icon            =   "FrmImp2.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   11550
   WindowState     =   2  'Maximized
   Begin ComctlLib.ListView ListView2 
      Height          =   375
      Left            =   6000
      TabIndex        =   69
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   128
      BackColor       =   14737632
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame fracodclie 
      Caption         =   "Proveedor"
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   150
      TabIndex        =   88
      Top             =   4500
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txt_cli 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   90
         MaxLength       =   8
         TabIndex        =   89
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblCliente 
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
         Left            =   75
         TabIndex        =   90
         Top             =   540
         Width           =   2565
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fradescto 
      Caption         =   "Condición  Descto.                                  Tipo de Descto:                  "
      Height          =   2055
      Left            =   4380
      TabIndex        =   85
      Top             =   4035
      Visible         =   0   'False
      Width           =   5175
      Begin VB.ListBox listat 
         Height          =   1635
         Left            =   2640
         Style           =   1  'Checkbox
         TabIndex        =   87
         Top             =   255
         Width           =   2175
      End
      Begin VB.ListBox listac 
         Height          =   1620
         Left            =   240
         TabIndex        =   86
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame frazonas 
      Height          =   3015
      Left            =   3225
      TabIndex        =   32
      Top             =   3135
      Visible         =   0   'False
      Width           =   6735
      Begin VB.OptionButton opzonas 
         Caption         =   "Zonas"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton opzonas 
         Caption         =   "Provincia"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   2175
      End
      Begin VB.OptionButton opzonas 
         Caption         =   "Distrito"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
      Begin VB.ListBox zonas 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2220
         Left            =   3480
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblzonas 
         AutoSize        =   -1  'True
         Caption         =   "Zonas :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2040
         TabIndex        =   33
         Top             =   600
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   5295
      Begin VB.Label lblformulas 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   840
         TabIndex        =   82
         Top             =   510
         Width           =   4215
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Formula:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   135
         TabIndex        =   81
         Top             =   495
         Width           =   600
      End
      Begin VB.Label LBLRUTA 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   840
         TabIndex        =   80
         Top             =   750
         Width           =   4350
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Archivo: "
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   79
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lblreporte 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   5100
      End
   End
   Begin VB.Frame fcontab 
      Height          =   1695
      Left            =   0
      TabIndex        =   42
      Top             =   1080
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CheckBox chenivel 
         Caption         =   "Nivel 6"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   49
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox chenivel 
         Caption         =   "Nivel 5"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   48
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chenivel 
         Caption         =   "Nivel 4"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   47
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chenivel 
         Caption         =   "Nivel 3"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox chenivel 
         Caption         =   "Nivel 2"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chenivel 
         Caption         =   "Nivel 1"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblcontab 
         Caption         =   "Seleccione los Niveles para impresión"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame FRASTOCK 
      Height          =   3015
      Left            =   3240
      TabIndex        =   25
      Top             =   3120
      Visible         =   0   'False
      Width           =   6735
      Begin VB.ComboBox CmbCalidad 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   600
         Width           =   2055
      End
      Begin VB.ListBox subfami 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   1155
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   1800
         Width           =   3735
      End
      Begin VB.ListBox fami 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   1155
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblcalidad 
         AutoSize        =   -1  'True
         Caption         =   "Calidad"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4320
         TabIndex        =   28
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Familia :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sub Familia :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Consolidar Compañias :"
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
      Height          =   975
      Left            =   5400
      TabIndex        =   77
      Top             =   0
      Width           =   4575
      Begin VB.CheckBox checia 
         Caption         =   "Consolidado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Frame fracli 
      Height          =   975
      Left            =   3750
      TabIndex        =   34
      Top             =   1065
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtDias2 
         Height          =   285
         Index           =   1
         Left            =   2760
         MaxLength       =   3
         TabIndex        =   38
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtDias2 
         Height          =   285
         Index           =   0
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   37
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtDias1 
         Height          =   285
         Index           =   1
         Left            =   2760
         MaxLength       =   3
         TabIndex        =   36
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtDias1 
         Height          =   285
         Index           =   0
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   35
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblcli 
         Caption         =   "2.- "
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   41
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblcli 
         Caption         =   "1.- "
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   40
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblcli 
         Caption         =   "Ingrese 2 Rangos :"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame frapasa 
      BackColor       =   &H008B4914&
      Caption         =   "Contabilidad"
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   10140
      TabIndex        =   70
      Top             =   15
      Visible         =   0   'False
      Width           =   1305
      Begin VB.CheckBox chepasa 
         BackColor       =   &H008B4914&
         Caption         =   "<--Marcar"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblpasa 
         BackColor       =   &H008B4914&
         Caption         =   "Pasar la Información al Periodo Contable"
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
         Height          =   855
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label cop_fecha1 
         BackColor       =   &H008B4914&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label cop_fecha2 
         BackColor       =   &H008B4914&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblpasa2 
         BackColor       =   &H008B4914&
         Caption         =   "al"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   72
         Top             =   1440
         Width           =   255
      End
   End
   Begin VB.Data transfer 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "FACART"
      Top             =   7080
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.CommandButton Chequeo 
      Caption         =   "chequeo FAR,ALL,CAR,CAA"
      Height          =   255
      Left            =   10320
      TabIndex        =   68
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar ProgBar 
      Height          =   195
      Left            =   0
      TabIndex        =   18
      Top             =   6720
      Visible         =   0   'False
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones  : "
      ForeColor       =   &H00000080&
      Height          =   3615
      Left            =   15
      TabIndex        =   61
      Top             =   2190
      Visible         =   0   'False
      Width           =   3135
      Begin VB.ListBox txttp 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   960
         Left            =   840
         Style           =   1  'Checkbox
         TabIndex        =   67
         Top             =   600
         Width           =   1452
      End
      Begin VB.TextBox max 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   288
         Left            =   840
         TabIndex        =   64
         Text            =   "65"
         Top             =   1800
         Width           =   972
      End
      Begin VB.TextBox fbtxt 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtserie 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   288
         Left            =   840
         TabIndex        =   62
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblmoneda1 
         Caption         =   "Lineas X pag."
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   60
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblmoneda1 
         Caption         =   "Tipo de Docu."
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   66
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblmoneda1 
         Caption         =   "Serie:"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fzonas 
      Caption         =   "Zonas"
      ForeColor       =   &H00808000&
      Height          =   2535
      Left            =   6600
      TabIndex        =   58
      Top             =   2400
      Visible         =   0   'False
      Width           =   3015
      Begin VB.ListBox solozonas 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   2010
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   59
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CheckBox chestands 
      Caption         =   "Multiplicar Monto por Nº de Stands"
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   3240
      TabIndex        =   57
      Top             =   3600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   3240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CheckBox cheasiento 
      Caption         =   "Pasar a Contabilidad"
      Height          =   255
      Left            =   0
      TabIndex        =   55
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame fradeudas 
      Height          =   1215
      Left            =   6600
      TabIndex        =   31
      Top             =   975
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CheckBox Check1 
         Caption         =   "Fechas de Nuevo Vcto."
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   54
         Top             =   120
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox Cheop2 
         Caption         =   "Deudas por Cobrar Vencidas"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Cheop1 
         Caption         =   "Deudas por Cobrar del Dia"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Cheop3 
         Caption         =   "Deudas por Cobrar por Vencer"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Value           =   1  'Checked
         Width           =   2535
      End
   End
   Begin MSMask.MaskEdBox txtfecha 
      Height          =   285
      Left            =   1680
      TabIndex        =   53
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   14737632
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.ComboBox cmdMoneda 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   315
      ItemData        =   "FrmImp2.frx":0442
      Left            =   120
      List            =   "FrmImp2.frx":044C
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Chesup 
      Caption         =   "Suprimir las Columnas con 0"
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   1560
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox vendmulti 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2490
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CheckBox chestock 
      Caption         =   "Incluir Valorado"
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox listVen 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2460
      Left            =   0
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox PROV 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   2310
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton pantalla 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Re&portar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   10230
      Picture         =   "FrmImp2.frx":0464
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Width           =   1185
   End
   Begin VB.CommandButton cerrar 
      BackColor       =   &H00FFFFFF&
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
      Height          =   795
      Left            =   10230
      Picture         =   "FrmImp2.frx":12DE
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4320
      Width           =   1185
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   15
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin MSMask.MaskEdBox txtCampo2 
      Height          =   285
      Left            =   4680
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   14737632
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdultima 
      Caption         =   "Ultima Edicíon."
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
      Left            =   3720
      TabIndex        =   30
      Top             =   1920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox txtCampo1 
      Height          =   285
      Left            =   3240
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   14737632
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Frame fracal 
      Caption         =   "Calidad del Producto :"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   0
      TabIndex        =   83
      Top             =   5760
      Visible         =   0   'False
      Width           =   2535
      Begin VB.ListBox listacal 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   510
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   84
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Label lblfecha 
      AutoSize        =   -1  'True
      Caption         =   "Fecha :"
      Height          =   195
      Left            =   1800
      TabIndex        =   52
      Top             =   1560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblmoneda 
      Caption         =   "Moneda :"
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblcampo2 
      AutoSize        =   -1  'True
      Caption         =   "Campo1"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4800
      TabIndex        =   24
      Top             =   2520
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblcampo1 
      AutoSize        =   -1  'True
      Caption         =   "Campo1"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3240
      TabIndex        =   23
      Top             =   2520
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblstock 
      Caption         =   "Proveedores"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblProceso 
      Alignment       =   2  'Center
      BackColor       =   &H008B4914&
      Caption         =   "Procesando ..."
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10155
      TabIndex        =   19
      Top             =   5280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4914&
      BorderStyle     =   1  'Fixed Single
      Height          =   7095
      Index           =   5
      Left            =   10095
      TabIndex        =   76
      Top             =   -120
      Width           =   1815
   End
End
Attribute VB_Name = "FrmImp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim loc_cp As String * 1
Dim loc_key As Integer
Dim xl As Object
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset
Dim PS_REP03 As rdoQuery
Dim llave_rep03 As rdoResultset
Dim PS_REP04 As rdoQuery
Dim llave_rep04 As rdoResultset
Dim PS_FACART As rdoQuery
Dim llave_FACART As rdoResultset

Dim PSTEMP_MAYOR As rdoQuery
Dim temp_mayor As rdoResultset



Dim wranF, wran1, wran2, WPAS

Dim c1 As Integer
Dim f1 As Integer
Dim xcuenta As Integer
Dim i As Integer
Dim Mensaje, Titulo, valorpred As String
Dim Wfile  As String
Dim WFORM  As String


Private Sub cerrar_Click()
Unload FrmImp2
End Sub

Private Sub cmdMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If txtFecha.Visible = True Then txtFecha.SetFocus
End If
End Sub

Private Sub cmdultima_Click()
Dim VAR
Dim nom
If Trim((listven.Text)) = "" Then Exit Sub
lblproceso.Visible = True
DoEvents
nom = Trim(Left(listven.Text, 3))
VAR = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\OFFICE\PLANVEN" & nom & ".XLS"

If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
End If
On Error GoTo no_existe
xl.Workbooks.Open VAR, 0, False, 4
xl.APPLICATION.Visible = True
On Error GoTo 0
Set xl = Nothing
lblproceso.Visible = False
Exit Sub
no_existe:
If Err.Number = 1004 Then MsgBox "Planilla no emitida ó ya Procesada.", 48, Pub_Titulo
Set xl = Nothing
lblproceso.Visible = False
Exit Sub

End Sub

Private Sub Chequeo_Click()
If LK_CODUSU <> "ADMIN" Then Exit Sub
Dim LLEVA_SALDO As Currency
Dim WNUM As Currency
Dim wser As Currency
Dim WNUMOPER
Dim wcodclie
Dim wfecha
Dim WTIPMOV
Dim WNUMDOC As Currency
Dim WSERDOC As Currency
If val(InputBox("1 Cheque de Saldo de Mercaderia .. 2 Chequeo de Tablas", "")) = 1 Then
GoTo chesal
ElseIf val(InputBox("1 Cheque de Saldo de Mercaderia .. 2 Chequeo de Tablas", "")) <> 2 Then
 Exit Sub
End If

pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCLIE = 334 AND FAR_CODCIA = ? AND (FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?) AND FAR_TIPMOV = 20 AND FAR_ESTADO <>'E' ORDER BY FAR_FECHA,FAR_NUMSER, FAR_NUMFAC "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = LK_FECHA_DIA
PS_REP01(2) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_TIPMOV = ? AND ALL_FECHA_DIA = ? AND ALL_CODCLIE = ? AND ALL_NUMOPER = ? AND ALL_NUMFAC = ?  AND ALL_FLAG_EXT <> 'E' "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = LK_FECHA_DIA
PS_REP02(1) = 0
PS_REP02(2) = LK_FECHA_DIA
PS_REP02(3) = 0
PS_REP02(4) = 0
PS_REP02(5) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM CARTERA WHERE CAR_IMPORTE <> 0  AND CAR_CODCIA = ? AND CAR_TIPMOV = ? AND CAR_FECHA_INGR= ? AND CAR_CODCLIE = ? AND CAR_NUMOPER = ? AND CAR_NUMFAC = ? "
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
PS_REP03(0) = LK_FECHA_DIA
PS_REP03(1) = 0
PS_REP03(2) = LK_FECHA_DIA
PS_REP03(3) = 0
PS_REP03(4) = 0
PS_REP03(5) = 0
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM CARACU WHERE CAA_CODCIA = ? AND CAA_TIPMOV = ? AND CAA_FECHA= ? AND CAA_CODCLIE = ? AND CAA_NUM_OPER = ? AND CAA_SERDOC = ? AND CAA_NUMDOC = ? AND CAA_ESTADO <> 'E' AND CAA_SIGNO_CAR = 1 "
Set PS_REP04 = CN.CreateQuery("", pub_cadena)
PS_REP04(0) = LK_FECHA_DIA
PS_REP04(1) = 0
PS_REP04(2) = LK_FECHA_DIA
PS_REP04(3) = 0
PS_REP04(4) = 0
PS_REP04(5) = 0
PS_REP04(6) = 0
Set llave_rep04 = PS_REP04.OpenResultset(rdOpenKeyset, rdConcurValues)


PS_REP01(0) = LK_CODCIA
PS_REP01(1) = CDate("2001/01/01")
PS_REP01(2) = CDate("2001/05/31")
llave_rep01.Requery
WNUM = llave_rep01!far_numfac
wser = llave_rep01!far_numser
WNUMOPER = llave_rep01!FAR_NUMOPER
wcodclie = llave_rep01!far_codclie
wfecha = llave_rep01!FAR_fecha
WTIPMOV = llave_rep01!FAR_TIPMOV

Do Until llave_rep01.EOF
If WNUM <> llave_rep01!far_numfac Then
    PS_REP02(0) = LK_CODCIA
    PS_REP02(1) = WTIPMOV
    PS_REP02(2) = wfecha
    PS_REP02(3) = wcodclie
    PS_REP02(4) = WNUMOPER
    PS_REP02(5) = WNUM
    llave_rep02.Requery
    If llave_rep02.EOF Then
      Stop
    Else
      If llave_rep02.RowCount > 1 Then
        MsgBox "mAS DE UNO "
      End If
      llave_rep02.Edit
      llave_rep02!ALL_FLAG_SO = "O"
      llave_rep02.Update
    End If
    PS_REP03(0) = LK_CODCIA
    PS_REP03(1) = WTIPMOV
    PS_REP03(2) = wfecha
    PS_REP03(3) = wcodclie
    PS_REP03(4) = WNUMOPER
    PS_REP03(5) = WNUM
    llave_rep03.Requery
    If llave_rep03.EOF Then
      Stop
    Else
      If llave_rep03.RowCount > 1 Then
        MsgBox "mAS DE UNO "
      End If
      llave_rep03.Edit
      llave_rep03!CAR_FLAG_SO = "O"
      llave_rep03.Update
    End If
    PS_REP04(0) = LK_CODCIA
    PS_REP04(1) = WTIPMOV
    PS_REP04(2) = wfecha
    PS_REP04(3) = wcodclie
    PS_REP04(4) = WNUMOPER
    If Not llave_rep03.EOF Then
    PS_REP04(5) = llave_rep03!car_SERDOC
    PS_REP04(6) = llave_rep03!car_NUMDOC
    llave_rep04.Requery
    End If
    If llave_rep04.EOF Then
      Stop
    Else
      If llave_rep04.RowCount > 1 Then
        MsgBox "mAS DE UNO "
      End If
      llave_rep04.Edit
      llave_rep04!caa_FLAG_SO = "O"
      llave_rep04.Update
    End If
    
    WNUM = llave_rep01!far_numfac
    wser = llave_rep01!far_numser
    WNUMOPER = llave_rep01!FAR_NUMOPER
    wcodclie = llave_rep01!far_codclie
    wfecha = llave_rep01!FAR_fecha
    WTIPMOV = llave_rep01!FAR_TIPMOV
    
End If
llave_rep01.Edit
llave_rep01!FAR_FLAG_SO = "O"
llave_rep01.Update
llave_rep01.MoveNext
Loop
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = WTIPMOV
PS_REP02(2) = wfecha
PS_REP02(3) = wcodclie
PS_REP02(4) = WNUMOPER
PS_REP02(5) = WNUM
llave_rep02.Requery
If llave_rep02.EOF Then
  Stop
Else
  If llave_rep02.RowCount > 1 Then
    MsgBox "mAS DE UNO "
  End If
  llave_rep02.Edit
  llave_rep02!ALL_FLAG_SO = "O"
  llave_rep02.Update
End If

PS_REP03(0) = LK_CODCIA
PS_REP03(1) = WTIPMOV
PS_REP03(2) = wfecha
PS_REP03(3) = wcodclie
PS_REP03(4) = WNUMOPER
PS_REP03(5) = WNUM
llave_rep03.Requery
If llave_rep03.EOF Then
  Stop
Else
  If llave_rep03.RowCount > 1 Then
    MsgBox "mAS DE UNO "
  End If
  llave_rep03.Edit
  llave_rep03!CAR_FLAG_SO = "O"
  llave_rep03.Update
End If
PS_REP04(0) = LK_CODCIA
 PS_REP04(1) = WTIPMOV
 PS_REP04(2) = wfecha
 PS_REP04(3) = wcodclie
 PS_REP04(4) = WNUMOPER
 PS_REP04(5) = llave_rep03!car_SERDOC
 PS_REP04(6) = llave_rep03!car_NUMDOC
 llave_rep04.Requery
 If llave_rep04.EOF Then
   Stop
 Else
   If llave_rep04.RowCount > 1 Then
     MsgBox "mAS DE UNO "
   End If
   llave_rep04.Edit
   llave_rep04!caa_FLAG_SO = "O"
   llave_rep04.Update
 End If

MsgBox "TERMINADO"

Exit Sub
chesal:

pub_cadena = "SELECT art_key FROM arti WHERE art_CODCIA = ? order by art_key"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP02(0) = LK_CODCIA



'pub_cadena = "SELECT far_fecha_compra, far_stock, far_cantidad, far_signo_arm FROM FACART WHERE FAR_CODCIA = ? AND (FAR_FECHA >= ? AND FAR_FECHA <= ?) AND FAR_CODART = ?  AND FAR_ESTADO <> 'E' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC , FAR_NUMOPER2"
pub_cadena = "SELECT FAR_DESCTO,FAR_FLETE,FAR_EQUIV,FAR_COSPRO_SUP,FAR_COSPRO_ANT, FAR_TIPMOV, FAR_CODCIA, FAR_PRECIO, FAR_PRECIO_NETO,FAR_COSPRO, FAR_SUBTRA, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_FECHA_COMPRA, FAR_CANTIDAD, FAR_SIGNO_ARM, FAR_COSPRO, FAR_CODART , FAR_TIPO_CAMBIO, FAR_MONEDA, FAR_STOCK  FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND FAR_CODART = ?  and far_estado <>'E' AND FAR_CODART <> 0 ORDER BY FAR_CODART, FAR_FECHA_COMPRA,FAR_SIGNO_ARM DESC, FAR_NUMOPER2 "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = LK_FECHA_DIA
PS_REP01(2) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = CDate("2000/07/01")
PS_REP01(2) = CDate("2001/07/31")
PS_REP01(3) = 0
llave_rep02.Requery
ProgBar.Min = 0
ProgBar.max = llave_rep02.RowCount
ProgBar.Value = 0
ProgBar.Visible = True
Do Until llave_rep02.EOF
 DoEvents
 Chequeo.Caption = str(llave_rep02.AbsolutePosition) & " / " & str(llave_rep02.RowCount)
 ProgBar.Value = llave_rep02.AbsolutePosition
 DoEvents
 PS_REP01(3) = llave_rep02!ART_KEY
 llave_rep01.Requery
 LLEVA_SALDO = 0
 Do Until llave_rep01.EOF
 LLEVA_SALDO = LLEVA_SALDO + (val(llave_rep01!far_cantidad) * val(llave_rep01!far_signo_arm))
 If llave_rep01!FAR_TIPMOV = 3 Or llave_rep01!FAR_TIPMOV = 10 Or llave_rep01!FAR_TIPMOV = 100 Or llave_rep01!FAR_TIPMOV = 5 Or llave_rep01!FAR_TIPMOV = 93 Or llave_rep01!FAR_TIPMOV = 6 Or llave_rep01!FAR_TIPMOV = 101 Or llave_rep01!FAR_TIPMOV = 97 Or llave_rep01!FAR_TIPMOV = 20 Then
 Else
 MsgBox "NO TIENE TIPMOV"
 End If
 If val(llave_rep01!FAR_STOCK) <> LLEVA_SALDO Then
   MsgBox "Recostear Nuevamente este Articulo " & llave_rep02!ART_KEY & "  Fecha: " & llave_rep01!FAR_fecha_compra, 48, ""
   GoTo SALE
  Stop
  llave_rep01.Edit
  llave_rep01!FAR_STOCK = LLEVA_SALDO
  llave_rep01.Update
 Else
 
 End If
 
 
 
  
 llave_rep01.MoveNext
 Loop
SALE:
 llave_rep02.MoveNext
Loop
ProgBar.Visible = False
MsgBox "proceso terminado"
Exit Sub

CHEBANCO:

pub_cadena = "SELECT ALL_SIGNO_CCM , ALL_FECHA_DIA , ALL_CODCIA, ALL_NUMOPER,ALL_CHENUM , ALL_CODBAN FROM ALLOG WHERE ALL_FLAG_EXT <> 'E' AND (ALL_CODTRA = 2748 OR ALL_CODTRA = 2735 )"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)
llave_rep02.Requery


pub_cadena = "SELECT * FROM CHEQUES WHERE CHE_CODCIA = ? AND CHE_FECHA = ? AND CHE_CHENUM = ? AND CHE_CODBAN = ? AND CHE_NUMOPER = ? "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = LK_FECHA_DIA
PS_REP01(2) = 0
PS_REP01(3) = 0
PS_REP01(4) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = 0
PS_REP01(2) = 0
PS_REP01(3) = 0

ProgBar.Min = 0
ProgBar.max = llave_rep02.RowCount
ProgBar.Value = 0
ProgBar.Visible = True
Do Until llave_rep02.EOF
 DoEvents
 Chequeo.Caption = str(llave_rep02.AbsolutePosition) & " / " & str(llave_rep02.RowCount)
 ProgBar.Value = llave_rep02.AbsolutePosition
 DoEvents
 PS_REP01(0) = "01"
 PS_REP01(1) = llave_rep02!ALL_FECHA_DIA
 PS_REP01(2) = llave_rep02!all_chenum
 PS_REP01(3) = llave_rep02!all_codban
 PS_REP01(4) = llave_rep02!ALL_NUMOPER
 llave_rep01.Requery
 If llave_rep02!all_CODCIA = "01" Or llave_rep02!all_CODCIA = "02" Then
 If llave_rep02!ALL_SIGNO_CCM <> 0 Then
   If llave_rep01.EOF Then
   Stop
     GoTo SALE1
   End If
   If llave_rep01.RowCount > 1 Then Stop
   llave_rep01.Edit
   'llave_rep01!CHE_CODCIA
   llave_rep01!CHE_CODCIA2 = llave_rep02!all_CODCIA
   llave_rep01.Update
 End If
 End If
SALE1:
 llave_rep02.MoveNext
Loop
ProgBar.Visible = False
MsgBox "proceso terminado"

Exit Sub

CHECARTERA:

pub_cadena = "SELECT * FROM CARTERA WHERE CAR_FECHA_SUNAT >= ? AND CAR_CP = 'P' AND (CAR_CODCIA = '02' ) ORDER BY CAR_FECHA_SUNAT "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = LK_FECHA_DIA
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP02(0) = "01/06/2001"
llave_rep02.Requery

ProgBar.Min = 0
ProgBar.max = llave_rep02.RowCount
ProgBar.Value = 0
ProgBar.Visible = True
Do Until llave_rep02.EOF
 DoEvents
 Chequeo.Caption = str(llave_rep02.AbsolutePosition) & " / " & str(llave_rep02.RowCount)
 ProgBar.Value = llave_rep02.AbsolutePosition
 DoEvents
SQ_OPER = 1
pu_cp = llave_rep02!CAR_cp
pu_codclie = llave_rep02!CAR_CODCLIE
pu_codcia = llave_rep02!car_codcia
PUB_TIPDOC = llave_rep02!car_TIPDOC
PUB_FECHA = llave_rep02!CAR_FECHA_INGR
PUB_NUM_OPER = llave_rep02!CAR_NUMOPER
LEER_CAA_LLAVE
If caa_histo.RowCount > 1 Then Stop
If caa_histo.EOF Then Stop
Do Until caa_histo.EOF
caa_histo.Edit
caa_histo!caa_FLAG_SO = "K"
caa_histo.Update
caa_histo.MoveNext
Loop

SALE8:
 llave_rep02.MoveNext
Loop
ProgBar.Visible = False
MsgBox "proceso terminado"

Exit Sub


CHECARTERA66:
pub_cadena = "SELECT * FROM FACART WHERE FAR_TIPMOV = 99 AND FAR_CODCIA = '02' AND FAR_ESTADO <> 'E' AND FAR_FECHA> '2001/06/25' ORDER BY FAR_FECHA "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ? AND ALL_NUMOPER = ? AND ALL_CODCLIE = ? AND ALL_FLAG_EXT <> 'E' "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = LK_FECHA_DIA
PS_REP02(2) = 0
PS_REP02(3) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)
llave_rep01.Requery
ProgBar.Min = 0
ProgBar.max = llave_rep01.RowCount
ProgBar.Value = 0
ProgBar.Visible = True
Do Until llave_rep01.EOF
 DoEvents
 Chequeo.Caption = str(llave_rep01.AbsolutePosition) & " / " & str(llave_rep01.RowCount)
 ProgBar.Value = llave_rep01.AbsolutePosition
 DoEvents
 PS_REP02(0) = llave_rep01!FAR_CODCIA
 PS_REP02(1) = llave_rep01!FAR_fecha
 PS_REP02(2) = llave_rep01!FAR_NUMOPER
 PS_REP02(3) = llave_rep01!far_codclie
 llave_rep02.Requery
 If llave_rep02.EOF Then Stop
 If llave_rep02.RowCount > 1 Then Stop
 If llave_rep02!all_numfac = llave_rep01!far_numfac Then
'   Stop
 Else
  llave_rep02.Edit
  llave_rep02!all_numfac = llave_rep01!far_numfac
  llave_rep02.Update
 End If
 llave_rep01.MoveNext
Loop
ProgBar.Visible = False
MsgBox "proceso terminado"

End Sub

Private Sub fami_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Pantalla.Enabled Then Pantalla.SetFocus
End If
End Sub

Private Sub Form_Activate()
 If CmbCalidad.ListCount <> 0 Then CmbCalidad.ListIndex = 0
 If Wfile = "CAJA_GENERAL" Then
  cmdmoneda.SetFocus
 End If
End Sub


Private Sub Form_Load()
If LK_CODUSU = "ADMIN" Then Chequeo.Visible = True
CentrarFormulario MDIForm1, Me
Screen.MousePointer = 11

If retra_llave.EOF Then
   Screen.MousePointer = 0
   Exit Sub
End If
Screen.MousePointer = 0
Wfile = Trim(retra_llave(3))
WFORM = Trim(retra_llave(7))
lblreporte.Caption = Trim(retra_llave(1))
'lblmoneda.Visible = True
'cmdMoneda.Visible = True
If Trim(par_llave!PAR_DEFAULT_FAC) = "D" Then
  cmdmoneda.ListIndex = 1
Else
  cmdmoneda.ListIndex = 0
End If

If Wfile = "IMP_STOCK" Then
 chestock.Visible = True
 lblStock.Visible = True
 pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CP = 'P'  AND CLI_CODCIA = ? ORDER BY CLI_NOMBRE"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
 PS_REP01(0) = LK_CODCIA
 llave_rep01.Requery
 Do Until llave_rep01.EOF
     PROV.AddItem llave_rep01!CLI_NOMBRE & String(25, " ") & llave_rep01!cli_codclie
     llave_rep01.MoveNext
 Loop
 PROV.Visible = True
 PUB_CODCIA = LK_CODCIA
End If
If Wfile = "HSTOCK" Or Wfile = "STOCKVA" Or Wfile = "STOCKSEM" Then    'CRYSTAL REPORT
 'fami.Height = 1830
 PUB_CODCIA = LK_CODCIA
 LLENADOS fami, 122
 'FRASTOCK.Left = 2280
 FRASTOCK.Visible = True
 subfami.Visible = False
 PUB_CODCIA = LK_CODCIA
 If Wfile = "STOCKVA" Or Wfile = "STOCKSEM" Then
   CmbCalidad.Visible = False
   lblcalidad.Visible = False
 Else
   LLENADOS_COMBO CmbCalidad, 2
 End If
 fami.TabIndex = 0
End If
If Wfile = "IMP_PLANILLA" Or Wfile = "IMP_CONSUPLAN" Or Wfile = "IMP_COMISION" Or Wfile = "IMP_COMI_NETAS" Then
 cmdmoneda.Visible = True
 Dim codi As String * 5
 lblStock.Caption = "Vendedor : "
 lblStock.Visible = True
 pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = ? ORDER BY VEM_CODVEN"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 PS_REP01(0) = 0
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
 PS_REP01(0) = LK_CODCIA
 llave_rep01.Requery
 listven.Clear
 Do Until llave_rep01.EOF
     codi = llave_rep01!VEM_CODVEN
     listven.AddItem codi & Trim(llave_rep01!VEM_NOMBRE)
     llave_rep01.MoveNext
 Loop
 listven.Visible = True
 lblcampo1.Caption = "Fecha de Inicial : "
 lblcampo1.Visible = True
 'txtCampo1.MaxLength = 10
 txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
 If Wfile = "IMP_PLANILLA" Then
   cmdultima.Visible = True
 End If
 If Wfile <> "IMP_CONSUPLAN" Then
  lblcampo2.Caption = "Fecha de Final: "
  lblcampo2.Visible = True
  'txtCampo2.MaxLength = 10
  txtcampo2.Mask = "##/##/####"
  txtcampo2.Visible = True
 End If
 listven.TabIndex = 0
End If
If Wfile = "PARTE_DIARIO" Or Wfile = "VENTA_X_VEND" Or Wfile = "VEN_NEGOCIOS" Or Wfile = "VEN_NEGOCIOS_CLI" Or Wfile = "HER_NEGOCIOS" Or Wfile = "HER_VENDEDOR" Or Wfile = "HER_REGISTRO" Or Wfile = "REGCOMP" Or Wfile = "RESUMEN_TRANSA" Or Wfile = "SECUENCIA" Or Wfile = "RESU_VENTA_DIA" Or Wfile = "RESU_REGISTRO" Then
  If Wfile = "HER_VENDEDOR" Or Wfile = "HER_NEGOCIOS" Or Wfile = "VEN_NEGOCIOS" Then
   lblMoneda.Visible = True
   cmdmoneda.Visible = True
  End If
 If Wfile = "HER_REGISTRO" Or Wfile = "REGCOMP" Or Wfile = "RESU_VENTA_DIA" Or Wfile = "RESU_REGISTRO" Then
   ''If par_llave!PAR_CONTABILIDAD = "A" Then
    ''txtCampo1.Text = Format(cop_llave!cop_fecha_proceso, "dd/mm/yyyy")
    ''txtCampo2.Text = Format(cop_llave!COP_FECHA_PROCESO2, "dd/mm/yyyy")
    ''cheasiento.Visible = True
   ''Else
    txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    txtcampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
  '' End If
   max.Text = 65
   txttp.Clear
   txttp.AddItem "F= Facturas"
   txttp.AddItem "B= boletas"
   txttp.AddItem "N= N/Cred."
   txttp.AddItem "D= N/Deb."
   Frame2.Visible = True
   fracodclie.Visible = True
   checia.Visible = True
   lblMoneda.Visible = True
   cmdmoneda.Visible = True

   If Wfile = "RESU_VENTA_DIA" Then
       frapasa.Visible = True
      If Not cop_llave.EOF Then
         cop_fecha1.Caption = Format(cop_llave!cop_fecha_proceso, "dd/mm/yy")
         cop_fecha2.Caption = Format(cop_llave!cop_fecha_proceso2, "dd/mm/yy")
      End If
   End If
   
   loc_cp = "P"
 Else
   txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 End If
 lblcampo1.Caption = "Fecha de Inicial : "
 lblcampo1.Visible = True
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
 lblcampo2.Caption = "Fecha de Final: "
 lblcampo2.Visible = True
 txtcampo2.Mask = "##/##/####"
 txtcampo2.Visible = True
 If Wfile = "VENTA_X_VEND" Or Wfile = "HER_VENDEDOR" Or Wfile = "VEN_NEGOCIOS" Or Wfile = "VEN_NEGOCIOS_CLI" Then
    lblStock.Caption = "Vendedor : "
    lblStock.Visible = True
    pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = ? ORDER BY VEM_CODVEN"
    Set PS_REP01 = CN.CreateQuery("", pub_cadena)
    PS_REP01(0) = 0
    Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    PS_REP01(0) = LK_CODCIA
    llave_rep01.Requery
    vendmulti.Clear
    Do Until llave_rep01.EOF
        codi = llave_rep01!VEM_CODVEN
        vendmulti.AddItem codi & Trim(llave_rep01!VEM_NOMBRE)
        llave_rep01.MoveNext
    Loop
    vendmulti.Visible = True
    If Wfile = "VENTA_X_VEND" Or Wfile = "VEN_NEGOCIOS" Or Wfile = "VEN_NEGOCIOS_CLI" Then Chesup.Visible = True
 End If
 If Wfile = "PARTE_DIARIO" Then
    cheasiento.Visible = False
    checia.Visible = True
 End If
End If
If Wfile = "VENTA_X_VEND" Then 'Or Wfile = "HER_VENDEDOR" Then
    'fami.Height = 1830
    PUB_CODCIA = LK_CODCIA
    LLENADOS fami, 122
    'FRASTOCK.Left = 5980
    FRASTOCK.Visible = True
    subfami.Visible = False
    lblcalidad.Visible = False
    CmbCalidad.Visible = False
    lblMoneda.Visible = True
    cmdmoneda.Visible = True
End If
If Wfile = "DEDVENC" Then
    fracodclie.Visible = True
    fracodclie.Caption = "CLIENTE"
    fracodclie.Top = 3200
    fracodclie.Left = 3200
    loc_cp = "C"
    pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = ? ORDER BY VEM_CODVEN"
    Set PS_REP01 = CN.CreateQuery("", pub_cadena)
    PS_REP01(0) = 0
    Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    PS_REP01(0) = LK_CODCIA
    llave_rep01.Requery
    vendmulti.Clear
    Do Until llave_rep01.EOF
        codi = llave_rep01!VEM_CODVEN
        vendmulti.AddItem codi & Trim(llave_rep01!VEM_NOMBRE)
        llave_rep01.MoveNext
    Loop
    vendmulti.Visible = True
    fradeudas.Visible = True
    
    lblcampo1.Caption = "Fecha de Inicial : "
    lblcampo1.Visible = True
    txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    txtCampo1.Mask = "##/##/####"
    txtCampo1.Visible = True
    lblcampo2.Caption = "Fecha de Final: "
    lblcampo2.Visible = True
    txtcampo2.Mask = "##/##/####"
    txtcampo2.Visible = True
ElseIf Wfile = "COMISIONES" Then
    Pantalla.TabIndex = 0
    lblcampo1.Caption = "Fecha de Inicial : "
    lblcampo1.Visible = True
    txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    txtCampo1.Mask = "##/##/####"
    txtCampo1.Visible = True
    lblcampo2.Caption = "Fecha de Final: "
    lblcampo2.Visible = True
    txtcampo2.Mask = "##/##/####"
    txtcampo2.Visible = True
ElseIf Wfile = "ZONA_X_NEGO" Then
 PUB_CODCIA = "00"
 LLENADOS zonas, 35
 FraZonas.Visible = True
 opzonas(0).Caption = BUSCA_ETIQUETA(10)
 opzonas(1).Caption = BUSCA_ETIQUETA(11)
 opzonas(2).Caption = BUSCA_ETIQUETA(12)
 Chesup.Visible = True
ElseIf Left(Wfile, 9) = "MOROSIDAD" Then 'MIC
 fracli.Visible = True
 txtDias1(0).Text = "1"
 txtDias1(1).Text = "7"
 txtDias2(0).Text = "8"
 txtDias2(1).Text = "+"
 lblStock.Caption = "Vendedor : "
 lblStock.Visible = True
 pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = ? ORDER BY VEM_CODVEN"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 PS_REP01(0) = 0
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
 PS_REP01(0) = LK_CODCIA
 llave_rep01.Requery
 vendmulti.Clear
 Do Until llave_rep01.EOF
     codi = llave_rep01!VEM_CODVEN
     vendmulti.AddItem codi & Trim(llave_rep01!VEM_NOMBRE)
     llave_rep01.MoveNext
 Loop
 vendmulti.Visible = True
End If
If Wfile = "RESUMEN" Or Wfile = "POR_ANULADOS" Then
    lblcampo1.Caption = "Fecha de Inicial : "
    lblcampo1.Visible = True
    txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    txtCampo1.Mask = "##/##/####"
    txtCampo1.Visible = True
    lblcampo2.Caption = "Fecha de Final: "
    lblcampo2.Visible = True
    txtcampo2.Mask = "##/##/####"
    txtcampo2.Visible = True
    cmdmoneda.Visible = True

End If
If Wfile = "TRANSFER" Then
Frame1.BackColor = QBColor(14)
lblreporte.BackColor = QBColor(14)
lblreporte.ForeColor = QBColor(0)
lblcampo1.Caption = "Rango de Fechas para el Proceso :"
lblcampo1.Visible = True
txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
txtCampo1.Mask = "##/##/####"
txtCampo1.Visible = True
lblcampo2.Caption = ""
lblcampo2.Visible = True
txtcampo2.Mask = "##/##/####"
txtcampo2.Visible = True
   'Transferencia
End If

If Wfile = "REPO_CAJA_GEN" Then
 lblMoneda.Visible = True
 cmdmoneda.Visible = True
 cmdmoneda.ListIndex = 0
 cmdmoneda.TabIndex = 0
 lblFecha.Visible = True
 txtFecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtFecha.Mask = "##/##/####"
 txtFecha.Visible = True
End If

If Wfile = "REPO_CAJA_GEN_FECHA" Then
 lblMoneda.Visible = True
 cmdmoneda.Visible = True
 cmdmoneda.ListIndex = 0
 cmdmoneda.TabIndex = 0
 'Modificado 20042004
 lblcampo1.Caption = "Fecha de Inicial : "
 lblcampo1.Visible = True
 txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
 lblcampo2.Caption = "Fecha de Final: "
 lblcampo2.Visible = True
 txtcampo2.Mask = "##/##/####"
 txtcampo2.Visible = True
 
End If


If Wfile = "RESU_ENTREGA" Or Wfile = "RESU_CAJA" Or Wfile = "RESU_CAJA_VENUS" Or Wfile = "REPO_CAJA_DET" Or Wfile = "REPO_CAJA_DET2" Then
 If LK_EMP = "3AA" Then checia.Visible = True
 lblFecha.Visible = True
 txtFecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtFecha.Mask = "##/##/####"
 txtFecha.Visible = True
End If
If Wfile = "CAJA_GENERAL" Then
 lblMoneda.Visible = True
 cmdmoneda.Visible = True
 lblFecha.Visible = True
 txtFecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtFecha.Mask = "##/##/####"
 txtFecha.Visible = True
End If

If Wfile = "REPO_CAJA" Then
 lblMoneda.Visible = True
 cmdmoneda.Visible = True
 cmdmoneda.ListIndex = 0
 cmdmoneda.TabIndex = 0
 lblFecha.Visible = True
 txtFecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtFecha.Mask = "##/##/####"
 txtFecha.Visible = True
End If
If Wfile = "ULTIMAS_VENTAS" Then
 lblMoneda.Visible = True
 cmdmoneda.Visible = True
 cmdmoneda.ListIndex = 0
 cmdmoneda.TabIndex = 0
 
 PUB_CODCIA = "00"
 LLENADOS zonas, 35
 FraZonas.Visible = True
 opzonas(0).Enabled = False
 opzonas(1).Enabled = False
 opzonas(2).Caption = "Zonas"
End If
If Wfile = "PRO_CO" Then
    lblcampo1.Caption = "Fecha de Vcto. de Pago : "
    lblcampo1.Visible = True
    txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    txtCampo1.Mask = "##/##/####"
    txtCampo1.Visible = True
    Combo1.Visible = True
    PUB_TIPREG = 8
    PUB_CODCIA = "00"
    SQ_OPER = 2
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
        If tab_mayor!tab_NOMLARGO <> "TE" And tab_mayor!tab_NOMLARGO <> "CT" Then
          Combo1.AddItem tab_mayor!tab_NOMLARGO & "=" & tab_mayor!tab_nomcorto & String(30, " ") & Left(tab_mayor!TAB_contable2, 1)
        End If
        tab_mayor.MoveNext
    Loop
    chestands.Visible = True
    Combo1.ListIndex = 0
    
End If

If Wfile = "VENTAS_DEUDA" Then
    lblcampo1.Caption = "Desde la Fecha ..."
    lblcampo1.Visible = True
    txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    txtCampo1.Mask = "##/##/####"
    txtCampo1.Visible = True
    PUB_CODCIA = "00"
    LLENADOS solozonas, 35
    fzonas.Visible = True
    lblMoneda.Visible = False
    cmdmoneda.Visible = False
End If
If Wfile = "DEDVENC" Or Left(Wfile, 9) = "MOROSIDAD" Then 'MIC
   lblMoneda.Visible = True
   cmdmoneda.Visible = True
End If

If Wfile = "FECHA_STOCK" Or Wfile = "STOCXxUNIDAD" Then
  PUB_CODCIA = LK_CODCIA
  LLENADOS vendmulti, 122
  vendmulti.Visible = True
  lblcampo1.Caption = "Fecha para el Stock : "
  lblcampo1.Visible = True
  txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
  txtCampo1.Mask = "##/##/####"
  txtCampo1.Visible = True
  PUB_TIPREG = 2
  PUB_CODCIA = LK_CODCIA
  SQ_OPER = 2
  LEER_TAB_LLAVE
  listacal.ToolTipText = "TAB_TIPREG = 2"
  listacal.Clear
  Do Until tab_mayor.EOF
      listacal.AddItem tab_mayor!tab_NOMLARGO & String(80, " ") & tab_mayor!TAB_NUMTAB
      If tab_mayor!TAB_NUMTAB = 1 Then listacal.Selected(tab_mayor.AbsolutePosition - 1) = True
      tab_mayor.MoveNext
  Loop
 fracal.Visible = True
End If

  

If Wfile = "LISTA_PRECIOS" Then
  PUB_CODCIA = LK_CODCIA
  LLENADOS vendmulti, 122
  vendmulti.Visible = True
  PUB_CODCIA = "00"
  PUB_TIPREG = 600
  SQ_OPER = 2
  LEER_TAB_LLAVE
  listat.Clear
  Do Until tab_mayor.EOF
     listat.AddItem tab_mayor!tab_NOMLARGO & String(80, " ") & tab_mayor!TAB_NUMTAB
     tab_mayor.MoveNext
  Loop
  
  PUB_CODCIA = LK_CODCIA
  PUB_TIPREG = 333
  SQ_OPER = 2
  LEER_TAB_LLAVE
  listac.Clear
  Do Until tab_mayor.EOF
     listac.AddItem tab_mayor!tab_NOMLARGO & String(80, " ") & tab_mayor!TAB_NUMTAB
     tab_mayor.MoveNext
  Loop
  
  fradescto.Visible = True
End If


End Sub
Public Sub carga_cartera()
Dim wwww_flag  As String * 1
Dim WIMPORTE_ORIG As Currency
Dim ws_file As String
Dim CONTADOR As Integer
Dim wreal_cont As Integer
pub_mensaje = "CONTINUAR PROCESO ...   ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
pub_cadena = "SELECT CAA_NUM_OPER FROM CARACU WHERE CAA_CP = 'C'  AND CAA_CODCIA = ? ORDER BY CAA_NUM_OPER DESC"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
FILAX = 2
caa_histo.Requery
CONTADOR = 0
PUB_CP = "C"
FrmImp2.lblproceso.Caption = "Verificando Datos .... Espere !!!"
wreal_cont = 0
wwww_flag = ""
pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CP=? AND CLI_auto2  = ? AND CLI_CODCIA = ? ORDER BY CLI_CP ,CLI_CODCLIE"
Set PSCLI_LLAVE = CN.CreateQuery("", pub_cadena)
PSCLI_LLAVE(0) = ""
PSCLI_LLAVE(1) = 0
PSCLI_LLAVE(2) = ""
Set cli_llave = PSCLI_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
   xl.APPLICATION.Visible = True
Do Until CONTADOR = 600000
   If Trim(xl.Cells(FILAX, 12)) = "P" Or Trim(xl.Cells(FILAX, 12)) = "A" Then GoTo mas
   If Trim(xl.Cells(FILAX, 1)) = "" Then
       Exit Do
      CONTADOR = CONTADOR + 1
      GoTo mas
   End If
   wreal_cont = wreal_cont + 1

  
   PSCLI_LLAVE(0) = PUB_CP
   PSCLI_LLAVE(1) = Trim(xl.Cells(FILAX, 1))
   PSCLI_LLAVE(2) = LK_CODCIA
   cli_llave.Requery
   If cli_llave.EOF Then
      MsgBox "Apunte ...... Error en Codigo de cliente, NO EXISTE ...FILA: " & FILAX & " - " & Trim(xl.Cells(FILAX, 1))
      wwww_flag = "A"
   Else
'      MsgBox " existe"
   End If
   
   'If Trim(Left(cli_llave!cli_nombre, 15)) = Trim(Left(xl.Cells(FILAX, 6), 15)) Then
   'Else
   '   MsgBox "Error en Codigo de cliente, NO EXISTE ...FILA:" & FILAX
   '   GoTo CANCELA
   'End If
   
   SQ_OPER = 1
   PUB_CODVEN = xl.Cells(FILAX, 11)
   pu_codcia = LK_CODCIA
   LEER_VEN_LLAVE
   If ven_llave.EOF Then '
      MsgBox "Apunte.....Error en Codigo de Vendedor, NO EXISTE ...FILA: " & FILAX
      wwww_flag = "A"
       ' GoTo CANCELA
   End If
   
  ' If xl.Cells(FILAX, 1) = "F" Or xl.Cells(FILAX, 1) = "B" Then  'Or xl.Cells(FILAX, 1) = "N" Or xl.Cells(FILAX, 1) = "D" Then
   ' Else
  ''     MsgBox "Solo Válido F,B,N,D : en Fila: " & FILAX
   '   wwww_flag = "A"
   'End If
   
   If IsNumeric(xl.Cells(FILAX, 2)) = False Then
      MsgBox "Serie Invalido ...FILA: " & FILAX
      wwww_flag = "A"
      'GoTo CANCELA
   End If
  'If Left(xl.Cells(FILAX, 2), 3) = "" Then Stop
  
  'If Mid(xl.Cells(FILAX, 2), 4, Len(Trim(xl.Cells(FILAX, 2)))) = "" Then Stop

  ' If IsNumeric(xl.Cells(FILAX, 3)) = False Then
 '     MsgBox "N.Documento Invalido ...FILA: " & FILAX
 ''     wwww_flag = "A"
      'GoTo CANCELA
 ''   End If
   
mas:
   FILAX = FILAX + 1
  
Loop
If wwww_flag = "A" Then
  MsgBox "Verificar Información....."
  GoTo CANCELA
End If
MsgBox "Listo para Grabar ..!!!!"

Stop
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.max = FILAX
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Visible = True
DoEvents
PS_REP01(0) = LK_CODCIA
llave_rep01.Requery
If llave_rep01.EOF Then
   PUB_NUM_OPER_XXX = 1
Else
  PUB_NUM_OPER_XXX = llave_rep01!CAA_NUM_OPER + 1
End If
FILAX = 2
Do Until FILAX > FrmImp2.ProgBar.max
DoEvents
   If Trim(xl.Cells(FILAX, 12)) = "P" Or Trim(xl.Cells(FILAX, 12)) = "A" Then GoTo MAS2
   PSCLI_LLAVE(0) = "C"
   PSCLI_LLAVE(1) = Trim(xl.Cells(FILAX, 1))
   PSCLI_LLAVE(2) = LK_CODCIA
   cli_llave.Requery
   If cli_llave.EOF Then
   GoTo MAS2
   End If
   
   PUB_IMPORTE_AMORT = Format(val(xl.Cells(FILAX, 5)) - val(xl.Cells(FILAX, 6)), "0.000")
   WIMPORTE_ORIG = val(xl.Cells(FILAX, 5))
   If PUB_IMPORTE_AMORT = 0 Then
      GoTo MAS2
   End If
   wreal_cont = wreal_cont - 1
   FrmImp2.lblproceso.Caption = "Act.... " + str(wreal_cont) '+ "  -- " + Trim(xl.Cells(FILAX, 6))
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   DoEvents
   

 
   pu_codclie = cli_llave!cli_codclie
   
'  If Trim(xl.Cells(FILAX, 1)) = "F" Or Trim(xl.Cells(FILAX, 1)) = "B" Then
    PUB_TIPDOC = "FA"
'  End If
SQ_OPER = 3
pu_codcia = LK_CODCIA
PUB_SERDOC = 0
pu_cp = PUB_CP
DoEvents
LEER_CAR_LLAVE
DoEvents
If car_menor.EOF = True Then
   PUB_NUMDOC = 1
Else
   PUB_NUMDOC = car_menor!car_NUMDOC + 1
End If

'PUB_CODVEN = xl.Cells(FILAX, 1)

car_llave.AddNew
car_llave!CAR_CODCLIE = pu_codclie
car_llave!car_codcia = LK_CODCIA
car_llave!car_numguia = 0
car_llave!car_TIPDOC = PUB_TIPDOC
car_llave!CAR_cp = PUB_CP
car_llave!car_SERDOC = PUB_SERDOC
car_llave!car_NUMDOC = PUB_NUMDOC
car_llave!CAR_FECHA_INGR = LK_FECHA_DIA
car_llave!CAR_FECHA_SUNAT = xl.Cells(FILAX, 3)
car_llave!car_fecha_vcto = xl.Cells(FILAX, 16)
car_llave!car_fecha_vcto_orig = xl.Cells(FILAX, 16)
car_llave!CAR_SITUACION = 0
car_llave!CAR_COMISION = 0
car_llave!CAR_NUM_REN = 0
car_llave!car_concepto = "Saldo Inicial "
car_llave!car_nombre_banco = " "
car_llave!car_NUM_CHEQUE = 0
car_llave!car_SIGNO_CAJA = 0
car_llave!car_numguia = 0
car_llave!CAR_MONEDA = "S"
car_llave!CAR_IMP_INI = WIMPORTE_ORIG
car_llave!car_importe = PUB_IMPORTE_AMORT
car_llave!car_codtra = 0
car_llave!car_PRECIO = PUB_IMPORTE_AMORT 'Val(xl.Cells(FILAX, 5))
car_llave!CAR_signo_car = 1
car_llave!car_NUMSER = 51
car_llave!car_NUMFAC = Trim(xl.Cells(FILAX, 2))
car_llave!CAR_TIPMOV = 10
car_llave!car_FBG = "B"
car_llave!CAR_codven = xl.Cells(FILAX, 11)
car_llave!CAR_COBRADOR = xl.Cells(FILAX, 11)
car_llave!CAR_NUMSER_C = 0
car_llave!CAR_NUMFAC_C = 0
car_llave!CAR_codban = 0
car_llave.Update

'SQ_OPER = 1
'pu_codcia = LK_CODCIA
'pu_cp = PUB_CP
'' LEER_CLI_LLAVE

'cli_llave.Edit
'cli_llave!cli_SALDO = Nulo_Valor0(cli_llave!cli_SALDO) + PUB_IMPORTE_AMORT
'cli_llave.Update

caa_histo.AddNew
caa_histo!CAA_CODCLIE = pu_codclie
caa_histo!caa_codcia = LK_CODCIA
caa_histo!CAA_TIPDOC = PUB_TIPDOC
caa_histo!CAA_CP = PUB_CP
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
caa_histo!CAA_NUM_OPER = PUB_NUM_OPER_XXX
caa_histo!caa_serdoc = PUB_SERDOC
caa_histo!CAA_NUMDOC = PUB_NUMDOC
caa_histo!CAA_FECHA = LK_FECHA_DIA
caa_histo!CAA_FECHA_VCTO = xl.Cells(FILAX, 16)
caa_histo!CAA_FECHA_COBRO = xl.Cells(FILAX, 16)
caa_histo!caa_situacion = 0
caa_histo!caa_concepto = "Saldo Inicial"
caa_histo!CAA_IMPORTE = PUB_IMPORTE_AMORT
caa_histo!CAA_TOTAL = WIMPORTE_ORIG ' PUB_IMPORTE_AMORT
caa_histo!CAA_SALDO = Nulo_Valor0(cli_llave!cli_SALDO)
caa_histo!caa_SALDO_car = PUB_IMPORTE_AMORT
caa_histo!CAA_SIGNO_CAJA = 0
caa_histo!CAA_SIGNO_CAJA_REAL = 0
caa_histo!CAA_SIGNO_CAR = 1
caa_histo!CAA_TIPMOV = 10
caa_histo!CAA_hora = Now
caa_histo!CAA_CODUSU = LK_CODUSU
caa_histo!CAA_ESTADO = "N"
caa_histo!CAA_NUMPLAN = 0
caa_histo!CAa_NUM_CHEQUE = ""
caa_histo!CAa_numser = 51
caa_histo!CAa_numfac = Trim(xl.Cells(FILAX, 2))
caa_histo!caa_numser_c = 0
caa_histo!caa_numfac_c = 0
caa_histo!CAa_numGUIA = 0
caa_histo!CAA_FBG = "B"
caa_histo!CAA_CODVEN = xl.Cells(FILAX, 11)
caa_histo!caa_codban = 0
caa_histo.Update
MAS2:
FILAX = FILAX + 1
Loop
MsgBox "PROCESO TERMINADO SATISFACTORIAMENTE..", 48, Pub_Titulo
GoTo CANCELA

WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  If xl Is Nothing Then
    DoEvents
    Set xl = CreateObject("Excel.Application")
    DoEvents
  End If
  lblproceso.Caption = "Abriendo , Archivo carga.xls . . . "
  DoEvents
  'xl.Workbooks.Open Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\OFFICE\CARTERA.xls", 0, True, 4, WPAS
  xl.Workbooks.Open "C:\CARGA\ctacte.xls", 0, True, 4, WPAS
  Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  xl.APPLICATION.Visible = True
  Set xl = Nothing
  Screen.MousePointer = 0
  Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo :SALDO_CAR" & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Resume Next
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.APPLICATION.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
 
End Sub

Public Sub IMP_STOCK()
'On Error GoTo FINTODO
Dim ws_clave As String
Dim LETRAS(24) As String * 1
Dim WSFECHA As Date
Dim wcodclie
Dim wvalor
Dim ws_ingresos As Currency
Dim ws_salidas As Currency
Dim val_ingresos As Currency
Dim val_salidas As Currency
Dim acu_val_ingresos As Currency
Dim acu_val_salidas As Currency
Dim acu_val_saldos As Currency
Dim Wche As Integer
Dim wkSELECT As String
Dim WPROVE, wfami, WSUBFAMI
Dim WCHE1, WCHE2, WCHE3
Dim WNOMCLIE
If Trim(PROV.Text) = "" Then
  MsgBox "Seleccione Datos, para Procesar ", 48, Pub_Titulo
  GoTo CANCELA
End If
WPROVE = ""
WCHE1 = 0
wkSELECT = ""
For i = 0 To PROV.ListCount - 1
  PROV.ListIndex = i
  If PROV.Selected(i) Then
    WCHE1 = 1
    WPROVE = WPROVE + " ART_CODCLIE = " & Trim(Right(PROV.Text, 10)) & " OR"
  End If
Next i
If WCHE1 = 1 Then
 WPROVE = Left(WPROVE, Len(WPROVE) - 3)
Else
 WPROVE = ""
End If


wfami = ""
WCHE2 = 0
wkSELECT = ""
For i = 0 To fami.ListCount - 1
  fami.ListIndex = i
  If fami.Selected(i) Then
    WCHE2 = 1
    wfami = wfami + " ART_FAMILIA = " & Trim(Right(fami.Text, 10)) & " OR"
  End If
Next i
If WCHE2 = 1 Then
 wfami = Left(wfami, Len(wfami) - 3)
Else
 wfami = ""
End If

WSUBFAMI = ""
WCHE3 = 0
wkSELECT = ""
For i = 0 To subfami.ListCount - 1
  subfami.ListIndex = i
  If subfami.Selected(i) Then
    WCHE3 = 1
    WSUBFAMI = WSUBFAMI + " ART_SUBFAM = " & Trim(Right(subfami.Text, 10)) & " OR"
  End If
Next i
If WCHE3 = 1 Then
 WSUBFAMI = Left(WSUBFAMI, Len(WSUBFAMI) - 3)
Else
 WSUBFAMI = ""
End If


If WCHE1 = 0 And WCHE2 = 0 And WCHE3 = 0 Then
  MsgBox "Seleccione Datos, para Procesar ", 48, Pub_Titulo
  GoTo CANCELA
End If


wkSELECT = ""
If WPROVE <> "" Then
 wkSELECT = wkSELECT + Trim(WPROVE) + " OR"
End If
If wfami <> "" Then
 wkSELECT = wkSELECT + Trim(wfami) + " OR"
End If
If WSUBFAMI <> "" Then
 wkSELECT = wkSELECT + Trim(WSUBFAMI)
End If
If Right(wkSELECT, 1) = "R" Then
wkSELECT = Left(wkSELECT, Len(wkSELECT) - 3)
End If


pub_cadena = "SELECT ART_CODCIA, ART_KEY, ART_NOMBRE, ART_CODCLIE  FROM ARTI WHERE ART_CODCIA = ? AND ART_CALIDAD = 1 AND (" & wkSELECT & ") ORDER BY ART_CODCLIE, ART_FAMILIA, ART_SUBFAM, ART_NOMBRE"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
DoEvents
pub_cadena = "SELECT FAR_TIPMOV, FAR_CANTIDAD, FAR_CODCIA, FAR_CODART, FAR_SIGNO_ARM, FAR_PRECIO_NETO, FAR_PRECIO, FAR_EQUIV FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ? AND FAR_FECHA = ? AND FAR_ESTADO <> 'E' ORDER BY FAR_CODCIA, FAR_FECHA"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT PRE_CODART, PRE_EQUIV, PRE_UNIDAD FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ? AND PRE_FLAG_UNIDAD = 'A' ORDER BY PRE_CODCIA"
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


FrmImp2.ProgBar.Value = 1

'*** VERFICA SI HAY DATOS , O ESTAN CORRECTOS
PS_REP01(0) = LK_CODCIA
llave_rep01.Requery
If llave_rep01.EOF = True Then
   MsgBox "!!! NO EXISTEN Articulos ...", 48, Pub_Titulo
   GoTo CANCELA
End If
ws_clave = "0"
usu.Requery
Do Until usu.EOF
 If LK_CODUSU = "ADMIN" And Trim(usu!USU_KEY) = "ADMIN" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  If Trim(usu!USU_KEY) = "SUPERVISOR" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  usu.MoveNext
Loop

FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.max = llave_rep01.RowCount
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
xl.Worksheets(1).Activate
GoSub LETRAS
xcuenta = 0
xl.Cells(1, 1) = Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia")))
xl.Cells(2, 2) = "'" & LK_FECHA_DIA
f1 = 5  'Fila Inicial
wcodclie = llave_rep01!art_codclie
SQ_OPER = 1
pu_codcia = LK_CODCIA
pu_cp = "P"
pu_codclie = Nulo_Valor0(llave_rep01!art_codclie)
LEER_CLI_LLAVE
WNOMCLIE = ""
If cli_llave.EOF Then
'  MsgBox "Error en Codigo de cliente, NO EXISTE ...", 48, Pub_Titulo
'  GoTo CANCELA
Else
 WNOMCLIE = Trim(cli_llave!CLI_NOMBRE)
End If
FrmImp2.lblproceso.Caption = "Procesando . . .  un Momento ."
DoEvents
f1 = f1 + 1
xl.Cells(f1, 1) = WNOMCLIE
xl.Cells(f1, 1).Font.Bold = True
f1 = f1 + 1
fila = 0
acu_val_ingresos = 0
acu_val_salidas = 0
acu_val_saldos = 0
Do Until llave_rep01.EOF
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   If wcodclie <> llave_rep01!art_codclie Then
        SQ_OPER = 1
        pu_codcia = LK_CODCIA
        pu_cp = "P"
        pu_codclie = llave_rep01!art_codclie
        LEER_CLI_LLAVE
        WNOMCLIE = ""
        If cli_llave.EOF Then
         ' MsgBox "Error en Codigo de cliente, NO EXISTE ...", 48, Pub_Titulo
'          GoTo CANCELA
        Else
          WNOMCLIE = Trim(cli_llave!CLI_NOMBRE)
        End If
        'F1 = F1 + 1
        xl.Cells(f1, 1) = WNOMCLIE
        xl.Cells(f1, 1).Font.Bold = True
        wcodclie = llave_rep01!art_codclie
        f1 = f1 + 1
   End If
   PUB_CODART = llave_rep01!ART_KEY
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ARM_LLAVE
   If arm_llave.EOF Then
      MsgBox "Error Grave en Articulo..."
      GoTo CANCELA
   End If
   PS_REP03(0) = LK_CODCIA
   PS_REP03(1) = llave_rep01!ART_KEY
   llave_rep03.Requery
   If llave_rep03.EOF = True Then
      MsgBox "Error en la unidad. " & llave_rep01!art_nombre
      GoTo CANCELA
   End If
   ws_ingresos = 0
   ws_salidas = 0
   val_ingresos = 0
   val_salidas = 0
   If chestock.Value = 0 Then
      GoTo PASADIA
   End If
   PS_REP02(0) = LK_CODCIA
   PS_REP02(1) = llave_rep01!ART_KEY
   PS_REP02(2) = LK_FECHA_DIA
   llave_rep02.Requery
   If llave_rep02.EOF = True Then
      GoTo PASADIA
   End If
   Do Until llave_rep02.EOF
   If llave_rep02!far_signo_arm = 1 Then
     ws_ingresos = ws_ingresos + (llave_rep02!far_cantidad / llave_rep03!PRE_EQUIV)
     If llave_rep02!FAR_TIPMOV = 20 Then
        val_ingresos = val_ingresos + llave_rep02!far_precio_neto
     Else
       val_ingresos = val_ingresos + ((llave_rep02!far_cantidad / llave_rep03!PRE_EQUIV) * llave_rep02!far_precio_neto)
     End If
   ElseIf llave_rep02!far_signo_arm = -1 Then
     ws_salidas = ws_salidas + (llave_rep02!far_cantidad / llave_rep03!PRE_EQUIV)
     val_salidas = val_salidas + ((llave_rep02!far_cantidad / llave_rep02!FAR_equiv) * llave_rep02!FAR_PRECIO)
   End If
   llave_rep02.MoveNext
   Loop
PASADIA:
   xl.Cells(f1, 1) = Trim(llave_rep01!art_nombre)
   xl.Cells(f1, 2) = Left(llave_rep03!pre_unidad, 10)
   If chestock.Value = 1 Then
     If ws_ingresos <> 0 Then xl.Cells(f1, 3) = ws_ingresos
     If ws_salidas <> 0 Then xl.Cells(f1, 4) = ws_salidas
     'xl.Cells(f1, 6) = Format(llave_rep01!ART_COSPRO * llave_rep03!PRE_EQUIV, "0.000")
     If val_ingresos <> 0 Then xl.Cells(f1, 7) = val_ingresos
     If val_salidas <> 0 Then xl.Cells(f1, 8) = val_salidas
     'xl.Cells(f1, 9) = (arm_llave!arm_stock / llave_rep03!PRE_EQUIV) * (llave_rep01!ART_COSPRO * llave_rep03!PRE_EQUIV)
     acu_val_saldos = acu_val_saldos + val(xl.Cells(f1, 9))
     acu_val_ingresos = acu_val_ingresos + val_ingresos
     acu_val_salidas = acu_val_salidas + val_salidas
   End If
   xl.Cells(f1, 5) = Format(arm_llave!arm_stock / llave_rep03!PRE_EQUIV, "0.000")
   f1 = f1 + 1
   llave_rep01.MoveNext
Loop
  If chestock.Value = 1 Then
    xl.Cells(f1, 1) = "Totales "
    xl.Cells(f1, 7) = acu_val_ingresos
    xl.Cells(f1, 7).Font.Bold = True
    xl.Cells(f1, 8) = acu_val_salidas
    xl.Cells(f1, 8).Font.Bold = True
    xl.Cells(f1, 9) = acu_val_saldos
    xl.Cells(f1, 9).Font.Bold = True
  End If
  If chestock.Value = 0 Then
   xl.Range("C4:D5").Delete 4
   xl.Range("D3:G5").Delete 4
  End If
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False

Exit Sub


LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"
Return

WEXCEL:
  FrmImp2.lblproceso.Caption = "Abriendo , Archivo Stock.xls . . . "
  DoEvents
  WPAS = ws_clave
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\OFFICE\stock.xls", 0, True, 4, WPAS, WPAS
Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.APPLICATION.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
 
End Sub


Private Sub listVen_Click()
 If txtCampo1.Visible Then
  txtCampo1.SetFocus
 ElseIf Pantalla.Enabled Then
  Pantalla.SetFocus
 End If

End Sub

Private Sub listven_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If txtCampo1.Visible Then
  txtCampo1.SetFocus
 ElseIf Pantalla.Enabled Then
  Pantalla.SetFocus
 End If
End If
End Sub

Private Sub ListView2_DblClick()
 loc_key = ListView2.SelectedItem.Index
 txt_cli.Text = Trim(ListView2.ListItems.Item(loc_key).Text) & " "
 txt_cli_KeyPress 13
End Sub

Private Sub ListView2_GotFocus()
If loc_key <> 0 Then
 Set ListView2.SelectedItem = ListView2.ListItems(loc_key)
 ListView2.ListItems.Item(loc_key).Selected = True
 ListView2.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub ListView2_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView2.SelectedItem.Index
 txt_cli.Text = Trim(ListView2.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub ListView2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 ListView2.Visible = False
 txt_cli.Text = ""
 txt_cli.SetFocus
 Exit Sub
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
ListView2_DblClick

End Sub

Private Sub opzonas_Click(Index As Integer)
Dim cod As Integer
lblZonas.Caption = Trim(opzonas(Index).Caption) & " :"
If Index = 0 Then
  cod = 20
ElseIf Index = 1 Then
  cod = 30
ElseIf Index = 2 Then
  cod = 35
End If
PUB_CODCIA = "00"
LLENADOS zonas, cod
zonas.SetFocus
End Sub

Private Sub Pantalla_Click()
Dim wsFECHA1
Dim wsFECHA2
'On Error GoTo SALE
If Wfile = "IMP_STOCK" Then
  Call IMP_STOCK
ElseIf Wfile = "COMISIONES" Then
  Call COMISIONES
ElseIf Wfile = "IMP_PLANILLA" Or Wfile = "IMP_CONSUPLAN" Or Wfile = "IMP_COMISION" Then
   If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
   Else
     wsFECHA1 = Trim(txtCampo1.Text)
   End If
   If Not IsDate(wsFECHA1) Or val(Mid(wsFECHA1, 4, 2)) > 12 Then
     MsgBox "Fecha invalida .... ", 48, Pub_Titulo
     txtCampo1.SetFocus
     Exit Sub
   End If
   If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
   Else
     wsFECHA2 = Trim(txtcampo2.Text)
   End If
   If txtcampo2.Visible Then
    If Not IsDate(wsFECHA2) Or val(Mid(wsFECHA2, 4, 2)) > 12 Then
      MsgBox "Fecha invalida .... ", 48, Pub_Titulo
      Azul2 txtCampo1, txtCampo1
      Exit Sub
    End If
   End If
   If Wfile <> "IMP_CONSUPLAN" Then
    If CDate(wsFECHA1) > CDate(wsFECHA2) Then
      MsgBox "Fecha invalida .... ", 48, Pub_Titulo
      Azul2 txtCampo1, txtCampo1
      Exit Sub
    End If
   End If
   If Trim(listven.Text) = "" Then
     MsgBox "Seleccione vendedor ... ", 48, Pub_Titulo
     Exit Sub
   End If
   If Wfile = "IMP_PLANILLA" Then
    Call IMP_PLANILLA
   ElseIf Wfile = "IMP_CONSUPLAN" Then
    Call IMP_CONSUPLAN
   ElseIf Wfile = "IMP_COMISION" Then
    Call IMP_COMISION
   End If
   
ElseIf Wfile = "PARTE_DIARIO" Or Wfile = "VENTA_X_VEND" Or Wfile = "VEN_NEGOCIOS" Or Wfile = "VEN_NEGOCIOS_CLI" Or Wfile = "HER_NEGOCIOS" Or Wfile = "HER_VENDEDOR" Or Wfile = "HER_REGISTRO" Or Wfile = "REGCOMP" Or Wfile = "RESU_VENTA_DIA" Or Wfile = "SECUENCIA" Or Wfile = "RESU_REGISTRO" Then
   If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
   Else
     wsFECHA1 = Trim(txtCampo1.Text)
   End If
   If Not IsDate(wsFECHA1) Or val(Mid(wsFECHA1, 4, 2)) > 12 Then
     MsgBox "Fecha invalida .... ", 48, Pub_Titulo
     Azul2 txtCampo1, txtCampo1
     Exit Sub
   End If
   If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
   Else
     wsFECHA2 = Trim(txtcampo2.Text)
   End If
   If txtcampo2.Visible Then
    If Not IsDate(wsFECHA2) Or val(Mid(wsFECHA2, 4, 2)) > 12 Then
      MsgBox "Fecha invalida .... ", 48, Pub_Titulo
      Azul2 txtcampo2, txtcampo2
      Exit Sub
    End If
   End If
   If CDate(wsFECHA1) > CDate(wsFECHA2) Then
     MsgBox "Fecha invalida .... ", 48, Pub_Titulo
     Azul2 txtCampo1, txtCampo1
     Exit Sub
   End If
   If Wfile = "HER_NEGOCIOS" Then
    Call HER_NEGOCIOS2
   ElseIf Wfile = "HER_VENDEDOR" Then
    Call HER_VENDEDOR
   ElseIf Wfile = "HER_REGISTRO" Then
    Call REG_VENTA
    'Call HER_REGISTRO
   ElseIf Wfile = "REGCOMP" Then
    Call REG_COMPRA
   ElseIf Wfile = "VEN_NEGOCIOS" Then
    Call VEN_NEGOCIOS
   ElseIf Wfile = "VEN_NEGOCIOS_CLI" Then
    Call VEN_NEGOCIOS_CLI
   ElseIf Wfile = "VENTA_X_VEND" Then
    Call VENTA_X_VEND
   ElseIf Wfile = "PARTE_DIARIO" Then
    Call PARTE_DIARIO
   ElseIf Wfile = "RESU_VENTA_DIA" Then
    Call RESU_VENTA_DIA
   ElseIf Wfile = "SECUENCIA" Then
     Call SECUENCIA
   ElseIf Wfile = "RESU_REGISTRO" Then
     Call RESU_REGISTRO
   End If
End If
If Wfile = "DEDVENC" Then
 Call DEDVENC
ElseIf Wfile = "HSTOCK" Then  'CRYSTAL REPORT
 Call HSTOCK
ElseIf Wfile = "STOCKVA" Or Wfile = "STOCKSEM" Then
 Call STOCKVA
ElseIf Wfile = "ZONA_X_NEGO" Then
 Call ZONA_X_NEGO
ElseIf Left(Wfile, 9) = "MOROSIDAD" Then
    If Wfile = "MOROSIDAD0" Then
        Call MOROSIDAD(0)
    Else
        Call MOROSIDAD(1)
    End If
 
ElseIf Wfile = "REPO_CAJA_GEN" Then
 Call REPO_CAJA_GEN
ElseIf Wfile = "REPO_CAJA" Then
  Call REPO_CAJA_GEN
ElseIf Wfile = "REPO_CAJA_GEN_FECHA" Then ' Modificado
  Call REPO_CAJA_GEN_FECHA
ElseIf Wfile = "PRO_CO" Then
 Call PRO_CO
ElseIf Wfile = "VENTAS_DEUDA" Then
 Call VENTAS_DEUDA
ElseIf Wfile = "ULTIMAS_VENTAS" Then
 Call ULTIMAS_VENTAS
ElseIf Wfile = "REPO_CAJA_DET" Then
 Call REPO_CAJA_DET
ElseIf Wfile = "REPO_CAJA_DET2" Then
 Call REPO_CAJA_DET2
ElseIf Wfile = "RESU_CAJA" Then
 Call RESU_CAJA
ElseIf Wfile = "RESU_CAJA_VENUS" Then
 Call RESU_CAJA_VENUS
ElseIf Wfile = "RESU_ENTREGA" Then
 Call RESU_ENTREGA
ElseIf Wfile = "CARGA_CARTERA" Then
 Call carga_cartera
ElseIf Wfile = "FECHA_STOCK" Then
 Call FECHA_STOCK
ElseIf Wfile = "STOCXxUNIDAD" Then
 Call STOCKxUNIDAD_EXCEL(" ")
ElseIf Wfile = "LISTA_PRECIOS" Then
 Call LISTA_PRECIOS
ElseIf Wfile = "IMP_COMI_NETAS" Then
 Call IMP_COMI_NETAS
ElseIf Wfile = "POR_ANULADOS" Then
 Call POR_ANULADOS
End If
If Wfile = "TRANSFER" Then
Call Transferencia
End If
If Wfile = "CAJA_GENERAL" Then
  Call CAJA_GENERAL
End If

Exit Sub
SALE:
ProgBar.Visible = False
lblproceso.Visible = False
Pantalla.Enabled = True
cerrar.Enabled = True
MsgBox Err.Description + "Intente Nuevamente.", 48, Pub_Titulo
End Sub
Public Sub IMP_PLANILLA()
'On Error GoTo FINTODO
Dim ws_clave As String
Dim LETRAS(24) As String * 1
Dim WSFECHA As Date
Dim wcodven As Integer
Dim wvalor
Dim ws_ingresos As Currency
Dim ws_salidas As Currency
Dim val_ingresos As Currency
Dim val_salidas As Currency
Dim acu_val_ingresos As Currency
Dim acu_val_salidas As Currency
Dim Wche As Integer
Dim wkSELECT As String
Dim wsfile As String
wsfile = ""
Dim wsFECHA1, wsFECHA2
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If

If Trim(listven.Text) = "" Then
  MsgBox "Seleccione Datos, para Procesar ", 48, Pub_Titulo
  GoTo CANCELA
End If
Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
wcodven = val(Left(listven.Text, 4))
pub_cadena = "SELECT * FROM CARTERA WHERE CAR_CODCIA = ? AND CAR_COBRADOR = ? AND CAR_FECHA_VCTO >= ? and CAR_FECHA_VCTO <= ? AND CAR_IMPORTE <> 0  ORDER BY CAR_CODCLIE, CAR_FECHA_VCTO, CAR_IMPORTE"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
DoEvents

'*** VERFICA SI HAY DATOS , O ESTAN CORRECTOS
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = wcodven
PS_REP01(2) = wsFECHA1
PS_REP01(3) = wsFECHA2
llave_rep01.Requery
If llave_rep01.EOF = True Then
   MsgBox "!!! NO EXISTEN Datos ...", 48, Pub_Titulo
   GoTo CANCELA
End If
ws_clave = "0"
usu.Requery
Do Until usu.EOF
 If LK_CODUSU = "ADMIN" And Trim(usu!USU_KEY) = "ADMIN" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  If Trim(usu!USU_KEY) = "SUPERVISOR" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  usu.MoveNext
Loop
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.max = llave_rep01.RowCount
FrmImp2.ProgBar.Value = 0
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
GoSub LETRAS
FrmImp2.ProgBar.Visible = True
DoEvents
xcuenta = 0
xl.Cells(1, 1) = Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia")))
xl.Cells(2, 2) = "'" & LK_FECHA_DIA
xl.Cells(3, 1) = "Vendedor"

xl.Cells(3, 2) = Trim(listven.Text)
f1 = 5  'Fila Inicial

FrmImp2.lblproceso.Caption = "Procesando . . .  un Momento ."
DoEvents
fila = 0
acu_val_ingresos = 0
acu_val_salidas = 0
Do Until llave_rep01.EOF
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   f1 = f1 + 1
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   pu_codclie = llave_rep01!CAR_CODCLIE
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
      MsgBox "Error en Codigo de cliente, NO EXISTE ...", 48, Pub_Titulo
      GoTo CANCELA
   End If
   xl.Cells(f1, 1) = Trim(cli_llave!cli_codclie)
   xl.Cells(f1, 2) = Trim(cli_llave!CLI_NOMBRE)
'   xl.Cells(F1, 2).Font.Bold = True
  ' xl.Cells(F1, 3) = llave_rep01!CAR_CODVEN
   xl.Cells(f1, 3).HorizontalAlignment = xlCenter
   If llave_rep01!car_FBG = "F" Then
     xl.Cells(f1, 3) = "FAC."
   ElseIf llave_rep01!car_FBG = "B" Then
     xl.Cells(f1, 3) = "BOL."
   ElseIf llave_rep01!car_FBG = "G" Then
     xl.Cells(f1, 3) = "GUIA"
   End If
   xl.Cells(f1, 4).HorizontalAlignment = xlCenter
   xl.Cells(f1, 4) = llave_rep01!car_NUMSER
   xl.Cells(f1, 5).HorizontalAlignment = xlCenter
   xl.Cells(f1, 5) = llave_rep01!car_NUMFAC
   xl.Cells(f1, 5).HorizontalAlignment = xlCenter
   xl.Cells(f1, 6) = "'" & llave_rep01!CAR_FECHA_INGR
   xl.Cells(f1, 7) = "'" & llave_rep01!car_fecha_vcto
   xl.Cells(f1, 8) = llave_rep01!car_importe
   xl.Cells(f1, 8).HorizontalAlignment = xlRight
   xl.Cells(f1, 9).NumberFormat = "#######.000"
   xl.Cells(f1, 10).NumberFormat = "dd/mm/yyyy"
   xl.Cells(f1, 11) = ""
   xl.Cells(f1, 12) = llave_rep01!CAR_CODCLIE
   xl.Cells(f1, 13) = llave_rep01!car_codcia
   xl.Cells(f1, 14) = llave_rep01!car_SERDOC
   xl.Cells(f1, 15) = llave_rep01!car_NUMDOC
   xl.Cells(f1, 16) = llave_rep01!car_TIPDOC
   
   llave_rep01.MoveNext
Loop
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(2, 2) = "V I S I T A S   A   C L I E N T E S"
  xl.ActiveCell.Range("I6").Activate
  wranF = "A" & 1 & ":J" & f1
  xl.Worksheets(1).Range(wranF).Font.Name = "Draft 17cpi"
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  xl.Workbooks(1).Save
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub

LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"
Return

WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  
  wsfile1 = "PlanVen" & wcodven & ".XLS"
  wsfile = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\OFFICE\" & wsfile1
  On Error GoTo OJO
  Kill wsfile
  On Error GoTo 0
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  FrmImp2.lblproceso.Caption = "Configurando Hoja de Calculo... un Momento ."
  DoEvents
  xl.SheetsInNewWorkbook = 1
  xl.Workbooks.Add
  xl.Worksheets(1).Name = "COBRANZAS"
  xl.Windows(1).Caption = " Vendedor :  " & Trim(listven.Text)
  xl.Windows(1).WindowState = xlMaximized
  xl.Workbooks(1).SaveAs wsfile
  xl.ActiveWindow.Zoom = 83
  xl.Worksheets(1).Columns("A").ColumnWidth = 6
  xl.Worksheets(1).Columns("B").ColumnWidth = 25
  xl.Worksheets(1).Columns("C").ColumnWidth = 4
  xl.Worksheets(1).Columns("D").ColumnWidth = 4
  xl.Worksheets(1).Columns("F").ColumnWidth = 7
  xl.Worksheets(1).Columns("F").ColumnWidth = 10
  xl.Worksheets(1).Columns("G").ColumnWidth = 10
  xl.Worksheets(1).Columns("H").ColumnWidth = 9
  xl.Worksheets(1).Columns("I").ColumnWidth = 9
  xl.Worksheets(1).Columns("J").ColumnWidth = 10
  xl.Worksheets(1).Columns("K").ColumnWidth = 10 'pu_cp
  xl.Worksheets(1).Columns("L").ColumnWidth = 0 'pu_codclie
  xl.Worksheets(1).Columns("M").ColumnWidth = 0 ' pu_codcia
  xl.Worksheets(1).Columns("N").ColumnWidth = 0 'PUB_SERDOC
  xl.Worksheets(1).Columns("O").ColumnWidth = 0 'PUB_NUMDOC
  xl.Worksheets(1).Columns("P").ColumnWidth = 0 'PUB_TIPDOC
  xl.Worksheets(1).Rows(4).RowHeight = 15
  xl.Worksheets(1).Rows(5).RowHeight = 15
  xl.Range("A4:K5").Font.Bold = True
  xl.Range("A4:J5").HorizontalAlignment = xlCenter
  xl.Cells(5, 1) = "Codigo"
  xl.Cells(5, 2) = "C L I E N T E "
  xl.Cells(5, 3) = " "
  xl.Cells(4, 4) = "Nro."
  xl.Cells(5, 4) = "Serie"
  xl.Cells(4, 5) = "Nro."
  xl.Cells(5, 5) = "Doc."
  xl.Cells(4, 6) = "Fecha"
  xl.Cells(5, 6) = "Emisión"
  xl.Cells(4, 7) = "Fecha"
  xl.Cells(5, 7) = "Vcto."
  xl.Cells(5, 8) = "Saldo"
  xl.Cells(4, 9) = "Nuevos"
  xl.Cells(5, 9) = "Pagos"
  xl.Cells(4, 10) = "Nuevos "
  xl.Cells(5, 10) = "Vcto."
  xl.Cells(5, 11) = "Cheque."
  xl.APPLICATION.Visible = False
  With xl.Worksheets(1).PageSetup
    .TopMargin = 28.6
    .HeaderMargin = 28.6
    .PrintTitleRows = "$1:$5"
  End With
  FrmImp2.lblproceso.Caption = "Abriendo , Archivo PLANILLAS.xls . . . "
  DoEvents
  WPAS = "131296"
  'Set xl = Nothing
Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Resume Next
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.APPLICATION.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
 
End Sub

Public Sub IMP_CONSUPLAN()
'On Error GoTo FINTODO
Dim ws_clave As String
Dim LETRAS(24) As String * 1
Dim WSFECHA As Date
Dim wcodven As Integer
Dim wvalor
Dim ws_ingresos As Currency
Dim ws_salidas As Currency
Dim val_ingresos As Currency
Dim val_salidas As Currency
Dim acu_val_ingresos As Currency
Dim acu_val_salidas As Currency
Dim Wche As Integer
Dim wkSELECT As String
Dim wsfile As String
Dim wsFECHA1, wsFECHA2
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If

wsfile = ""
Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
wcodven = val(Left(listven.Text, 4))
pub_cadena = "SELECT * FROM CARACU WHERE CAA_CODCIA = ? AND CAA_CODVEN = ? AND CAA_FECHA = ? AND CAA_ESTADO <> 'E' AND CAA_SIGNO_CAR = -1 ORDER BY CAA_CODCLIE, CAA_FECHA,CAA_NUM_OPER, CAA_SALDO_CAR"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
DoEvents

PS_REP01(0) = LK_CODCIA
PS_REP01(1) = wcodven
PS_REP01(2) = wsFECHA1
llave_rep01.Requery
If llave_rep01.EOF = True Then
   MsgBox "!!! NO EXISTEN Datos ...", 48, Pub_Titulo
   GoTo CANCELA
End If
ws_clave = "0"
usu.Requery
Do Until usu.EOF
 If LK_CODUSU = "ADMIN" And Trim(usu!USU_KEY) = "ADMIN" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  If Trim(usu!USU_KEY) = "SUPERVISOR" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  usu.MoveNext
Loop
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.max = llave_rep01.RowCount
FrmImp2.ProgBar.Value = 0
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
FrmImp2.ProgBar.Visible = True
DoEvents
xl.Worksheets(1).Activate
GoSub LETRAS
xcuenta = 0
xl.Cells(1, 1) = Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia")))
xl.Cells(2, 2) = "'" & LK_FECHA_DIA
xl.Cells(3, 2) = "Vendedor : " & wcodven
f1 = 5  'Fila Inicial

FrmImp2.lblproceso.Caption = "Procesando . . .  un Momento ."
DoEvents
fila = 0
acu_val_ingresos = 0
acu_val_salidas = 0
Do Until llave_rep01.EOF
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   f1 = f1 + 1
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   pu_codclie = llave_rep01!CAA_CODCLIE
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
      MsgBox "Error en Codigo de cliente, NO EXISTE ...", 48, Pub_Titulo
      GoTo CANCELA
   End If
   xl.Cells(f1, 1) = Trim(cli_llave!cli_codclie)
   xl.Cells(f1, 2) = Trim(cli_llave!CLI_NOMBRE)
   xl.Cells(f1, 2).Font.Bold = True
   xl.Cells(f1, 3).HorizontalAlignment = xlCenter
   If llave_rep01!CAA_FBG = "F" Then
     xl.Cells(f1, 4) = "FAC."
   ElseIf llave_rep01!CAA_FBG = "B" Then
     xl.Cells(f1, 4) = "BOL."
   ElseIf llave_rep01!CAA_FBG = "G" Then
     xl.Cells(f1, 4) = "GUIA"
   End If
   xl.Cells(f1, 4).HorizontalAlignment = xlCenter
   xl.Cells(f1, 5) = "'" & llave_rep01!CAa_numser & " - " & llave_rep01!CAa_numfac
   xl.Cells(f1, 5).HorizontalAlignment = xlCenter
   xl.Cells(f1, 6) = "'" & llave_rep01!CAA_FECHA
   xl.Cells(f1, 7) = "'" & llave_rep01!CAA_FECHA_VCTO
   xl.Cells(f1, 8) = llave_rep01!CAA_SALDO - llave_rep01!CAA_IMPORTE
   xl.Cells(f1, 9).HorizontalAlignment = xlRight
   xl.Cells(f1, 9) = val(llave_rep01!CAA_IMPORTE)
   xl.Cells(f1, 10) = "'" & llave_rep01!CAA_FECHA_VCTO
   llave_rep01.MoveNext
Loop
  wran1 = "I" & 6
  wran2 = "I" & f1
  wranF = "I" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(2, 2) = "I N F O R M E  D E  C O B R A N Z A"
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub

LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"
Return

WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblproceso.Caption = "Abriendo , Archivo CONSUPLAN.xls . . . "
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\OFFICE\CONSUPLAN.xls", 0, True, 4, WPAS, WPAS
 
Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.APPLICATION.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
 
End Sub

Private Sub solozonas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Pantalla.Enabled Then Pantalla.SetFocus
End If
End Sub

Private Sub txt_cli_GotFocus()
Azul txt_cli, txt_cli
lblCliente.Caption = ""
End Sub
Private Sub txt_cli_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView2.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txt_cli.Text = "" Then
  loc_key = 1
  Set ListView2.SelectedItem = ListView2.ListItems(loc_key)
  ListView2.ListItems.Item(loc_key).Selected = True
  ListView2.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView2.ListItems.count Then loc_key = ListView2.ListItems.count
  GoTo posicion
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo posicion
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView2.ListItems.count Then loc_key = ListView2.ListItems.count
 GoTo posicion
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo posicion
End If
GoTo fin
posicion:
'  KeyCode = 0
  ListView2.ListItems.Item(loc_key).Selected = True
  ListView2.ListItems.Item(loc_key).EnsureVisible
  txt_cli.Text = Trim(ListView2.ListItems.Item(loc_key).Text) & " "
  DoEvents
  txt_cli.SelStart = Len(txt_cli.Text)
  DoEvents
fin:

End Sub
Private Sub txt_cli_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem    ' Variable FoundItem.
If KeyAscii = 27 Then
 ListView2.Visible = False
 txt_cli.Text = ""
 Exit Sub
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo ERROR_CODIGO
pu_codclie = val(txt_cli.Text)
On Error GoTo 0
If Len(txt_cli.Text) = 0 Then
   Exit Sub
End If

If pu_codclie <> 0 And IsNumeric(txt_cli.Text) = True Then
   SQ_OPER = 1
   pu_cp = loc_cp
   pu_codcia = LK_CODCIA
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
     lblCliente.Caption = ""
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Azul txt_cli, txt_cli
     GoTo fin
   Else
     lblCliente.Caption = Trim(cli_llave!CLI_NOMBRE)
     'LOC_RUC = Trim(cli_llave!cli_ruc_esposo)
   End If
   If Pantalla.Visible And Pantalla.Enabled Then
     Pantalla.SetFocus
   End If
Else
   If loc_key > ListView2.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(ListView2.ListItems.Item(loc_key).Text)
   If Trim(UCase(txt_cli.Text)) = Left(valor, Len(Trim(txt_cli.Text))) Then
   Else
      Exit Sub
   End If
   txt_cli.Text = Trim(ListView2.ListItems.Item(loc_key).SubItems(1))
   pu_codclie = val(txt_cli.Text)
   SQ_OPER = 1
   pu_cp = loc_cp
   pu_codcia = LK_CODCIA
   LEER_CLI_LLAVE
   If Not cli_llave.EOF Then
    lblCliente.Caption = Trim(ListView2.ListItems.Item(loc_key).Text)
    'LOC_RUC = Trim(cli_llave!cli_ruc_esposo)
   End If
   
   If Pantalla.Visible And Pantalla.Enabled Then
     Pantalla.SetFocus
   End If
End If

dale:
ListView2.Visible = False
fin:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul txt_cli, txt_cli

End Sub

Private Sub txt_cli_KeyUp(KeyCode As Integer, Shift As Integer)
Dim VAR
If Len(txt_cli.Text) = 0 Or IsNumeric(txt_cli.Text) = True Then
   ListView2.Visible = False
   Exit Sub
End If
If ListView2.Visible = False And KeyCode <> 13 Then
    VAR = Asc(txt_cli.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    ElseIf VAR = 58 Then
       VAR = "A"
    Else
       VAR = Chr(VAR)
    End If
    numarchi = 1
    'archi = "SELECT CLI_CODCLIE, CLI_CODCIA, CLI_CP, CLI_NOMBRE,CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM  FROM CLIENTES WHERE  CLI_CP = '" & loc_cp & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txt_cli.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
    archi = "SELECT CLI_CODCLIE , CLI_CODCIA, CLI_CP, CLI_NOMBRE, CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM, TAB_NOMLARGO  FROM CLIENTES,TABLAS WHERE (TAB_CODCIA = '00') AND (TAB_TIPREG = 35) AND (TAB_NUMTAB = CLI_ZONA_NEW) AND CLI_CP = '" & loc_cp & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txt_cli.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
'    If Trim(txt_cli.text) <> "" And ListView1.ListItems.count = 0 Then
'    Else
     PROC_LISVIEW ListView2
     loc_key = 0
     If ListView2.Visible Then
      loc_key = 1
     End If
 '   End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView2.Visible Then
  Set itmFound = ListView2.FindItem(LTrim(txt_cli.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView2.ListItems.count Then
      ListView2.ListItems.Item(ListView2.ListItems.count).EnsureVisible
   Else
     ListView2.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If


End Sub

Private Sub txtcampo1_GotFocus()
'Azul txtCampo1, txtCampo1
End Sub

Private Sub txtCampo1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
 Exit Sub
End If
If KeyAscii = 13 Then
  If fzonas.Visible Then
  solozonas.SetFocus
  Exit Sub
  End If
End If
If txtcampo2.Visible Then
 If Not IsDate(txtcampo2) Then
   txtcampo2.Text = Format(txtCampo1.Text, "dd/mm/yyyy")
 End If
 Azul2 txtcampo2, txtcampo2
Else
Pantalla.SetFocus
End If
 

End Sub

Private Sub txtcampo2_GotFocus()
'Azul txtCampo2, txtCampo2
End Sub

Private Sub txtCampo2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
 Exit Sub
End If
If Pantalla.Enabled Then
   Pantalla.SetFocus
End If

End Sub
Public Sub IMP_COMISION()
'On Error GoTo FINTODO
Dim RQ As rdoQuery
Dim rs As rdoResultset
Dim WSVALOR As String
Dim ws_clave As String
Dim LETRAS(24) As String * 1
Dim WSFECHA As Date
Dim wcodven As Integer
Dim wvalor
Dim ws_ingresos As Currency
Dim ws_salidas As Currency
Dim val_ingresos As Currency
Dim val_salidas As Currency
Dim acu_val_ingresos As Currency
Dim acu_val_salidas As Currency
Dim Wche As Integer
Dim wkSELECT As String
Dim wsfile As String
Dim WS_CALCULADA  As Currency
Dim WS_DIAS As Integer
Dim WS_COMI As Currency

Dim IdDivision As Integer
Dim PorComision As Double
Dim ValComision As Double
Dim PorCancelado As Double
Dim v_ws  As Double 'por E6G
'Variables para realcionar con cartera
Dim sqlCartera As String

'WSVALOR = InputBox("1= Solo Creditos , 2 = Solo Contados , 3 = Creditos y Contados", "Seleccion de Documentos", "1")
'If WSVALOR = "" Then Exit Sub
'If Val(WSVALOR) >= 1 And Val(WSVALOR) <= 3 Then
'Else
'  MsgBox "Fuera de Rango ", 48, Pub_Titulo
'  Exit Sub
'End If

wsfile = ""
Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
wcodven = val(Left(listven.Text, 4))
'If Val(WSVALOR) = 1 Then
' pub_cadena = "SELECT * FROM CARACU WHERE CAA_CODCIA = ? AND CAA_CODVEN = ? AND CAA_FECHA >= ? AND CAA_FECHA <= ? AND CAA_ESTADO <> 'E'  AND CAA_CONCEPTO <> 'Extorno -' AND CAA_SIGNO_CAR = -1 AND CAA_TIPDOC ='FA' AND (CAA_FBG ='F' OR CAA_FBG ='B' )  ORDER BY CAA_CODCLIE, CAA_FECHA,CAA_NUM_OPER, CAA_SALDO_CAR" 'OR CAA_FBG ='G'
'End If
'If Val(WSVALOR) = 2 Then
' pub_cadena = "SELECT * FROM CARACU WHERE CAA_CODCIA = ? AND CAA_CODVEN = ? AND CAA_FECHA >= ? AND CAA_FECHA <= ? AND CAA_ESTADO <> 'E'  AND CAA_CONCEPTO <> 'Extorno -' AND CAA_SIGNO_CAR = -1 AND CAA_TIPDOC ='CC' AND (CAA_FBG ='F' OR CAA_FBG ='B' )  ORDER BY CAA_CODCLIE, CAA_FECHA,CAA_NUM_OPER, CAA_SALDO_CAR" 'OR CAA_FBG ='G'
'End If
'If Val(WSVALOR) = 3 Then
 pub_cadena = "SELECT * FROM CARACU WHERE CAA_TIPMOV<>97 AND  CAA_CODCIA = ? AND CAA_CODVEN = ? AND CAA_FECHA >= ? AND CAA_FECHA <= ? AND CAA_ESTADO <> 'E'  AND CAA_CONCEPTO <> 'Extorno -' AND CAA_SIGNO_CAR = -1 AND (CAA_TIPDOC ='CC' OR CAA_TIPDOC ='FA' OR CAA_TIPDOC ='LE' ) AND (CAA_FBG ='F' OR CAA_FBG ='B' OR CAA_FBG =' ' ) and caa_cp='C' and CAA_CODTRA <> 1455 ORDER BY CAA_CODCLIE, CAA_FECHA,CAA_NUM_OPER, CAA_SALDO_CAR" 'OR CAA_FBG ='G'
'End If

Set RQ = CN.CreateQuery("", pub_cadena)
RQ.rdoParameters(0) = ""
RQ.rdoParameters(1) = 0
RQ.rdoParameters(2) = LK_FECHA_DIA
RQ.rdoParameters(3) = LK_FECHA_DIA
Set rs = RQ.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
DoEvents
Dim wsFECHA1, wsFECHA2
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If

RQ(0) = LK_CODCIA
RQ(1) = wcodven
RQ(2) = wsFECHA1
RQ(3) = wsFECHA2
rs.Requery
If rs.EOF = True Then
   MsgBox "!!! NO EXISTEN Datos ...", 48, Pub_Titulo
   GoTo CANCELA
End If
ws_clave = PUB_CLAVE
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.max = rs.RowCount
FrmImp2.ProgBar.Value = 0
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
FrmImp2.ProgBar.Visible = True
DoEvents
xl.Worksheets(1).Activate
GoSub LETRAS
xcuenta = 0
xl.Cells(1, 1) = Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia")))
xl.Cells(3, 1) = "'Comisiones del " & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy") & " en " & cmdmoneda.Text
xl.Cells(4, 1) = "Vendedor : " & Trim(listven.Text)

f1 = 6  'Fila Inicial

FrmImp2.lblproceso.Caption = "Procesando . . .  un Momento ."
DoEvents
fila = 0
acu_val_ingresos = 0
acu_val_salidas = 0
Do Until rs.EOF
    ValComision = 0
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   pu_codclie = rs!CAA_CODCLIE
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
      MsgBox "Error en Codigo de cliente, NO EXISTE ...", 48, Pub_Titulo
      GoTo CANCELA
   End If
   
   'Jala del cartera para verificar la moneda mic
    sqlCartera = "select * from cartera where car_codcia='" & LK_CODCIA & "' and car_tipdoc='" & rs("caa_tipdoc") & "' and car_numDOC=" & rs("caa_numDOC")
    Set PSX = CN.CreateQuery("", sqlCartera)
    Set x = PSX.OpenResultset(rdOpenKeyset)
    x.Requery
    If Not x.EOF Then
        If x("car_moneda") <> Left(cmdmoneda.Text, 1) Then
            GoTo OTRITO
        End If
        ' SI ES LETRA Y ESTA CANCELADA O AMORTIZADA SACAR EL NUMSER Y NUMFAC DE CAR_CONCEPTO CREAR ALGORITMO
        
        'ALGORITMO
        
        
        'PU_NUMSER = SERIE ENCONTRADA
        'PU_FBG = FBG ENCONTRADO
        'PU_NUMFAC = NUMFAC ENCONTRADO
        GoTo ENVIAFACART
    End If
   
   ' CALCULO DE COMISION POR ARTICULOS
    PU_NUMSER = rs("CAA_NUMSER")
    PU_FBG = rs("CAA_FBG")
    PU_NUMFAC = rs("CAA_NUMFAC")
ENVIAFACART:
'se agrego las tres var anteriores
    PU_NUMSER = rs("CAA_NUMSER")
    PU_FBG = rs("CAA_FBG")
    PU_NUMFAC = rs("CAA_NUMFAC")
    SQ_OPER = 1
    PU_TIPMOV = 10
    pu_codcia = LK_CODCIA
    LEER_FAR_LLAVE
    If Not far_llave.EOF Then
        If far_llave("far_moneda") <> Left(cmdmoneda.Text, 1) Then GoTo OTRITO
    End If
    f1 = f1 + 1
     While Not far_llave.EOF 'Do
        SQ_OPER = 1
        PUB_KEY = far_llave("FAR_CODART")
        pu_codcia = LK_CODCIA
        LEER_ART_LLAVE
        If Not art_LLAVE.EOF Then
            IdDivision = art_LLAVE("art_familia")
             If IdDivision <> 17 Then
                If rs("caa_tipdoc") = "CC" Then PorComision = art_LLAVE("art_por1") / 100
                If rs("caa_tipdoc") = "FA" Then PorComision = art_LLAVE("art_por2") / 100
                ValComision = ValComision + PorComision * far_llave("far_precio") * (far_llave("far_cantidad")) / (1 + LK_IGV / 100) * (1 - val(far_llave("far_pordesctos")) / 100)
            Else
                'MODIFICADO POR EDIN Y GLENDA
                SQ_OPER = 1
                PUB_TIPREG = 200
                PUB_NUMTAB = 1
                PUB_CODCIA = LK_CODCIA
                LEER_TAB_LLAVE
                '******************
                'If rs("caa_tipdoc") = "CC" Then PorComision = art_LLAVE("art_por3") / 100
                'If rs("caa_tipdoc") = "FA" Then PorComision = art_LLAVE("art_por4") / 100
                
                If val(far_llave("far_pordesctos")) <= val(tab_llave("TAB_NOMCORTO")) Then ' MODIFICADO POR E&G
                    PorComision = art_LLAVE("art_por3") / 100
                    ValComision = ValComision + (PorComision * far_llave("far_precio") * (far_llave("far_cantidad")) / (1 + LK_IGV / 100)) * (1 - val(far_llave("far_pordesctos")) / 100)
                    
                Else
                ' MODIFICADO POR EDIN Y GLENDA
                PUB_NUMTAB = 2
                LEER_TAB_LLAVE
                '*************************
                If val(far_llave("far_pordesctos")) >= val(tab_llave("TAB_NOMLARGO")) Then ' MODF E&g
                    PorComision = art_LLAVE("art_por4") / 100
                    ValComision = ValComision + (PorComision * far_llave("far_precio") * (far_llave("far_cantidad")) / (1 + LK_IGV / 100)) * (1 - val(far_llave("far_pordesctos")) / 100)
                End If
                End If
            End If
        End If
        far_llave.MoveNext
    Wend 'Loop
   'FIN
   xl.Cells(f1, 1) = wcodven
   If rs!CAA_FBG = "F" Then
     xl.Cells(f1, 2) = "FA"
   ElseIf rs!CAA_FBG = "B" Then
     xl.Cells(f1, 2) = "BO"
   ElseIf rs!CAA_FBG = " " Then
     xl.Cells(f1, 2) = "LE"
   End If
   xl.Cells(f1, 3) = "'" & rs!CAa_numser & " - " & Format(rs!CAa_numfac, "000000") & " / " & rs!CAA_TIPDOC
   xl.Cells(f1, 4) = Trim(cli_llave!CLI_NOMBRE)
   SQ_OPER = 1
   pu_cp = "C"
   pu_codclie = cli_llave!cli_codclie
   pu_codcia = LK_CODCIA
   PUB_SERDOC = rs!caa_serdoc
   PUB_NUMDOC = rs!CAA_NUMDOC
   PUB_TIPDOC = rs!CAA_TIPDOC
   LEER_CAR_LLAVE
   If car_llave.EOF Then
    '
    'MsgBox "Documento Extornado.... ", 48, Pub_Titulo
    GoTo OTRITO
   End If
   If val(car_llave!CAR_IMP_INI) <> 0 Then
   PorCancelado = (val(rs!CAA_IMPORTE) * -1) / (val(car_llave!CAR_IMP_INI))
   Else
   PorCancelado = (val(rs!CAA_IMPORTE) * -1) / (val(rs!CAA_IMPORTE))
   End If
   
   xl.Cells(f1, 5) = " " & car_llave!CAR_FECHA_INGR      'apostrofe entre "'" quitado por GTS p Dirome
   xl.Cells(f1, 6) = " " & car_llave!car_fecha_vcto_orig   'apostrofe entre "'" quitado por GTS p Dirome
   xl.Cells(f1, 7) = " " & rs!CAA_FECHA                    'apostrofe entre "'" quitado por GTS p Dirome
   xl.Cells(f1, 8) = val(rs!CAA_IMPORTE) * -1 ' Monto Pagado
   xl.Cells(f1, 9) = val(car_llave!CAR_IMP_INI) 'deuda original
   xl.Cells(f1, 10) = val(car_llave!car_importe) 'saldo actual
   
   'v_comision = Val(ValComision)
   If PUB_TIPDOC = "LE" Then
   xl.Cells(f1, 11) = val(car_llave!CAR_IMP_INI) * 0.014232435
   Else
   xl.Cells(f1, 11) = ValComision 'Val(car_llave!CAR_COMISION) 'comision kardex cambiado  por mic
   End If
      'If Val(car_llave!car_importe) > 0 Then MsgBox "jgkj"
   'WS_CALCULADA = (rs!CAA_IMPORTE * -1 * car_llave!CAR_COMISION) / (car_llave!CAR_IMP_INI) bloq por mic
   WS_CALCULADA = PorCancelado * ValComision 'agre por mic p comi calculada
   'WS_DIAS = DateDiff("d", RS!CAA_FECHA, car_llave!CAR_FECHA_VCTO_ORIG)
   WS_DIAS = DateDiff("d", car_llave!car_fecha_vcto_orig, rs!CAA_FECHA)
   WS_COMI = POR_COMI(WS_DIAS, 444)
   If WS_COMI = -99 Then
      GoTo CANCELA
   End If
   v_ws = val(WS_CALCULADA)
   If PUB_TIPDOC = "LE" Then
   xl.Cells(f1, 12) = val(rs!CAA_IMPORTE) * -0.014232435
   Else
   xl.Cells(f1, 12) = v_ws 'comision calculada
   End If
   xl.Cells(f1, 13) = WS_COMI '% comision
   If PUB_TIPDOC = "LE" Then
   xl.Cells(f1, 14) = val(rs!CAA_IMPORTE) * -0.014232435 * WS_COMI
   Else
   xl.Cells(f1, 14) = WS_CALCULADA * WS_COMI 'comision pagar
   End If
   xl.Cells(f1, 15) = WS_DIAS
OTRITO:
   rs.MoveNext
Loop
  wran1 = "H" & 7
  wran2 = "H" & f1
  wranF = "H" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "I" & 7
  wran2 = "I" & f1
  wranF = "I" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "J" & 7
  wran2 = "J" & f1
  wranF = "J" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "L" & 7
  wran2 = "L" & f1
  wranF = "L" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 
  wran1 = "A" & 7 & ":O" & f1
  xl.APPLICATION.Worksheets("HOJA DE COMISIONES x VENDEDOR").Range(wran1).Sort Key1:=xl.APPLICATION.Worksheets("HOJA DE COMISIONES x VENDEDOR").Range("O7")
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(2, 2) = "INFORME DE COMISIONES x VENDEDOR"
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub


LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"
Return

WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblproceso.Caption = "Abriendo , Archivo Comisiones.xls . . . "
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\OFFICE\Comisiones.xls", 0, True, 4, WPAS, WPAS

Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
  
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.APPLICATION.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
 
End Sub


Public Function POR_COMI(WDIAS As Integer, NroMN As String) As Currency
Dim WINI As Integer
Dim WFIN As Integer
SQ_OPER = 2
PUB_TIPREG = NroMN ' 444 MINORISTA 445 MAYORISTA
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_mayor.EOF Then
 MsgBox "Definir Tablas de Rangos... Verificar ", 48, Pub_Titulo
 POR_COMI = -99
 Exit Function
End If
WINI = 0
WFIN = 0
'If LK_EMP = "HER" Then
  If WDIAS < 0 Then
    WDIAS = 0
  End If
'End If
Do Until tab_mayor.EOF
 WINI = val(tab_mayor!tab_nomcorto)
 WFIN = val(tab_mayor!TAB_CODART)
 If WFIN = 0 And WINI <> 0 Then
   If WDIAS >= WINI Then
    POR_COMI = val(tab_mayor!tab_NOMLARGO)
    Exit Function
   End If
 ElseIf WDIAS >= WINI And WDIAS <= WFIN Then
   POR_COMI = val(tab_mayor!tab_NOMLARGO)
   Exit Function
 End If
 tab_mayor.MoveNext
Loop
'MsgBox "No entro al rango.. Verificar " & wdias, 48, Pub_Titulo
POR_COMI = 0
End Function
Public Sub COMISIONES()
Dim PSCARR As rdoQuery
Dim CARR As rdoResultset
Dim WS_COMISION, WS_IMPORTE As Currency
Dim PSNUC_REPO As rdoQuery
Dim nuc_repo As rdoResultset
Dim wsFECHA1, wsFECHA2
Dim wver
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
'Dim PSCARR As rdoQuery
'Dim CARR As rdoResultset

Screen.MousePointer = 11
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Verificando Datos . . ."
DoEvents
pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_CP = ? AND CLI_NUCLEO = ? AND CLI_ESTADO ='A'  ORDER BY CLI_NOMBRE"
Set PSNUC_REPO = CN.CreateQuery("", pub_cadena)
Set nuc_repo = PSNUC_REPO.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM CARTERA WHERE  CAR_TIPMOV = 10 AND CAR_CODCIA = ? AND CAR_FECHA_INGR >= ? AND CAR_FECHA_INGR <= ? "
Set PSCARR = CN.CreateQuery("", pub_cadena)
Set CARR = PSCARR.OpenResultset(rdOpenKeyset, rdConcurValues)
PSCARR(0) = LK_CODCIA
PSCARR(1) = wsFECHA1
PSCARR(2) = wsFECHA2
CARR.Requery
If CARR.EOF Then
 MsgBox "No existen Datos para Procesar..", 48, Pub_Titulo
 GoTo fin
End If
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.max = CARR.RowCount
FrmImp2.ProgBar.Value = 0
DoEvents
FrmImp2.lblproceso.Visible = True
DoEvents
FrmImp2.lblproceso.Caption = "Procesando Datos . . ."
Screen.MousePointer = 11
FrmImp2.ProgBar.Visible = True
DoEvents
Do Until CARR.EOF
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   SQ_OPER = 1
   PU_TIPMOV = Nulo_Valor0(CARR!CAR_TIPMOV)
   PU_NUMSER = CARR!car_NUMSER
   PU_NUMFAC = CARR!car_NUMFAC
   pu_codcia = CARR!car_codcia
   PU_FBG = Nulo_Valors(CARR!car_FBG)
   LEER_FAR_LLAVE
   WS_COMISION = 0
   pu_codcia = LK_CODCIA
   Do Until far_llave.EOF = True
      PUB_KEY = far_llave!far_codart
      LEER_ART_LLAVE
      WS_IMPORTE = 0
      If far_llave!far_num_precio = "1" Then
         WS_IMPORTE = redondea(art_LLAVE!ART_POR1 * far_llave!far_cantidad * far_llave!FAR_PRECIO / (100 * far_llave!FAR_equiv))
      ElseIf far_llave!far_num_precio = "2" Then
         WS_IMPORTE = redondea(art_LLAVE!ART_POR2 * far_llave!far_cantidad * far_llave!FAR_PRECIO / (100 * far_llave!FAR_equiv))
      ElseIf far_llave!far_num_precio = "3" Then
         WS_IMPORTE = redondea(art_LLAVE!ART_POR3 * far_llave!far_cantidad * far_llave!FAR_PRECIO / (100 * far_llave!FAR_equiv))
      ElseIf far_llave!far_num_precio = "4" Then
         WS_IMPORTE = redondea(art_LLAVE!ART_POR4 * far_llave!far_cantidad * far_llave!FAR_PRECIO / (100 * far_llave!FAR_equiv))
      ElseIf far_llave!far_num_precio = "5" Then
         WS_IMPORTE = redondea(art_LLAVE!ART_POR5 * far_llave!far_cantidad * far_llave!FAR_PRECIO / (100 * far_llave!FAR_equiv))
      Else
'        Stop
      End If
      WS_COMISION = WS_COMISION + WS_IMPORTE
      far_llave.MoveNext
      DoEvents
   Loop
   CARR.Edit
   CARR!CAR_COMISION = WS_COMISION
   CARR.Update
   CARR.MoveNext

 Loop
 FrmImp2.lblproceso.Visible = False
 FrmImp2.ProgBar.Visible = False
 Screen.MousePointer = 0
 FrmImp2.Pantalla.Enabled = True
 FrmImp2.cerrar.Enabled = True
 MsgBox "Proceso Terminado Correctamente ", 48, Pub_Titulo

Exit Sub
fin:
 FrmImp2.lblproceso.Visible = False
 FrmImp2.ProgBar.Visible = False
 Screen.MousePointer = 0
 FrmImp2.Pantalla.Enabled = True
 FrmImp2.cerrar.Enabled = True
Exit Sub
CANCELA:
End Sub

Public Sub LLENADOS(cont As ListBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
'    cont.AddItem " "
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub
Public Sub LLENADOS_COMBO(cont As ComboBox, tip As Integer)
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        tab_mayor.MoveNext
    Loop
End Sub

Public Sub HER_NEGOCIOS()
On Error GoTo FINTODO
Dim wRuta As String
Dim wmonto As Currency
Dim wcodclie As Currency
Dim var_ACUTOT As Currency
Dim var_ACUATE As Currency
Dim var_ACUPED As Currency
Dim wnumfac As Currency
Dim ws_clave As String
Dim WFLAG As String * 1
Dim wflag2 As String * 1
Dim wsFECHA1, wsFECHA2
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = "0"
usu.Requery
Do Until usu.EOF
  If LK_CODUSU = "ADMIN" And Trim(usu!USU_KEY) = "ADMIN" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  If Trim(usu!USU_KEY) = "SUPERVISOR" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  usu.MoveNext
Loop
pub_cadena = "SELECT CLI_CODCLIE,CLI_CODCIA,CLI_GRUPO FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_GRUPO = ? ORDER BY CLI_CODCIA, CLI_CODCLIE"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCLIE = ? AND FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_ESTADO <> 'E' ORDER BY FAR_CODCIA, FAR_FECHA"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

' el PS_REP1(0) ESTA MAS ABAJO
PS_REP01(1) = LK_CODCIA
PS_REP01(2) = 10
PS_REP01(3) = wsFECHA1
PS_REP01(4) = wsFECHA2
FrmImp2.ProgBar.Visible = True
DoEvents
FrmImp2.ProgBar.Value = 0
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Activando Reporte. . . "
DoEvents
SQ_OPER = 2
PUB_TIPREG = 222
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_mayor.EOF Then
  MsgBox "NO existe Tipos de Negocis ..", 48, Pub_Titulo
  GoTo CANCELA
End If
PS_REP02(0) = LK_CODCIA
f1 = 5  'Fila Inicial
WFLAG = ""
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.max = tab_mayor.RowCount
FrmImp2.lblproceso.Caption = "Procesando . . . "
Do Until tab_mayor.EOF
 FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
 PS_REP02(1) = Nulo_Valor0(tab_mayor!TAB_NUMTAB)
 llave_rep02.Requery
 If llave_rep02.EOF Then
 
 Else
    var_ACUTOT = 0
    var_ACUATE = 0
    var_ACUPED = 0
    wnumfac = -1
    Do Until llave_rep02.EOF
      wcodclie = llave_rep02!cli_codclie
      GoSub FAR_RECORRE
      llave_rep02.MoveNext
    Loop
    If wnumfac <> -1 Then
     If Trim(WFLAG) = "" Then
       GoSub WEXCEL
       FrmImp2.lblproceso.Caption = "Procesando . . . "
       DoEvents
       WFLAG = "A"
     End If
     f1 = f1 + 1
     xl.Cells(f1, 1) = f1 - 5
     xl.Cells(f1, 2) = Trim(tab_mayor!tab_NOMLARGO)
     xl.Cells(f1, 3) = Format(var_ACUPED, "###")
     xl.Cells(f1, 4) = Format(var_ACUATE, "###")
     xl.Cells(f1, 5) = Format(var_ACUTOT, "##,##0.000")
   End If
 End If
 tab_mayor.MoveNext
Loop
 If WFLAG <> "A" Then
   FrmImp2.lblproceso.Visible = False
   MsgBox "NO Existe Ventas ...", 48, Pub_Titulo
   GoTo CANCELA
 End If
  xl.Cells(1, 1) = Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia")))
  wranF = "B" & f1 + 1
  xl.Range(wranF) = "TOTAL "
  wran1 = "C" & 6
  wran2 = "C" & f1
  wranF = "C" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "D" & 6
  wran2 = "D" & f1
  wranF = "D" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "E" & 6
  wran2 = "E" & f1
  wranF = "E" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wmonto = val(xl.Range(wranF))
  wran1 = "F" & 6
  wran2 = "F" & f1
  wranF = "F" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wranF = "A" & f1 + 1 & ":F" & f1 + 1
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  For fila = 1 To f1 - 5
    wran1 = "E" & fila + 5
    wranF = "F" & fila + 5
    If wmonto <> 0 Then xl.Range(wranF).Formula = "=(" & wran1 & "* 100) /" & wmonto
  Next fila
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(2, 1) = "LISTA DE ANALISIS DE VENTAS POR TIPO DE NEGOCIO"
  xl.Cells(3, 1) = "'" & LK_FECHA_DIA
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub
FAR_RECORRE:
    PS_REP01(0) = wcodclie
    llave_rep01.Requery
    If llave_rep01.EOF = True Then
       GoTo dale_otro
    End If
    var_ACUATE = var_ACUATE + 1
    wnumfac = llave_rep01!far_numfac
    var_ACUPED = var_ACUPED + 1
    Do Until llave_rep01.EOF
       If wnumfac <> llave_rep01!far_numfac Then
          wnumfac = llave_rep01!far_numfac
          var_ACUPED = var_ACUPED + 1
       End If
       If llave_rep01!FAR_equiv <> 0 And llave_rep01!FAR_DESCTO = 0 Then
        var_ACUTOT = var_ACUTOT + (llave_rep01!FAR_PRECIO * llave_rep01!far_cantidad) / llave_rep01!FAR_equiv
       End If
       llave_rep01.MoveNext
    Loop
     
dale_otro:
Return

WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  wRuta = PUB_RUTA_OTRO
  wsfile1 = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\hernego.xls"
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblproceso.Caption = "Abriendo , Archivo hernego.xls . . . "
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open wsfile1, 0, True, 4, WPAS, WPAS

Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Exit Sub
FINTODO:
 MsgBox Err.Description + " Reintente Nuevamente ..", 48, Pub_Titulo
' Resume Next
 GoTo CANCELA
End Sub

Public Sub HER_NEGOCIOS2()
On Error GoTo FINTODO
Dim WPEDIDO As Integer
Dim wRuta As String
Dim wmonto As Currency
Dim tUm As Currency
Dim tLitros As Currency
Dim wcodclie As Currency
Dim var_ACUTOT As Currency
Dim var_ACUATE As Currency
Dim var_ACUPED As Currency
Dim VAR_TUM As Currency
Dim VAR_TOTALLITROS As Currency

Dim wnumfac As Currency
Dim ws_clave As String
Dim WFLAG As String * 1
Dim wflag2 As String * 1
Dim wsFECHA1, wsFECHA2
Dim xcuenta
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE

pub_cadena = "SELECT CLI_CODCLIE FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_GRUPO = ? "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

'pub_cadena = "SELECT SUM(CAR_IMPORTE)AS DEUDA, CAR_CODCLIE FROM CARTERA WHERE CAR_CODCIA = '02' AND CAR_CP = 'C' AND CAR_TIPDOC <> 'CH' AND CAR_IMPORTE <> 0  GROUP BY CAR_CODCLIE "
'Set PS_CON03 = CN.CreateQuery("", pub_cadena)
'Set llave_con03 = PS_CON03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


'pub_cadena = "SELECT COUNT(FAR_NUMFAC) AS PEDIDOS, SUM(FAR_CANTIDAD * FAR_PRECIO) AS TOTAL   FROM FACART WHERE FAR_CODCLIE = ? AND FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_ESTADO <> 'E' GROUP BY FAR_FBG, FAR_NUMSER,FAR_NUMFAC"
'                                                                   {FACART.FAR_PRECIO} * {FACART.FAR_CANTIDAD}) /{FACART.FAR_EQUIV}- {FACART.FAR_DESCTO}
pub_cadena = "SELECT FAR_CODCLIE, COUNT(FAR_NUMFAC) AS PEDIDOS, SUM((FAR_CANTIDAD * FAR_PRECIO) / FAR_EQUIV - FAR_DESCTO) AS TOTAL  FROM FACART, CLIENTES, TABLAS  WHERE (FAR_CODCLIE = CLI_CODCLIE AND CLI_GRUPO = TAB_NUMTAB AND TAB_TIPREG = 222) AND CLI_GRUPO = ? AND  FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_ESTADO <> 'E' AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "' GROUP BY FAR_CODCLIE,FAR_FBG, FAR_NUMSER,FAR_NUMFAC" 'mic
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = 0
PS_REP01(3) = LK_FECHA_DIA
PS_REP01(4) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT SUM((FAR_CANTIDAD / FAR_EQUIV) * FAR_PRECIO - FAR_DESCTO) AS TOTAL,SUM((FAR_CANTIDAD / FAR_EQUIV)) AS UM ,SUM((FAR_CANTIDAD / FAR_EQUIV)*FAR_LITRO) AS LITROS  FROM FACART WHERE FAR_CODCLIE = ? AND FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_ESTADO <> 'E' AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "' GROUP BY FAR_FBG, FAR_NUMSER,FAR_NUMFAC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
PS_REP02(2) = 0
PS_REP02(3) = LK_FECHA_DIA
PS_REP02(4) = LK_FECHA_DIA
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


' el PS_REP1(0) ESTA MAS ABAJO
PS_REP01(1) = LK_CODCIA
PS_REP01(2) = 10
PS_REP01(3) = wsFECHA1
PS_REP01(4) = wsFECHA2

PS_REP02(1) = LK_CODCIA
PS_REP02(2) = 10
PS_REP02(3) = wsFECHA1
PS_REP02(4) = wsFECHA2

FrmImp2.ProgBar.Visible = True
DoEvents
FrmImp2.ProgBar.Value = 0
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Activando Reporte. . . "
DoEvents
SQ_OPER = 2
PUB_TIPREG = 222
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_mayor.EOF Then
  MsgBox "NO existe Tipos de Negocios ..", 48, Pub_Titulo
  GoTo CANCELA
End If
PS_REP02(0) = LK_CODCIA
f1 = 5  'Fila Inicial
WFLAG = ""
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.max = tab_mayor.RowCount
FrmImp2.lblproceso.Caption = "Procesando . . . "
GoSub WEXCEL
FrmImp2.lblproceso.Caption = "Procesando . . . "
WFLAG = "A"
Do Until tab_mayor.EOF
    FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
    PS_REP01(0) = Nulo_Valor0(tab_mayor!TAB_NUMTAB)
    DoEvents
    var_ACUTOT = 0
    var_ACUATE = 0
    var_ACUPED = 0
    VAR_TUM = 0
    VAR_TOTALLITROS = 0
    llave_rep01.Requery
    If Not llave_rep01.EOF Then
'      xl.Application.Visible = True
       GoSub FAR_RECORRE
       f1 = f1 + 1
       xl.Cells(f1, 1) = f1 - 5
       xl.Cells(f1, 2) = Trim(tab_mayor!tab_NOMLARGO)
       xl.Cells(f1, 3) = Format(var_ACUPED, "###")
       xl.Cells(f1, 4) = Format(var_ACUATE, "###")
       xl.Cells(f1, 5) = Format(var_ACUTOT, "##,##0.000")
       xl.Cells(f1, 7) = Format(VAR_TUM, "##,##0.000")
       xl.Cells(f1, 9) = Format(VAR_TOTALLITROS, "##,##0.000")
       WFLAG = ""
    End If
 tab_mayor.MoveNext
Loop
 If WFLAG = "A" Then
   FrmImp2.lblproceso.Visible = False
   MsgBox "NO Existe Ventas ...", 48, Pub_Titulo
   GoTo CANCELA
 End If
  xl.Cells(1, 1) = Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia")))
  wranF = "B" & f1 + 1
  xl.Range(wranF) = "TOTAL "
  wran1 = "C" & 6
  wran2 = "C" & f1
  wranF = "C" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "D" & 6
  wran2 = "D" & f1
  wranF = "D" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "E" & 6
  wran2 = "E" & f1
  wranF = "E" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wmonto = val(xl.Range(wranF))
  wran1 = "F" & 6
  wran2 = "F" & f1
  wranF = "F" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wranF = "A" & f1 + 1 & ":J" & f1 + 1
  
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  
  wran1 = "G" & 6
  wran2 = "G" & f1
  wranF = "G" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  tUm = val(xl.Range(wranF))
  
  wran1 = "H" & 6
  wran2 = "H" & f1
  wranF = "H" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  
  wran1 = "I" & 6
  wran2 = "I" & f1
  wranF = "I" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  tLitros = val(xl.Range(wranF))
  
  wran1 = "J" & 6
  wran2 = "J" & f1
  wranF = "J" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  
  For fila = 1 To f1 - 5
    wran1 = "E" & fila + 5
    wranF = "F" & fila + 5
    If wmonto <> 0 Then xl.Range(wranF).Formula = "=(" & wran1 & "* 100) /" & wmonto
    '% um
    wran1 = "G" & fila + 5
    wranF = "H" & fila + 5
    If wmonto <> 0 Then xl.Range(wranF).Formula = "=(" & wran1 & "* 100) /" & tUm
    '% litros
    wran1 = "I" & fila + 5
    wranF = "J" & fila + 5
    If wmonto <> 0 Then xl.Range(wranF).Formula = "=(" & wran1 & "* 100) /" & tLitros
    
  Next fila
  
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(2, 1) = "LISTA DE ANALISIS DE VENTAS POR TIPO DE NEGOCIO"
  xl.Cells(3, 1) = "'DEL " & wsFECHA1 & " AL " & wsFECHA1
  If Left(cmdmoneda.Text, 1) = "S" Then
     xl.Cells(4, 5) = "S/."
  Else
     xl.Cells(4, 5) = "US$."
  End If
  'xl.Cells(4, ) = "%"
  xl.Cells(4, 7) = "TOTAL"
  xl.Cells(5, 7) = "U.M."
  
  xl.Cells(4, 8) = "(%)"
  xl.Cells(4, 10) = "(%)"
  
  xl.Cells(4, 9) = "TOTAL"
  xl.Cells(5, 9) = "LITROS"
  
  xl.Cells(1, 9) = Date
  xl.Cells(2, 9) = Time
  
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub
FAR_RECORRE:
'llave_rep01.Requery
    wnumfac = val(llave_rep01!pedidos)
    wnumfac = -1
    WPEDIDO = 0
    xcuenta = 0
    WPEDIDO = val(llave_rep01.RowCount)
    Do Until llave_rep01.EOF
      If wnumfac <> llave_rep01!far_codclie Then
        xcuenta = xcuenta + 1
        wnumfac = llave_rep01!far_codclie
        PS_REP02(0) = llave_rep01!far_codclie
        llave_rep02.Requery
        Do Until llave_rep02.EOF
           var_ACUTOT = var_ACUTOT + val(llave_rep02!Total)
           VAR_TUM = VAR_TUM + val(llave_rep02!UM)
           VAR_TOTALLITROS = VAR_TOTALLITROS + val(llave_rep02!LITROS)
           
           llave_rep02.MoveNext
        Loop
      End If
      llave_rep01.MoveNext
   Loop
   var_ACUATE = xcuenta
   var_ACUPED = WPEDIDO
Return

WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  wRuta = PUB_RUTA_OTRO
  wsfile1 = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\hernego.xls"
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblproceso.Caption = "Abriendo , Archivo hernego.xls . . . "
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open wsfile1, 0, True, 4, WPAS, WPAS

Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Exit Sub
FINTODO:
 MsgBox Err.Description + " Reintente Nuevamente ..", 48, Pub_Titulo
 
 Resume Next
 GoTo CANCELA
End Sub

Public Sub HSTOCK()
Dim CADENITA, wformula, wformula1, wformula2, wformula3, wformula4
Dim Modo, Modo1
Dim Wche, wkSELECT
Dim wfecha, wfiltra1
Dim wcodcia As String
lblproceso.Visible = True
Pantalla.Enabled = False
cerrar.Enabled = False
If retra_llave!tra_rep1 = "1" Then
  If LK_EMP_PTO = "A" Then
    Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\PTOVTA\" & "hstock.rpt"
    wcodcia = "00"
  Else
    Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hstock.rpt"
    wcodcia = LK_CODCIA
  End If
Else
   Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hstock.rpt"
   wcodcia = LK_CODCIA
End If
Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  " & Trim(retra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
DoEvents
ProgBar.Min = 0
ProgBar.max = 7
ProgBar.Value = 0
ProgBar.Visible = True
ProgBar.Value = ProgBar.Value + 1
wfecha = Format(LK_FECHA_DIA, "dd/mm/yyyy")
CADENITA = ""
If Trim(CmbCalidad.Text) <> "" Then
 CADENITA = "{ARTI.ART_CALIDAD} = " & Trim(Right(CmbCalidad.Text, 4)) & " AND "
End If
pub_cadena = CADENITA + "{ARTI.ART_CODCIA} = '" & wcodcia & "' and {TABLAS.TAB_TIPREG} = 122 and {PRECIOS.PRE_FLAG_UNIDAD} = 'A' and "
CADENITA = ""
wfiltra1 = ""
Modo1 = ""
Wche = 0
Modo1 = "{ARTI.ART_FAMILIA} in [" ''F']"
For fila = 0 To fami.ListCount - 1
  fami.ListIndex = fila
  If fami.Selected(fila) Then
    Wche = 1
    wkSELECT = str(val(Right(fami.Text, 6)))
    wfiltra1 = wfiltra1 + wkSELECT + ","
    Modo1 = Modo1 + wkSELECT + ","
  End If
Next fila
If Wche <> 0 Then
  CADENITA = Left(Modo1, Len(Modo1) - 1) & "] "
  wfiltra1 = Left(wfiltra1, Len(wfiltra1) - 1)
Else
  ProgBar.Visible = False
  lblproceso.Visible = False
  MsgBox "Seleccionar Familia ", 48, Pub_Titulo
  GoTo Cancel
 ' ¡Exit Sub
End If
pub_cadena = pub_cadena + CADENITA
ProgBar.Value = ProgBar.Value + 1
Reportes.Formulas(0) = ""
Reportes.Formulas(1) = ""
Reportes.Formulas(2) = ""
Reportes.Formulas(3) = ""
ProgBar.Value = ProgBar.Value + 1
DoEvents
'wformula1 = "FECHA=  '" & wFecha & "'"
wformula1 = "TITULO=  'LISTADO DE STOCK PARA TOMA DE INVENTARIO '"
wformula2 = "CIA=  '" & Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))) & "'"
wformula3 = "LINEAS=  '" & wfiltra1 & "'"
wformula4 = "CALIDAD=  '" & Trim(Left(CmbCalidad.Text, 40)) & "'"

ProgBar.Value = ProgBar.Value + 1
ProgBar.Value = ProgBar.Value + 1
Reportes.Formulas(0) = wformula1
Reportes.Formulas(1) = wformula2
Reportes.Formulas(2) = wformula3
Reportes.Formulas(3) = wformula4
Reportes.SelectionFormula = pub_cadena
Reportes.WindowTitle = Reportes.WindowTitle & " [ " & Trim(Reportes.ReportFileName) & "]"
Reportes.Action = 1
ProgBar.Value = ProgBar.Value + 1
ProgBar.Value = ProgBar.Value + 1
lblproceso.Visible = False
Pantalla.Enabled = True
cerrar.Enabled = True
ProgBar.Visible = False
Exit Sub
procancela:
MsgBox Err.Description, 48, Pub_Titulo

Exit Sub
Cancel:
ProgBar.Visible = False
lblproceso.Visible = False
Pantalla.Enabled = True
cerrar.Enabled = True

End Sub
Public Sub HER_VENDEDOR()
Dim CADENITA, wformula, wformula1, wformula2, wformula3, wformula4
Dim Modo, Modo1
Dim Wche, wkSELECT, DIA, MES, ANO, DIA1, MES1, ANO1
Dim wfecha, wfiltra1
Dim wsFECHA1 As String
Dim wsFECHA2 As String
pub_cadena = ""
lblproceso.Visible = True
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 GoTo SALE
End If

Wche = 0
Modo1 = "{ARTI.ART_FAMILIA} in [" ''F']"
For fila = 0 To fami.ListCount - 1
  fami.ListIndex = fila
  If fami.Selected(fila) Then
    Wche = 1
    Modo1 = Modo1 + Trim(Right(fami.Text, 4)) + ","
  End If
Next fila
If Wche <> 0 Then
 pub_cadena = Left(Modo1, Len(Modo1) - 1) & "] AND "
Else
 pub_cadena = ""
End If
Wche = 0
Modo1 = "{FACART.FAR_CODVEN} in [" ''F']"
For fila = 0 To vendmulti.ListCount - 1
  vendmulti.ListIndex = fila
  If vendmulti.Selected(fila) Then
    Wche = 1
    Modo1 = Modo1 + str(val(Left(vendmulti.Text, 3))) + ","
  End If
Next fila
If Wche <> 0 Then
 pub_cadena = pub_cadena + Left(Modo1, Len(Modo1) - 1) & "] AND "
Else
pub_cadena = ""
End If

DIA = Day(wsFECHA1)
MES = Month(wsFECHA1)
ANO = Year(wsFECHA1)
DIA1 = Day(wsFECHA2)
MES1 = Month(wsFECHA2)
ANO1 = Year(wsFECHA2)
pub_cadena = pub_cadena + "{FACART.FAR_MONEDA} = '" & Left(cmdmoneda.Text, 1) & "' AND {FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} in [10, 97] AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_FECHA} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {FACART.FAR_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  " & Trim(retra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
DoEvents
ProgBar.Min = 0
ProgBar.max = 7
ProgBar.Value = 0
ProgBar.Visible = True
ProgBar.Value = ProgBar.Value + 1
wsFECHA1 = txtCampo1.Text
wsFECHA2 = txtcampo2.Text
wfecha = "DEL " & wsFECHA1 & " AL " & wsFECHA2
ProgBar.Value = ProgBar.Value + 1
Reportes.Formulas(0) = ""
Reportes.Formulas(1) = ""
Reportes.Formulas(2) = ""
Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hvend.rpt"
ProgBar.Value = ProgBar.Value + 1
DoEvents
wformula1 = "TITULO=  'VENTAS ACUMULADAS x VENDEDOR '"
wformula2 = "CIA=  '" & Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))) & "'"
wformula3 = "LINEAS=  '" & wfecha & "'"
ProgBar.Value = ProgBar.Value + 1
ProgBar.Value = ProgBar.Value + 1
Reportes.Formulas(0) = wformula1
Reportes.Formulas(1) = wformula2
Reportes.Formulas(2) = wformula3
Reportes.SelectionFormula = pub_cadena
Reportes.WindowTitle = Reportes.WindowTitle & " [ " & Trim(Reportes.ReportFileName) & "]"
On Error GoTo SALE
Reportes.Action = 1
On Error GoTo 0
ProgBar.Value = ProgBar.Value + 1
ProgBar.Value = ProgBar.Value + 1
lblproceso.Visible = False
Pantalla.Enabled = True
cerrar.Enabled = True
ProgBar.Visible = False
Exit Sub
procancela:
MsgBox Err.Description, 48, Pub_Titulo

Exit Sub
Cancel:
ProgBar.Visible = False
lblproceso.Visible = False
Pantalla.Enabled = True
cerrar.Enabled = True
Exit Sub
SALE:
 If Err.Number = 20504 Then
   MsgBox "No esta el informe : " & Reportes.ReportFileName, 48, Pub_Titulo
 Else
   MsgBox "Fechas Invalidas .. intente nuevamente .. ", 48, Pub_Titulo
  
 End If
 Azul2 txtCampo1, txtCampo1
 lblproceso.Visible = False
 Pantalla.Enabled = True
 cerrar.Enabled = True
 ProgBar.Visible = False

End Sub
Public Sub HER_REGISTRO()
Dim CADENITA, wformula, wformula1, wformula2, wformula3, wformula4
Dim Modo, Modo1
Dim Wche, wkSELECT, DIA, MES, ANO, DIA1, MES1, ANO1
Dim wfecha, wfiltra1
Dim wsFECHA1 As String
Dim wsFECHA2 As String
pub_cadena = ""
lblproceso.Visible = True
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 GoTo SALE
End If

DIA = Day(wsFECHA1)
MES = Month(wsFECHA1)
ANO = Year(wsFECHA1)

DIA1 = Day(wsFECHA2)
MES1 = Month(wsFECHA2)
ANO1 = Year(wsFECHA2)


pub_cadena = "{FACART.FAR_FBG} <> ' '  AND {FACART.FAR_CP} = 'C' AND {FACART.FAR_TIPMOV} = 10 AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_FECHA} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {FACART.FAR_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  " & Trim(retra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
DoEvents
ProgBar.Min = 0
ProgBar.max = 7
ProgBar.Value = 0
ProgBar.Visible = True
ProgBar.Value = ProgBar.Value + 1
wsFECHA1 = txtCampo1.Text
wsFECHA2 = txtcampo2.Text
wfecha = "DEL " & wsFECHA1 & " AL " & wsFECHA2
ProgBar.Value = ProgBar.Value + 1
Reportes.Formulas(0) = ""
Reportes.Formulas(1) = ""
Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "regvent.rpt"
ProgBar.Value = ProgBar.Value + 1
DoEvents
wformula1 = "TITULO=  'REGISTRO DE VENTAS DIARIAS '"
wformula2 = "CIA=  '" & Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))) & "'"
ProgBar.Value = ProgBar.Value + 1
ProgBar.Value = ProgBar.Value + 1
Reportes.Formulas(0) = wformula1
Reportes.Formulas(1) = wformula2
Reportes.SelectionFormula = pub_cadena
Reportes.WindowTitle = Reportes.WindowTitle & " [ " & Trim(Reportes.ReportFileName) & "]"
On Error GoTo SALE
Reportes.Action = 1
On Error GoTo 0
ProgBar.Value = ProgBar.Value + 1
ProgBar.Value = ProgBar.Value + 1
lblproceso.Visible = False
Pantalla.Enabled = True
cerrar.Enabled = True
ProgBar.Visible = False
Exit Sub
procancela:
MsgBox Err.Description, 48, Pub_Titulo

Exit Sub
Cancel:
ProgBar.Visible = False
lblproceso.Visible = False
Pantalla.Enabled = True
cerrar.Enabled = True
Exit Sub
SALE:
 MsgBox "Fechas Invalidas .. intente nuevamente .. ", 48, Pub_Titulo
 Azul2 txtCampo1, txtCampo1
 lblproceso.Visible = False
 Pantalla.Enabled = True
 cerrar.Enabled = True
 ProgBar.Visible = False

End Sub

Private Sub txtDias1_KeyPress(Index As Integer, KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 And Index = 0 Then
 Azul txtDias1(1), txtDias1(1)
ElseIf KeyAscii = 13 And Index = 1 Then
 Azul txtDias2(0), txtDias2(0)
End If
End Sub

Private Sub txtDias2_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 0 Then
SOLO_ENTERO KeyAscii
Else
 Dim car
 car = Chr$(KeyAscii)
 car = UCase$(Chr$(KeyAscii))
 KeyAscii = Asc(car)
 If car < "0" Or car > "9" Then
     If KeyAscii <> 8 And KeyAscii <> 13 And car <> "+" Then
         KeyAscii = 0
         Beep
     End If
 End If
End If
If KeyAscii = 13 And Index = 0 Then
 Azul txtDias2(1), txtDias2(1)
ElseIf KeyAscii = 13 And Index = 1 Then
 Pantalla.SetFocus
End If
End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Pantalla.Enabled Then Pantalla.SetFocus
End If

End Sub

Private Sub vendmulti_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If txtCampo1.Visible Then
   Azul2 txtCampo1, txtCampo1
 End If
 If txtDias1(0).Visible And txtDias1(0).Enabled Then
  Azul txtDias1(0), txtDias1(0)
 End If
End If
End Sub
Public Sub VEN_NEGOCIOS()
On Error GoTo FINTODO

Dim LETRAS(64) As String * 2
Dim Wche As Integer
Dim wRuta As String
Dim wmonto As Currency
Dim wcodclie As Currency
Dim var_ACUTOT As Currency
Dim var_ACUATE As Currency
Dim var_ACUPED As Currency
Dim wnumfac As Currency
Dim ws_clave As String
Dim WFLAG As String * 1
Dim wsFECHA1, wsFECHA2
Dim Acodven() As Currency
Dim Modo1 As String
Dim wcodven As Currency
Dim xcuenta As Integer
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
GoSub WEXCEL
pub_cadena = ""
xcuenta = 0

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE

pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = '" & LK_CODCIA & "'"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM FACART WHERE FAR_CODVEN = ? AND FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_ESTADO <> 'E' AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "'ORDER BY FAR_CODCIA, FAR_FECHA, FAR_NUMSER, FAR_NUMFAC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
PS_REP02(2) = 0
PS_REP02(3) = LK_FECHA_DIA
PS_REP02(4) = LK_FECHA_DIA
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


' el PS_REP1(0) ESTA MAS ABAJO
PS_REP02(1) = LK_CODCIA
PS_REP02(2) = 10
PS_REP02(3) = wsFECHA1
PS_REP02(4) = wsFECHA2
DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep01.Requery
If llave_rep01.EOF Then
  GoTo CANCELA
End If
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.max = llave_rep01.RowCount
f1 = 5  'Fila Inicial
WFLAG = ""
SQ_OPER = 2
PUB_TIPREG = 222
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_mayor.EOF Then
   MsgBox "NO existe Tipos de Negocis ..", 48, Pub_Titulo
   GoTo CANCELA
End If
SQ_OPER = 1
pu_cp = "C"
pu_codcia = LK_CODCIA
Do Until llave_rep01.EOF
  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  f1 = f1 + 1
  xl.Cells(f1, 1) = Trim(Format(llave_rep01!VEM_CODVEN, "000")) & " " & Trim(llave_rep01!VEM_NOMBRE)
  PS_REP02(0) = llave_rep01!VEM_CODVEN
  llave_rep02.Requery
  If llave_rep02.EOF Then
     GoTo otroVEN
  End If
  c1 = 1
  tab_mayor.MoveFirst
  Do Until tab_mayor.EOF
    c1 = c1 + 1
    If Trim(WFLAG) = "" Then
     xl.Cells(5, c1) = Left(tab_mayor!tab_NOMLARGO, 8)
    End If
    llave_rep02.MoveFirst
    wnumfac = -1
    xcuenta = 0
    Do Until llave_rep02.EOF
      If wnumfac <> llave_rep02!far_numfac Then
        wnumfac = llave_rep02!far_numfac
        pu_codclie = llave_rep02!far_codclie
        LEER_CLI_LLAVE
        If cli_llave.EOF Then
         MsgBox "Intente Nuevamente..", 48, Pub_Titulo
         GoTo CANCELA
        End If
        If cli_llave!CLI_GRUPO = tab_mayor!TAB_NUMTAB Then
            xcuenta = xcuenta + 1
        End If
     End If
     llave_rep02.MoveNext
    Loop
    xl.Cells(f1, c1) = Format(xcuenta, "###")
    tab_mayor.MoveNext
  Loop
  WFLAG = "A"
otroVEN:
 llave_rep01.MoveNext
Loop
  GoSub LETRAS
  xcuenta = c1 + 1
  For fila = 6 To f1
    wranF = Trim(LETRAS(xcuenta)) & fila
    wran1 = Trim(LETRAS(2)) & fila
    wran2 = Trim(LETRAS(c1)) & fila
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  Next
  For fila = 2 To c1 + 1
    wranF = Trim(LETRAS(fila)) & f1 + 1
    wran1 = Trim(LETRAS(fila)) & 6
    wran2 = Trim(LETRAS(fila)) & f1
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  Next
  wranF = "A" & f1 + 1
  xl.Range(wranF) = "Totales :"
  wranF = Trim(LETRAS(c1 + 1)) & 5
  xl.Range(wranF) = "Total"
  wranF = "A" & f1 + 1 & ":" & Trim(LETRAS(c1 + 1)) & f1 + 1
  xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  If Chesup.Value = 1 Then
    FrmImp2.lblproceso.Caption = "Suprimiendo  0 ..."
    fila = 1
    Do Until fila >= c1 + 1
     fila = fila + 1
     wranF = Trim(LETRAS(fila)) & f1 + 1
     If val(xl.Range(wranF).Value) = 0 Then
         xl.Range(wranF).Delete 4
         fila = fila - 1
         c1 = c1 - 1
     End If
     Loop
  End If
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'" & LK_FECHA_DIA
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub

CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblproceso.Caption = "Abriendo , Archivo Comisiones.xls . . . "
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\VENXNEGO.xls", 0, True, 4, WPAS, WPAS
Return

LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"
LETRAS(25) = "AA"
LETRAS(26) = "AB"
LETRAS(27) = "AC"
LETRAS(28) = "AD"
LETRAS(29) = "AE"
LETRAS(30) = "AF"
LETRAS(31) = "AG"
LETRAS(32) = "AH"
LETRAS(33) = "AI"
LETRAS(34) = "AJ"
LETRAS(35) = "AK"
LETRAS(36) = "AL"
LETRAS(37) = "AM"
LETRAS(38) = "AN"
LETRAS(39) = "AO"
LETRAS(40) = "AP"
LETRAS(41) = "AQ"
LETRAS(42) = "AR"
LETRAS(43) = "AS"
LETRAS(44) = "AT"
LETRAS(45) = "AU"
LETRAS(46) = "AV"
LETRAS(47) = "AW"
LETRAS(48) = "AX"
LETRAS(49) = "BA"
LETRAS(50) = "BB"
LETRAS(51) = "BC"
LETRAS(52) = "BD"
LETRAS(53) = "BE"
LETRAS(54) = "BF"
LETRAS(55) = "BG"
LETRAS(56) = "BH"
LETRAS(57) = "BI"
LETRAS(58) = "BJ"
LETRAS(59) = "BK"
LETRAS(60) = "BL"
LETRAS(61) = "BM"
LETRAS(62) = "BN"
LETRAS(63) = "BO"
LETRAS(64) = "BP"

Return

FINTODO:
 MsgBox "Reintente Nuevamente ..", 48, Pub_Titulo
Resume Next
 GoTo CANCELA
End Sub

Public Sub VENTA_X_VEND()
'On Error GoTo FINTODO
Dim LETRAS(48) As String * 2
Dim Wche As Integer
Dim wfami As Integer
Dim wcodclie As Currency
Dim wcodven As Currency
Dim wnumfac As Currency
Dim ws_clave As String
Dim WFLAG As String * 1
Dim wsFECHA1, wsFECHA2
Dim Modo1 As String
Dim xcuenta As Integer
Dim CADENA As String
Dim TOT_VEN As Integer
Dim CANTIDAD_VEN, wcantidad As Currency
Dim wtot_soles As Currency
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
pub_cadena = ""
xcuenta = 0
Wche = 0
Modo1 = "FAR_CODVEN = "
xcuenta = 0
Dim WCOLUNNA() As Integer
ReDim WCOLUNNA(1000)
'Dim WCOLUNNA(vendmulti.ListCount - 1) As Integer
For fila = 0 To vendmulti.ListCount - 1
  vendmulti.ListIndex = fila
  If vendmulti.Selected(fila) Then
    Wche = 1
    Modo1 = Modo1 + str(val(Left(vendmulti.Text, 3))) + " OR FAR_CODVEN = "
    xcuenta = xcuenta + 1
    WCOLUNNA(xcuenta) = val(Left(vendmulti.Text, 3))
  End If
Next fila
If Wche = 0 Then
 MsgBox " Seleccione al menos un Vendedor ", 48, Pub_Titulo
 vendmulti.SetFocus
 GoTo CANCELA
End If
ReDim Preserve WCOLUNNA(xcuenta)
TOT_VEN = xcuenta
If Wche <> 0 Then
 If Trim(Right(Modo1, 2)) = "=" Then
   Modo1 = Left(Modo1, Len(Modo1) - 16)
 End If
 
 pub_cadena = "SELECT * FROM FACART WHERE FAR_CODART = ? AND FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_ESTADO <> 'E' AND (" & Modo1 & ") AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "'ORDER BY FAR_CODCIA, FAR_CODVEN , FAR_FECHA, FAR_NUMSER, FAR_NUMFAC"
End If

Wche = 0
Modo1 = "ART_FAMILIA = "
For fila = 0 To fami.ListCount - 1
  fami.ListIndex = fila
  If fami.Selected(fila) Then
    Wche = 1
    Modo1 = Modo1 + Trim((Right(fami.Text, 5))) + " OR ART_FAMILIA = "
  End If
Next fila
If Wche = 0 Then
 MsgBox " Seleccione al menos una Familia ", 48, Pub_Titulo
 fami.SetFocus
 GoTo CANCELA
End If
If Wche <> 0 Then
 If Trim(Right(Modo1, 2)) = "=" Then
  CADENA = Left(Modo1, Len(Modo1) - 17)
 End If
End If

Pantalla.Enabled = False
cerrar.Enabled = False
GoSub WEXCEL
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
For fila = 1 To TOT_VEN
  xl.Cells(5, fila + 3) = "V-" & Format(WCOLUNNA(fila), "00")
Next fila
ws_clave = PUB_CLAVE

Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
PS_REP02(2) = 0
PS_REP02(3) = wsFECHA1
PS_REP02(4) = wsFECHA2

Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = '" & LK_CODCIA & "' AND (" & CADENA & ") ORDER BY ART_CODCIA, ART_FAMILIA,ART_NUMERO, ART_SUBGRU, ART_NOMBRE"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)

Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT PRE_UNIDAD,PRE_EQUIV FROM PRECIOS WHERE PRE_CODCIA = '" & LK_CODCIA & "' and PRE_CODART = ? AND PRE_FLAG_UNIDAD = 'A' ORDER BY PRE_CODART"
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
PS_REP03(0) = 0
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


' el PS_REP1(0) ESTA MAS ABAJO
PS_REP02(1) = LK_CODCIA
PS_REP02(2) = 10
PS_REP02(3) = wsFECHA1
PS_REP02(4) = wsFECHA2
DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.lblproceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep01.Requery
If llave_rep01.EOF Then
  GoTo CANCELA
End If
WFLAG = ""
FrmImp2.lblproceso.Caption = "Procesando... un Momento ."
DoEvents
FrmImp2.ProgBar.max = llave_rep01.RowCount
f1 = 5  'Fila Inicial
SQ_OPER = 1
PUB_TIPREG = 122
PUB_CODCIA = LK_CODCIA
PUB_NUMTAB = llave_rep01!art_familia
LEER_TAB_LLAVE
If tab_llave.EOF = False Then f1 = f1 + 1: xl.Cells(f1, 1) = ">" & Trim(tab_llave!tab_NOMLARGO)
wfami = llave_rep01!art_familia
Do Until llave_rep01.EOF
  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  If wfami <> llave_rep01!art_familia Then
    PUB_NUMTAB = llave_rep01!art_familia
    LEER_TAB_LLAVE
    If tab_llave.EOF = False Then f1 = f1 + 1: xl.Cells(f1, 1) = ">" & Trim(tab_llave!tab_NOMLARGO)
    wfami = llave_rep01!art_familia
  End If
  wfami = llave_rep01!art_familia
  f1 = f1 + 1
  wtot_soles = 0
  xl.Cells(f1, 1) = Trim(llave_rep01!ART_ALTERNO)
  xl.Cells(f1, 2) = Trim(llave_rep01!art_nombre)
  PS_REP03(0) = llave_rep01!ART_KEY
  llave_rep03.Requery
  If llave_rep03.EOF Then GoTo CANCELA
  xl.Cells(f1, 3) = Trim(llave_rep03!pre_unidad)
  PS_REP02(0) = llave_rep01!ART_KEY
  llave_rep02.Requery
  If llave_rep02.EOF Then
     GoTo otroVEN
  End If
  WFLAG = "A"
  wcodven = llave_rep02!FAR_CODVEN
  xcuenta = 0
  wcantidad = 0
   Do Until llave_rep02.EOF
     If wcodven = llave_rep02!FAR_CODVEN Then
        wcantidad = wcantidad + llave_rep02!far_cantidad / llave_rep03!PRE_EQUIV
        ''If llave_rep02!FAR_EQUIV <> 0 And llave_rep02!FAR_DESCTO = 0 Then MIC
          wtot_soles = wtot_soles + llave_rep02!FAR_PRECIO * (llave_rep02!far_cantidad / llave_rep02!FAR_equiv) - llave_rep02!FAR_DESCTO
        ''End If
        xcuenta = 0
     Else
        CANTIDAD_VEN = wcantidad
        GoSub PONE_CANTIDAD
        wcodven = llave_rep02!FAR_CODVEN
        wcantidad = llave_rep02!far_cantidad / llave_rep03!PRE_EQUIV
        ''If llave_rep02!FAR_EQUIV <> 0 And llave_rep02!FAR_DESCTO = 0 Then MIC
        ''If llave_rep02!FAR_EQUIV <> 0 And llave_rep02!FAR_PRECIO <> 0 Then
          wtot_soles = wtot_soles + llave_rep02!FAR_PRECIO * (llave_rep02!far_cantidad / llave_rep02!FAR_equiv) - llave_rep02!FAR_DESCTO
        ''End If
        xcuenta = 1
     End If
     wcodven = llave_rep02!FAR_CODVEN
     llave_rep02.MoveNext
   Loop
     CANTIDAD_VEN = wcantidad
     GoSub PONE_CANTIDAD
 
otroVEN:
 llave_rep01.MoveNext
Loop
  GoSub LETRAS
'    xl.Application.Visible = True
 If WFLAG <> "A" Then
  MsgBox " No hay Información para el filtro", 48, Pub_Titulo
  GoTo SALTA
 End If
  xcuenta = TOT_VEN + 4
  For fila = 6 To f1
    wranF = Trim(LETRAS(xcuenta)) & fila
    wran1 = Trim(LETRAS(4)) & fila
    wran2 = Trim(LETRAS(TOT_VEN + 3)) & fila
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
    'If Val(xl.Range(wranF).Value) = 0 Then  xl.Range(wranF).Value = ""
  Next
  For fila = 4 To TOT_VEN + 4 + 1
    wranF = Trim(LETRAS(fila)) & f1 + 1
    wran1 = Trim(LETRAS(fila)) & 6
    wran2 = Trim(LETRAS(fila)) & f1
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  Next
  wranF = "A" & f1 + 1
  xl.Range(wranF) = "Totales :"
  wranF = Trim(LETRAS(TOT_VEN + 4)) & 5
  xl.Range(wranF) = "Total"
  wranF = Trim(LETRAS(TOT_VEN + 5)) & 5
  If Left(cmdmoneda.Text, 1) = "S" Then
   xl.Range(wranF) = "S/."
  Else
   xl.Range(wranF) = "US$."
  End If
  wranF = "A" & f1 + 1 & ":" & Trim(LETRAS(TOT_VEN + 5)) & f1 + 1
  xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  If Chesup.Visible And Chesup.Value = 1 Then
    FrmImp2.lblproceso.Caption = "Suprimiendo  0 ..."
    fila = 3
    Do Until fila >= TOT_VEN + 4
     fila = fila + 1
     wranF = Trim(LETRAS(fila)) & f1 + 1

     If val(xl.Range(wranF).Value) = 0 Then
         xl.Range(wranF).Delete 4
         fila = fila - 1
         TOT_VEN = TOT_VEN - 1
     End If
    Loop
' xl.Application.Visible = True
    For fila = 7 To f1
     wranF = Trim(LETRAS(TOT_VEN + 5)) & fila
     If val(xl.Range(wranF).Value) = 0 And Trim(xl.Cells(fila, 2)) <> "" Then
         xl.Range(wranF).Delete 3
         fila = fila - 1
         'TOT_VEN = TOT_VEN - 1
     End If
    Next fila
  End If
SALTA:
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'" & "DEL " & wsFECHA1 & " AL " & wsFECHA2
  xl.DisplayAlerts = False
  'xl.Worksheets(1).Protect ws_clave  'gts quitar proteccion
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub
PONE_CANTIDAD:
For fila = 1 To TOT_VEN
  If WCOLUNNA(fila) = wcodven Then
    xl.Cells(f1, fila + 3) = Format(CANTIDAD_VEN, "####")
    Exit For
  End If
Next fila
xl.Cells(f1, TOT_VEN + 5) = Format(wtot_soles, "0.000")


Return

CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblproceso.Caption = "Abriendo , Archivo Comisiones.xls . . . "
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\VENTAXVEN.xls", 0, True, 4, WPAS, WPAS
Return

LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"
LETRAS(25) = "AA"
LETRAS(26) = "AB"
LETRAS(27) = "AC"
LETRAS(28) = "AD"
LETRAS(29) = "AE"
LETRAS(30) = "AF"
LETRAS(31) = "AG"
LETRAS(32) = "AH"
LETRAS(33) = "AI"
LETRAS(34) = "AJ"
LETRAS(35) = "AK"
LETRAS(36) = "AL"
LETRAS(37) = "AM"
LETRAS(38) = "AN"
LETRAS(39) = "AO"
LETRAS(40) = "AP"
LETRAS(41) = "AQ"
LETRAS(42) = "AR"
LETRAS(43) = "AS"
LETRAS(44) = "AT"
LETRAS(45) = "AU"
LETRAS(46) = "AV"
LETRAS(47) = "AW"
LETRAS(48) = "AX"

Return

FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.APPLICATION.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
 
End Sub
Public Sub VEN_NEGOCIOS_CLI()
On Error GoTo FINTODO
Dim LETRAS(64) As String * 2
Dim Wche As Integer
Dim wRuta As String
Dim wmonto As Currency
Dim wcodclie As Currency
Dim var_ACUTOT As Currency
Dim var_ACUATE As Currency
Dim var_ACUPED As Currency
Dim WCODCLI As Currency
Dim ws_clave As String
Dim WFLAG As String * 1
Dim wsFECHA1, wsFECHA2
Dim Acodven() As Currency
Dim Modo1 As String
Dim wcodven As Currency
Dim xcuenta As Integer
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
GoSub WEXCEL
pub_cadena = ""
xcuenta = 0
Wche = 0
Modo1 = "VEM_CODVEN = "
For fila = 0 To vendmulti.ListCount - 1
  vendmulti.ListIndex = fila
  If vendmulti.Selected(fila) Then
    Wche = 1
    Modo1 = Modo1 + str(val(Left(vendmulti.Text, 3))) + " OR VEM_CODVEN = "
  End If
Next fila
If Wche <> 0 Then
 If Trim(Right(Modo1, 2)) = "=" Then
   Modo1 = Left(Modo1, Len(Modo1) - 16)
 End If
 pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = '" & LK_CODCIA & "' AND (" & Modo1 & ")"
Else
 pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = '" & LK_CODCIA & "'"
End If

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = "0"
usu.Requery
Do Until usu.EOF
  If LK_CODUSU = "ADMIN" And Trim(usu!USU_KEY) = "ADMIN" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  If Trim(usu!USU_KEY) = "SUPERVISOR" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  usu.MoveNext
Loop

Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM FACART WHERE FAR_CODVEN = ? AND FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_ESTADO <> 'E' ORDER BY FAR_CODCIA, FAR_CODCLIE " ', FAR_FECHA, FAR_NUMSER, FAR_NUMFAC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
PS_REP02(2) = 0
PS_REP02(3) = LK_FECHA_DIA
PS_REP02(4) = LK_FECHA_DIA

Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

' el PS_REP1(0) ESTA MAS ABAJO
PS_REP02(1) = LK_CODCIA
PS_REP02(2) = 10
PS_REP02(3) = wsFECHA1
PS_REP02(4) = wsFECHA2
DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.lblproceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep01.Requery
If llave_rep01.EOF Then
  GoTo CANCELA
End If
FrmImp2.lblproceso.Caption = "Procesando... un Momento ."
DoEvents
FrmImp2.ProgBar.max = llave_rep01.RowCount
f1 = 5  'Fila Inicial
WFLAG = ""
SQ_OPER = 2
PUB_TIPREG = 222
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_mayor.EOF Then
   MsgBox "NO existe Tipos de Negocis ..", 48, Pub_Titulo
   GoTo CANCELA
End If
SQ_OPER = 1
pu_cp = "C"
pu_codcia = LK_CODCIA
Do Until llave_rep01.EOF
  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  f1 = f1 + 1
  xl.Cells(f1, 1) = Trim(llave_rep01!VEM_NOMBRE)
  PS_REP02(0) = llave_rep01!VEM_CODVEN
  llave_rep02.Requery
  If llave_rep02.EOF Then
     GoTo otroVEN
  End If
  c1 = 1
  tab_mayor.MoveFirst
  Do Until tab_mayor.EOF
    c1 = c1 + 1
    If Trim(WFLAG) = "" Then xl.Cells(5, c1) = Left(tab_mayor!tab_NOMLARGO, 8)
    llave_rep02.MoveFirst
    WCODCLI = -1
    xcuenta = 0
    Do Until llave_rep02.EOF
     If WCODCLI <> llave_rep02!far_codclie Then
        WCODCLI = llave_rep02!far_codclie
        pu_codclie = llave_rep02!far_codclie
        LEER_CLI_LLAVE
        If cli_llave.EOF Then MsgBox "Intente Nuevamente..", 48, Pub_Titulo: GoTo CANCELA
        If cli_llave!CLI_GRUPO = tab_mayor!TAB_NUMTAB Then
         xcuenta = xcuenta + 1
        End If
     End If
     llave_rep02.MoveNext
    Loop
    xl.Cells(f1, c1) = Format(xcuenta, "###")
    tab_mayor.MoveNext
  Loop
  WFLAG = "A"
otroVEN:
 llave_rep01.MoveNext
Loop
  GoSub LETRAS
  xcuenta = c1 + 1
  For fila = 6 To f1
    wranF = Trim(LETRAS(xcuenta)) & fila
    wran1 = Trim(LETRAS(2)) & fila
    wran2 = Trim(LETRAS(c1)) & fila
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  Next
  For fila = 2 To c1 + 1
    wranF = Trim(LETRAS(fila)) & f1 + 1
    wran1 = Trim(LETRAS(fila)) & 6
    wran2 = Trim(LETRAS(fila)) & f1
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  Next
  wranF = "A" & f1 + 1
  xl.Range(wranF) = "Totales :"
  wranF = Trim(LETRAS(c1 + 1)) & 5
  xl.Range(wranF) = "Total"
  wranF = "A" & f1 + 1 & ":" & Trim(LETRAS(c1 + 4)) & f1 + 1
  xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  If Chesup.Value = 1 Then
    FrmImp2.lblproceso.Caption = "Suprimiendo  0 ..."
    fila = 1
    Do Until fila >= c1 + 1
     fila = fila + 1
     wranF = Trim(LETRAS(fila)) & f1 + 1
     If val(xl.Range(wranF).Value) = 0 Then
         xl.Range(wranF).Delete 4
         fila = fila - 1
         c1 = c1 - 1
     End If
     Loop
  End If
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'" & "DEL " & wsFECHA1 & " AL " & wsFECHA2
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub

CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblproceso.Caption = "Abriendo , Archivo Comisiones.xls . . . "
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\VENXNEGO.xls", 0, True, 4, WPAS, WPAS
Return

LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"
LETRAS(25) = "AA"
LETRAS(26) = "AB"
LETRAS(27) = "AC"
LETRAS(28) = "AD"
LETRAS(29) = "AE"
LETRAS(30) = "AF"
LETRAS(31) = "AG"
LETRAS(32) = "AH"
LETRAS(33) = "AI"
LETRAS(34) = "AJ"
LETRAS(35) = "AK"
LETRAS(36) = "AL"
LETRAS(37) = "AM"
LETRAS(38) = "AN"
LETRAS(39) = "AO"
LETRAS(40) = "AP"
LETRAS(41) = "AQ"
LETRAS(42) = "AR"
LETRAS(43) = "AS"
LETRAS(44) = "AT"
LETRAS(45) = "AU"
LETRAS(46) = "AV"
LETRAS(47) = "AW"
LETRAS(48) = "AX"
LETRAS(49) = "BA"
LETRAS(50) = "BB"
LETRAS(51) = "BC"
LETRAS(52) = "BD"
LETRAS(53) = "BE"
LETRAS(54) = "BF"
LETRAS(55) = "BG"
LETRAS(56) = "BH"
LETRAS(57) = "BI"
LETRAS(58) = "BJ"
LETRAS(59) = "BK"
LETRAS(60) = "BL"
LETRAS(61) = "BM"
LETRAS(62) = "BN"
LETRAS(63) = "BO"
LETRAS(64) = "BP"


Return

FINTODO:
 MsgBox "Verificar , Reintente  Nuevamente ..", 48, Pub_Titulo
 Resume Next
 xl.DisplayAlerts = False
 xl.APPLICATION.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
 
End Sub

Public Sub STOCKVA()
On Error GoTo SALE

Dim CADENITA, wformula, wformula1, wformula2, wformula3, wformula4
Dim Modo, Modo1
Dim Wche, wkSELECT
Dim wfecha, wfiltra1
lblproceso.Visible = True
Pantalla.Enabled = False
cerrar.Enabled = False

Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  " & Trim(retra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
DoEvents
ProgBar.Min = 0
ProgBar.max = 7
ProgBar.Value = 0
ProgBar.Visible = True
ProgBar.Value = ProgBar.Value + 1
wfecha = Format(LK_FECHA_DIA, "dd/mm/yyyy")
'{PRECIOS.PRE_FLAG_UNIDAD} = "A" and
'{TABLAS.TAB_TIPREG} = 2 and
'{ARTI.ART_CALIDAD} = 1 and
'{ARTI.ART_CODCIA} = "01"
If Wfile = "STOCKVA" Then
 pub_cadena = "{ARTI.ART_CODCIA} = '" & LK_CODCIA & "' and {TABLAS.TAB_TIPREG} = 2 and {PRECIOS.PRE_FLAG_UNIDAD} = 'A' and {ARTI.ART_CALIDAD} = 1 "
Else
 pub_cadena = "{ARTI.ART_CODCIA} = '" & LK_CODCIA & "' and {TABLAS.TAB_TIPREG} = 2 and {PRECIOS.PRE_FLAG_UNIDAD} = 'A' and {ARTI.ART_CALIDAD} = 1 "
End If
CADENITA = ""
wfiltra1 = ""
Modo1 = ""
Wche = 0
Modo1 = "{ARTI.ART_FAMILIA} in [" ''F']"
For fila = 0 To fami.ListCount - 1
  fami.ListIndex = fila
  If fami.Selected(fila) Then
    Wche = 1
    wkSELECT = str(val(Right(fami.Text, 6)))
    wfiltra1 = wfiltra1 + Left(fami.Text, 8) + ","
    Modo1 = Modo1 + wkSELECT + ","
  End If
Next fila
If Wche <> 0 Then
  CADENITA = " AND " & Left(Modo1, Len(Modo1) - 1) & "] "
  wfiltra1 = Left(wfiltra1, Len(wfiltra1) - 1)
Else
  CADENITA = ""
End If
pub_cadena = pub_cadena + CADENITA
ProgBar.Value = ProgBar.Value + 1
Reportes.Formulas(0) = ""
Reportes.Formulas(1) = ""
Reportes.Formulas(2) = ""
Reportes.Formulas(3) = ""

If Wfile = "STOCKSEM" Then
   If LK_EMP_PTO = "A" Then
     Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\PTOVTA\" & "STOCKSEM.rpt"
   Else
     Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "STOCKSEM.rpt"
   End If
ElseIf Wfile = "STOCKVA" Then
 Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "STOCKVA.rpt"
End If
ProgBar.Value = ProgBar.Value + 1
DoEvents
'wformula1 = "FECHA=  '" & wFecha & "'"
'wformula1 = "TITULO=  'LISTADO DE STOCK PARA TOMA DE INVENTARIO '"
wformula1 = "CIA=  '" & Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))) & "'"
wformula2 = "LINEAS=  '" & wfecha & "  -  " & wfiltra1 & "'"
'wformula4 = "CALIDAD=  '" & Trim(Left(CmbCalidad.text, 40)) & "'"

ProgBar.Value = ProgBar.Value + 1
ProgBar.Value = ProgBar.Value + 1
Reportes.Formulas(0) = wformula1
Reportes.Formulas(1) = wformula2
'Reportes.Formulas(2) = wformula3
'Reportes.Formulas(3) = wformula4
Reportes.SelectionFormula = pub_cadena
Reportes.WindowTitle = Reportes.WindowTitle & " [ " & Trim(Reportes.ReportFileName) & "]"
Reportes.Action = 1
ProgBar.Value = ProgBar.Value + 1
ProgBar.Value = ProgBar.Value + 1
lblproceso.Visible = False
Pantalla.Enabled = True
cerrar.Enabled = True
ProgBar.Visible = False
Exit Sub
SALE:
procancela:
MsgBox Err.Description, 48, Pub_Titulo

Exit Sub
Cancel:
ProgBar.Visible = False
lblproceso.Visible = False
Pantalla.Enabled = True
cerrar.Enabled = True

End Sub
Public Sub DEDVENC()
Dim wv
Dim CADENITA, wformula, wformula1, wformula2, wformula3, wformula4
Dim Modo, Modo1
Dim Wche, wkSELECT
Dim wfecha, wfiltra1
Dim wdia, wmes, wano
Dim DIA, MES, ANO
Dim DIA1, MES1, ANO1
Dim wsFECHA1
Dim wsFECHA2
Dim wfecha1 As String
lblproceso.Visible = True
Pantalla.Enabled = False
cerrar.Enabled = False

Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  " & Trim(retra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
DoEvents
ProgBar.Min = 0
ProgBar.max = 7
ProgBar.Value = 0
ProgBar.Visible = True
ProgBar.Value = ProgBar.Value + 1
wfecha = Format(LK_FECHA_DIA, "dd/mm/yyyy")
wdia = Day(wfecha)
wmes = Month(wfecha)
wano = Year(wfecha)
'DIA1 = Day(wFecha)
'MES1 = Month(wFecha)
'ANO1 = Year(wFecha)
If Check1.Value = 1 Then
  If Right(txtCampo1.Text, 2) = "__" Then
       wsFECHA1 = Left(txtCampo1.Text, 8)
  Else
       wsFECHA1 = Trim(txtCampo1.Text)
  End If
  If Right(txtcampo2.Text, 2) = "__" Then
       wsFECHA2 = Left(txtcampo2.Text, 8)
  Else
       wsFECHA2 = Trim(txtcampo2.Text)
  End If
  If Not IsDate(wsFECHA1) Then
   MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
   GoTo Cancel
  End If
  If Not IsDate(wsFECHA2) Then
   MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
   GoTo Cancel
  End If
  If CDate(wsFECHA1) > CDate(wsFECHA2) Then
   MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
   GoTo Cancel
  End If
  DIA = Day(wsFECHA1)
  MES = Month(wsFECHA1)
  ANO = Year(wsFECHA1)
  DIA1 = Day(wsFECHA2)
  MES1 = Month(wsFECHA2)
  ANO1 = Year(wsFECHA2)
  wfecha1 = "{CARTERA.CAR_FECHA_VCTO} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {CARTERA.CAR_FECHA_VCTO} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
End If

pub_cadena = "{CARTERA.CAR_FECHA_VCTO} = 'C' AND {CLIENTES.CLI_CODCIA} = '" & LK_CODCIA & "' and {CARTERA.CAR_NUMFAC} >= 1.00 and {@FECACTUAL} = {CARTERA.CAR_FECHA_VCTO} "
If Cheop1.Value = 1 And Cheop2.Value = 0 And Cheop3.Value = 0 Then
 pub_cadena = "{CLIENTES.CLI_CP} = 'C' AND {CLIENTES.CLI_CODCIA} = '" & LK_CODCIA & "' and {CARTERA.CAR_NUMFAC} >= 1.00 and {@FECACTUAL} = {CARTERA.CAR_FECHA_VCTO} "
ElseIf Cheop1.Value = 0 And Cheop2.Value = 1 And Cheop3.Value = 0 Then
 pub_cadena = "{CLIENTES.CLI_CP} = 'C' AND {CLIENTES.CLI_CODCIA} = '" & LK_CODCIA & "' and {CARTERA.CAR_NUMFAC} >= 1.00 and {@DIAVENC} > 0.00 "
ElseIf Cheop1.Value = 0 And Cheop2.Value = 0 And Cheop3.Value = 1 Then
 pub_cadena = "{CLIENTES.CLI_CP} = 'C' AND {CLIENTES.CLI_CODCIA} = '" & LK_CODCIA & "' and {CARTERA.CAR_NUMFAC} >= 1.00 and {@DIAVENC} < 0.00 "
ElseIf Cheop1.Value = 1 And Cheop2.Value = 1 And Cheop3.Value = 0 Then
 pub_cadena = "{CLIENTES.CLI_CP} = 'C' AND {CLIENTES.CLI_CODCIA} = '" & LK_CODCIA & "' and {CARTERA.CAR_NUMFAC} >= 1.00 and ({@FECACTUAL} = {CARTERA.CAR_FECHA_VCTO} OR {@DIAVENC} >= 0.00) "
ElseIf Cheop1.Value = 1 And Cheop2.Value = 1 And Cheop3.Value = 1 Then
 pub_cadena = "{CLIENTES.CLI_CP} = 'C' AND {CLIENTES.CLI_CODCIA} = '" & LK_CODCIA & "' and {CARTERA.CAR_NUMFAC} >= 1.00 "
ElseIf Cheop1.Value = 0 And Cheop2.Value = 1 And Cheop3.Value = 1 Then
 pub_cadena = "{CLIENTES.CLI_CP} = 'C' AND {CLIENTES.CLI_CODCIA} = '" & LK_CODCIA & "' and {CARTERA.CAR_NUMFAC} >= 1.00 and ({@FECACTUAL} <> {CARTERA.CAR_FECHA_VCTO} OR {@DIAVENC} <> 0.00) "
ElseIf Cheop1.Value = 1 And Cheop2.Value = 0 And Cheop3.Value = 1 Then
 pub_cadena = "{CLIENTES.CLI_CP} = 'C' AND {CLIENTES.CLI_CODCIA} = '" & LK_CODCIA & "' and {CARTERA.CAR_NUMFAC} >= 1.00 and ({@FECACTUAL} = {CARTERA.CAR_FECHA_VCTO} OR {@DIAVENC} <= 0.00) "
Else
 pub_cadena = "{CLIENTES.CLI_CP} = 'C' AND {CLIENTES.CLI_CODCIA} = '" & LK_CODCIA & "' and {CARTERA.CAR_NUMFAC} >= 1.00 "
End If
wfiltra1 = ""
CADENITA = ""
Wche = 0
Modo1 = "{CARTERA.CAR_CODVEN} in [" ''F']"
For fila = 0 To vendmulti.ListCount - 1
  vendmulti.ListIndex = fila
  If vendmulti.Selected(fila) Then
    Wche = 1
    Modo1 = Modo1 + str(val(Left(vendmulti.Text, 3))) + ","
    wfiltra1 = wfiltra1 + str(val(Left(vendmulti.Text, 3))) + ","
  End If
Next fila
If Wche <> 0 Then
 CADENITA = " AND " + Left(Modo1, Len(Modo1) - 1) & "] "
 wfiltra1 = Left(wfiltra1, Len(wfiltra1) - 1)
Else
 CADENITA = ""
 wfiltra1 = "(*)"
End If
pub_cadena = pub_cadena + CADENITA
If Trim(pub_cadena) <> "" Then
 pub_cadena = pub_cadena + "AND {CARTERA.CAR_MONEDA} = '" & Left(cmdmoneda.Text, 1) & "'"
End If
ProgBar.Value = ProgBar.Value + 1
Reportes.Formulas(0) = ""
Reportes.Formulas(1) = ""
Reportes.Formulas(2) = ""
Reportes.Formulas(3) = ""
Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "DEDVENC.rpt"
ProgBar.Value = ProgBar.Value + 1
DoEvents

wformula1 = "TITULO=  'LISTADO DE DEUDAS DE CLIENTES'"
wformula2 = "CIA=  '" & Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))) & "'"
wformula3 = "CONCEPTO=  '" & wfecha & "  -  Vend." & wfiltra1 & "'"
wformula4 = "FECACTUAL=  Date ( " & wano & "," & wmes & "," & wdia & ")"
ProgBar.Value = ProgBar.Value + 1
ProgBar.Value = ProgBar.Value + 1
Reportes.Formulas(0) = wformula1
Reportes.Formulas(1) = wformula2
Reportes.Formulas(2) = wformula3
Reportes.Formulas(3) = wformula4
dale:
wv = InputBox("1=Solo Creditos , 2 = Solo Contados , 3 = Credito y Contados ......... Nº:", "Filtro", 3)
If val(wv) < 1 Or val(wv) > 3 Then GoTo dale
If val(wv) = 1 Then pub_cadena = pub_cadena + " AND {CARTERA.CAR_TIPDOC} IN ['FA']"
If val(wv) = 2 Then pub_cadena = pub_cadena + " AND {CARTERA.CAR_TIPDOC} IN ['CC']"

If val(txt_cli.Text) <> 0 Then
    pub_cadena = pub_cadena & " AND {CARTERA.CAR_CODCLIE} = " & val(txt_cli.Text)
End If

If Check1.Value = 1 Then
 Reportes.SelectionFormula = pub_cadena + " AND {CARTERA.CAR_IMPORTE} > 0.001 AND " & wfecha1
Else
 Reportes.SelectionFormula = pub_cadena + " AND {CARTERA.CAR_IMPORTE} > 0.001 "
End If
Reportes.WindowTitle = Reportes.WindowTitle & " [ " & Trim(Reportes.ReportFileName) & "]"
Reportes.Action = 1
ProgBar.Value = ProgBar.Value + 1
ProgBar.Value = ProgBar.Value + 1
lblproceso.Visible = False
Pantalla.Enabled = True
cerrar.Enabled = True
ProgBar.Visible = False
Exit Sub
procancela:
MsgBox Err.Description, 48, Pub_Titulo

Exit Sub
Cancel:
ProgBar.Visible = False
lblproceso.Visible = False
Pantalla.Enabled = True
cerrar.Enabled = True

End Sub

Public Sub ZONA_X_NEGO()
'On Error GoTo FINTODO
Dim LETRAS(48) As String * 2
Dim Wche As Integer
Dim wcodtipo As Currency
Dim wcodclie As Currency
Dim ws_clave As String
Dim WFLAG As String * 1
Dim Modo1 As String
Dim xcuenta As Integer
Dim CADENA As String
Dim TOT_COLU As Integer
Dim CANTIDAD_CLI, wcantidad As Currency
Dim wtot As Currency
Dim cod As Integer
Dim wsolos As String * 1
Pantalla.Enabled = False
cerrar.Enabled = False

Dim WCOLUNNA() As String
ReDim WCOLUNNA(9000)
If opzonas(0).Value Then
  cod = 20
ElseIf opzonas(1).Value Then
  cod = 30
ElseIf opzonas(2).Value Then
  cod = 35
End If

SQ_OPER = 2
PUB_TIPREG = 222
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_mayor.EOF Then
 MsgBox "DEfinir Tipo de Negocio", 48, Pub_Titulo
 GoTo CANCELA
End If
xcuenta = 0
Do Until tab_mayor.EOF
 xcuenta = xcuenta + 1
 WCOLUNNA(xcuenta) = tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
 tab_mayor.MoveNext
Loop
ReDim Preserve WCOLUNNA(xcuenta)
TOT_COLU = xcuenta
pub_cadena = ""
xcuenta = 0
Wche = 0
Modo1 = "TAB_NUMTAB = "
xcuenta = 0
For fila = 0 To zonas.ListCount - 1
  zonas.ListIndex = fila
  If zonas.Selected(fila) Then
    Wche = 1
    Modo1 = Modo1 + str(val(Right(zonas.Text, 5))) + " OR TAB_NUMTAB = "
  End If
Next fila
'If Wche = 0 Then
' MsgBox " Seleccione al menos una zona ", 48, Pub_Titulo
' zonas.SetFocus
' GoTo CANCELA
'End If
If Wche <> 0 Then
 If Trim(Right(Modo1, 2)) = "=" Then
   Modo1 = Left(Modo1, Len(Modo1) - 16)
 End If
 pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = '00' AND TAB_TIPREG = " & cod & " AND (" & Modo1 & ") ORDER BY TAB_NUMTAB "
Else
 pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = '00' AND TAB_TIPREG = " & cod & "   ORDER BY TAB_NUMTAB "
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT CLI_CODCLIE, CLI_GRUPO  FROM CLIENTES WHERE CLI_CASA_ZONA = ? AND CLI_CODCIA = '" & LK_CODCIA & "' and CLI_CP = 'C' ORDER BY CLI_GRUPO, CLI_CASA_ZONA"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

Pantalla.Enabled = False
cerrar.Enabled = False
GoSub WEXCEL
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
For fila = 1 To TOT_COLU
  xl.Cells(5, fila + 1) = Left(WCOLUNNA(fila), 8)
Next fila
ws_clave = "0"
usu.Requery
Do Until usu.EOF
  If LK_CODUSU = "ADMIN" And Trim(usu!USU_KEY) = "ADMIN" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  If Trim(usu!USU_KEY) = "SUPERVISOR" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  usu.MoveNext
Loop

DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.lblproceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep01.Requery
If llave_rep01.EOF Then
  GoTo CANCELA
End If
wtot = 0
WFLAG = ""
FrmImp2.lblproceso.Caption = "Procesando... un Momento ."
DoEvents
FrmImp2.ProgBar.max = llave_rep01.RowCount
f1 = 5  'Fila Inicial
pub_cadena = ""
Do Until llave_rep01.EOF
  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  f1 = f1 + 1
  xl.Cells(f1, 1) = Trim(llave_rep01!tab_NOMLARGO)
  PS_REP02(0) = llave_rep01!TAB_NUMTAB
  llave_rep02.Requery
  If llave_rep02.EOF Then
    GoTo OTRO
  End If
  wcodtipo = llave_rep02!CLI_GRUPO
'  If llave_rep02!CLI_GRUPO = 0 Then MsgBox llave_rep02!CLI_CODCLIE
  xcuenta = 0
  wcantidad = 0
  Do Until llave_rep02.EOF
     If wcodtipo = llave_rep02!CLI_GRUPO Then
        wcantidad = wcantidad + 1
     Else
        CANTIDAD_CLI = wcantidad
        GoSub PONE_CANTIDAD
        wcodtipo = llave_rep02!CLI_GRUPO
        wcantidad = 1
     End If
     wcodtipo = llave_rep02!CLI_GRUPO
     wcodclie = llave_rep02!cli_codclie
     llave_rep02.MoveNext
     WFLAG = "A"
   Loop
   CANTIDAD_CLI = wcantidad
   GoSub PONE_CANTIDAD
  
OTRO:
 llave_rep01.MoveNext
Loop
  GoSub LETRAS
 If WFLAG <> "A" Then
  MsgBox " No hay Información para el filtro", 48, Pub_Titulo
  GoTo SALTA
 End If
 If wtot <> 0 Then MsgBox "Existe " & wtot & " Cliente(s) sin Tipo de Negocio, su(s) Codigo(s) son : " & pub_cadena & " . se recomienda relacionarlo.", 48, Pub_Titulo
  xcuenta = TOT_COLU + 2
  For fila = 6 To f1
    wranF = Trim(LETRAS(xcuenta)) & fila
    wran1 = Trim(LETRAS(2)) & fila
    wran2 = Trim(LETRAS(TOT_COLU + 1)) & fila
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
    If val(xl.Range(wranF).Value) = 0 Then xl.Range(wranF).Value = ""
  Next
  For fila = 2 To TOT_COLU + 2
    wranF = Trim(LETRAS(fila)) & f1 + 1
    wran1 = Trim(LETRAS(fila)) & 6
    wran2 = Trim(LETRAS(fila)) & f1
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  Next
  wranF = "A" & f1 + 1
  xl.Range(wranF) = "Totales :"
  wranF = Trim(LETRAS(TOT_COLU + 2)) & 5
  xl.Range(wranF) = "Total"
  wranF = "A5"
  xl.Range(wranF) = Left(Trim(lblZonas.Caption), Len(lblZonas.Caption) - 2)

  wranF = "A" & f1 + 1 & ":" & Trim(LETRAS(TOT_COLU + 2)) & f1 + 1
  xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  If Chesup.Visible And Chesup.Value = 1 Then
    FrmImp2.lblproceso.Caption = "Suprimiendo  0 ..."
    fila = 1
    Do Until fila >= TOT_COLU + 2
     fila = fila + 1
     wranF = Trim(LETRAS(fila)) & f1 + 1
     If val(xl.Range(wranF).Value) = 0 Then
         xl.Range(wranF).Delete 4
         fila = fila - 1
         TOT_COLU = TOT_COLU - 1
     End If
     Loop
  End If
SALTA:
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub

PONE_CANTIDAD:
wsolos = ""
For fila = 1 To TOT_COLU
  If Trim(Right(WCOLUNNA(fila), 6)) = Trim(str(wcodtipo)) Then
    xl.Cells(f1, fila + 1) = CANTIDAD_CLI
    wsolos = "A"
    Exit For
  End If
Next fila
If wsolos <> "A" Then wtot = wtot + 1: pub_cadena = pub_cadena + str(wcodclie) + ", "
Return

CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblproceso.Caption = "Abriendo , Archivo Comisiones.xls . . . "
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\ZONAXNEGO.xls", 0, True, 4, WPAS, WPAS
Return

LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"
LETRAS(25) = "AA"
LETRAS(26) = "AB"
LETRAS(27) = "AC"
LETRAS(28) = "AD"
LETRAS(29) = "AE"
LETRAS(30) = "AF"
LETRAS(31) = "AG"
LETRAS(32) = "AH"
LETRAS(33) = "AI"
LETRAS(34) = "AJ"
LETRAS(35) = "AK"
LETRAS(36) = "AL"
LETRAS(37) = "AM"
LETRAS(38) = "AN"
LETRAS(39) = "AO"
LETRAS(40) = "AP"
LETRAS(41) = "AQ"
LETRAS(42) = "AR"
LETRAS(43) = "AS"
LETRAS(44) = "AT"
LETRAS(45) = "AU"
LETRAS(46) = "AV"
LETRAS(47) = "AW"
LETRAS(48) = "AX"

Return

FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.APPLICATION.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
 
End Sub

Private Sub zonas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Pantalla.Enabled Then Pantalla.SetFocus
End If
End Sub
Public Sub MOROSIDAD(ByVal sTipo As Integer)
Dim CADENITA, wformula, wformula1, wformula2, wformula3, wformula4
Dim wformula5, wformula6, wformula7, wformula8
Dim Modo, Modo1
Dim Wche, wkSELECT
Dim wfecha, wfiltra1
Dim DIA, MES, ANO
Dim valor As Integer
valor = 0
lblproceso.Visible = True
Pantalla.Enabled = False
cerrar.Enabled = False

If val(txtDias1(0).Text) > val(txtDias1(1).Text) Then
 MsgBox "Rango Invalido ...", 48, Pub_Titulo
 Azul txtDias1(0), txtDias1(0)
 GoTo CANCELA
ElseIf val(txtDias1(0).Text) < 0 Or val(txtDias1(1).Text) < 0 Then
 MsgBox "Rango  No Procede ...", 48, Pub_Titulo
 Azul txtDias1(0), txtDias1(0)
 GoTo CANCELA
End If
If Trim(txtDias2(1).Text) <> "+" Then
 If Not IsNumeric(txtDias2(1).Text) Then
   MsgBox "Dato Incorrecto", 48, Pub_Titulo
   Azul txtDias2(1), txtDias2(1)
   GoTo CANCELA
 End If
 If val(txtDias2(0).Text) > val(txtDias2(1).Text) Then
  MsgBox "Rango Invalido ...", 48, Pub_Titulo
  Azul txtDias2(0), txtDias2(0)
  GoTo CANCELA
 ElseIf val(txtDias2(0).Text) < 0 Or val(txtDias2(1).Text) < 0 Then
  MsgBox "Rango  No Procede ...", 48, Pub_Titulo
  Azul txtDias2(0), txtDias2(0)
  GoTo CANCELA
 End If
 valor = val(txtDias2(1).Text)
Else
 If val(txtDias2(0).Text) < 0 Then
  MsgBox "Rango  No Procede ...", 48, Pub_Titulo
  Azul txtDias2(0), txtDias2(0)
  GoTo CANCELA
 End If
 valor = 9999
End If
Wche = 0
pub_cadena = ""
Modo1 = "{VEMAEST.VEM_CODVEN} in [" ''F']"
For fila = 0 To vendmulti.ListCount - 1
  vendmulti.ListIndex = fila
  If vendmulti.Selected(fila) Then
    Wche = 1
    Modo1 = Modo1 + str(val(Left(vendmulti.Text, 3))) + ","
  End If
Next fila
If Wche <> 0 Then
 pub_cadena = pub_cadena + Left(Modo1, Len(Modo1) - 1) & "] "
Else
pub_cadena = ""
End If

Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  " & Trim(retra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
DoEvents
ProgBar.Min = 0
ProgBar.max = 7
ProgBar.Value = 0
ProgBar.Visible = True
ProgBar.Value = ProgBar.Value + 1
wfecha = Format(LK_FECHA_DIA, "dd/mm/yyyy")
DIA = Day(wfecha)
MES = Month(wfecha)
ANO = Year(wfecha)
If pub_cadena = "" Then
 pub_cadena = "{CARTERA.CAR_MONEDA} = '" & Left(cmdmoneda.Text, 1) & "' AND {CARTERA.CAR_IMPORTE} >= 0.001 AND {CARTERA.CAR_TIPDOC} = 'FA' AND {CARTERA.CAR_NUMFAC} >= 0.001 AND {CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "'"
Else
 pub_cadena = pub_cadena + " AND {CARTERA.CAR_MONEDA} = '" & Left(cmdmoneda.Text, 1) & "' AND {CARTERA.CAR_IMPORTE} >= 0.001 AND {CARTERA.CAR_TIPDOC} = 'FA' AND {CARTERA.CAR_NUMFAC} >= 0.001 AND {CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "'"
End If
ProgBar.Value = ProgBar.Value + 1
Reportes.Formulas(0) = ""
Reportes.Formulas(1) = ""
Reportes.Formulas(2) = ""
Reportes.Formulas(3) = ""
Reportes.Formulas(4) = ""
Reportes.Formulas(5) = ""
Reportes.Formulas(6) = ""
Reportes.Formulas(7) = ""
If sTipo = 0 Then
    Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "MOROSIDAD.rpt"
Else
    Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "MOROSIDADCLIE.rpt"
End If
ProgBar.Value = ProgBar.Value + 1
DoEvents
wformula1 = "TITULO=  'MOROSIDAD X VENDEDORES'"
wformula2 = "CIA=  '" & Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))) & "'"
wformula3 = "DIA=  '" & wfecha & "'"
wformula4 = "DIAS1.1=  " & txtDias1(0).Text
wformula5 = "DIAS1.2=  " & txtDias1(1).Text
wformula6 = "DIAS2.1=  " & txtDias2(0).Text
wformula7 = "DIAS2.2=  " & valor

ProgBar.Value = ProgBar.Value + 1
ProgBar.Value = ProgBar.Value + 1
wformula8 = "DIA_FECHA = Date ( " & ANO & "," & MES & "," & DIA & ")"
Reportes.Formulas(0) = wformula1
Reportes.Formulas(1) = wformula2
Reportes.Formulas(2) = wformula3
Reportes.Formulas(3) = wformula4
Reportes.Formulas(4) = wformula5
Reportes.Formulas(5) = wformula6
Reportes.Formulas(6) = wformula7
Reportes.Formulas(7) = wformula8

Reportes.SelectionFormula = pub_cadena
Reportes.WindowTitle = Reportes.WindowTitle & " [ " & Trim(Reportes.ReportFileName) & "]"
Reportes.Action = 1
ProgBar.Value = ProgBar.Value + 1
ProgBar.Value = ProgBar.Value + 1
lblproceso.Visible = False
Pantalla.Enabled = True
cerrar.Enabled = True
ProgBar.Visible = False
Exit Sub
CANCELA:
ProgBar.Visible = False
lblproceso.Visible = False
Pantalla.Enabled = True
cerrar.Enabled = True

End Sub
Public Sub REPO_CAJA_GEN()
Dim ws_conta As Integer
Dim WS_NOMCLI As String * 20
Dim WS_NOMBAN As String * 20
Dim WS_SALDO As Currency
Dim WS_MONEDA As String * 1
Dim ww_moneda
Dim ws_mensaje
Dim WS_FBG, WS_LARGO, ws_nomche, ws_codclie
Dim WS_SALDO_ING As Currency
Dim WS_SALDO_SAL As Currency
Dim todos_cli
Dim wsFECHA1


Dim PSTEMP_MAYOR As rdoQuery
Dim temp_mayor As rdoResultset


Dim temp_llave As rdoResultset

If Right(txtFecha.Text, 2) = "__" Then
     wsFECHA1 = Left(txtFecha.Text, 8)
Else
     wsFECHA1 = Trim(txtFecha.Text)
End If

If Right(txtFecha.Text, 2) = "__" Then
     wsFECHA1 = Left(txtFecha.Text, 8)
Else
     wsFECHA1 = Trim(txtFecha.Text)
End If

If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If


pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ? AND ALL_CODTRA = 5310"
    Set PSTEMP_MAYOR = CN.CreateQuery("", pub_cadena)
    PSTEMP_MAYOR(0) = LK_CODCIA
    PSTEMP_MAYOR(1) = wsFECHA1
    PSTEMP_MAYOR.MaxRows = 1
    Set temp_mayor = PSTEMP_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
    
    PSTEMP_MAYOR(0) = LK_CODCIA
    temp_mayor.Requery
         
   ' If temp_mayor.EOF Then
   ' If LK_CODUSU <> "SUPERVISOR" Then
   ' MsgBox "No ha cerrado caja......favor de ingresar el monto final de caja"
   ' GoTo CANCELA
   ' End If
   ' Else
   ' End If

FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL

xl.Worksheets(1).Activate
WS_SALDO = 0
xcuenta = 0
f1 = 5

WS_MONEDA = Left(cmdmoneda.Text, 1)

If WS_MONEDA = "S" Then
   xl.Cells(4, 2) = "MONEDA:" & " SOLES"
ElseIf WS_MONEDA = "D" Then
   xl.Cells(4, 2) = "MONEDA:" & " DOLLARES"
End If


PUB_FECHA = wsFECHA1
SQ_OPER = 1
pu_codcia = LK_CODCIA
LEER_ALL_LLAVE
If all_llave.EOF Then GoTo VAMOS
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.max = all_llave.RowCount
FrmImp2.ProgBar.Value = 0
f1 = f1 + 1
ws_conta = 0
If WS_MONEDA = "S" Then
     WS_SALDO = all_llave!ALL_IMPORTE
Else
     WS_SALDO = all_llave!ALL_IMPORTE_DOLL
End If
'MsgBox all_llave!all_codtra
xl.Cells(f1, 2) = "Saldo Anterior:"
xl.Cells(f1, 5) = WS_SALDO
all_llave.MoveNext
Do Until all_llave.EOF
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   If all_llave!ALL_SIGNO_CAJA = 0 Then GoTo OTRO
   If all_llave!all_CODCIA <> LK_CODCIA Then GoTo OTRO
   If LK_EMP = "HER" And all_llave!ALL_CODTRA = 2727 And all_llave!ALL_SIGNO_CAR = 1 Then GoTo OTRO
   If all_llave!all_flag_ext = "E" Then GoTo OTRO
   If (all_llave!ALL_SIGNO_CAR = 0 And all_llave!ALL_tipmov = 0) Or (all_llave("all_codtra") = 2725 Or all_llave("all_codtra") = 5360 Or all_llave("all_codtra") = 2770 Or all_llave("all_codtra") = 2774) Then  'agregado gts 5360
      WS_IMPORTE = all_llave!ALL_IMPORTE
   Else
      WS_IMPORTE = all_llave!ALL_IMPORTE_AMORT
   End If
   'bloqueado por mic para diro
'   If all_llave!ALL_SIGNO_CAR <> 0 And all_llave!ALL_tipmov = 0 And LK_EMP = "HER" Then
'      WS_IMPORTE = all_llave!ALL_IMPORTE
'   End If
  
   If Trim(all_llave!ALL_moneda_ccm) <> " " And val(all_llave!all_codban) <> 0 Then
      ww_moneda = all_llave!ALL_moneda_ccm
   ElseIf Trim(all_llave!ALL_MONEDA_CLI) <> " " And val(all_llave!ALL_CODCLIE) <> 0 Then
      ww_moneda = all_llave!ALL_MONEDA_CLI
   ElseIf Trim(all_llave!ALL_MONEDA_CAJA) <> " " Then
      ww_moneda = all_llave!ALL_MONEDA_CAJA
   End If
   If ww_moneda <> WS_MONEDA Then GoTo OTRO
   WS_NOMCLI = ""
   WS_NOMBAN = ""
   ws_mensaje = ""
   WS_FBG = ""
   WS_LARGO = ""
   ws_nomche = ""
   ws_codclie = val(all_llave!ALL_CODCLIE)
   If all_llave!all_numfac <> 0 Then
       If all_llave!ALL_FBG = "F" Then
          WS_FBG = "Fact. " & all_llave!ALL_NUMSER & "-" & all_llave!all_numfac
       ElseIf all_llave!ALL_FBG = "B" Then
          WS_FBG = "Bolet." & all_llave!ALL_NUMSER & "-" & all_llave!all_numfac
       Else
          WS_FBG = "Guia. " & all_llave!ALL_NUMSER & "-" & all_llave!all_numfac
       End If
   End If
   If all_llave!all_chenum <> 0 And all_llave!ALL_SIGNO_CCM = -1 Then
       ws_nomche = "O/Pago: " & all_llave!all_chenum
   End If
      
   If ws_codclie <> 0 Then
         pu_cp = all_llave!ALL_CP
         pu_codcia = LK_CODCIA
         pu_codclie = all_llave!ALL_CODCLIE
         SQ_OPER = 1
         LEER_CLI_LLAVE
         WS_NOMCLI = Left(cli_llave!CLI_NOMBRE, 18) & ":"
   End If
   If val(all_llave!all_codban) <> 0 Then
         pu_codcia = LK_CODCIA
         PUB_CODBAN = all_llave!all_codban
         SQ_OPER = 1
         LEER_CCM_LLAVE
         WS_NOMBAN = ccm_llave!CCM_NOMBRE & ":"
   End If
   If all_llave!ALL_SIGNO_CAR <> 0 Then
       SQ_OPER = 1
       pu_cp = all_llave!ALL_CP
       pu_codclie = ws_codclie
       pu_codcia = LK_CODCIA
       PUB_SERDOC = Nulo_Valor0(all_llave!ALL_serdoc)
       PUB_NUMDOC = all_llave!ALL_NUMDOC
       PUB_TIPDOC = all_llave!ALL_TIPDOC
       LEER_CAR_LLAVE
      ' WS_NOMCLI = Left(WS_NOMCLI, 30) & car_llave!car_FBG & "-." & car_llave!car_NUMSER & "-" & car_llave!car_NUMFAC
   End If
   WS_LARGO = Trim(WS_FBG) & Trim(ws_mensaje) & Trim(WS_NOMBAN) & Trim(ws_nomche) & "-" & Trim(WS_NOMCLI)  '&
  ' If WS_LARGO = "" Then WS_LARGO = all_llave!ALL_CONCEPTO   'QUITADO GTS
      
   f1 = f1 + 1
   If all_llave!ALL_SIGNO_CAJA = 1 Then
      WS_SALDO = WS_SALDO + WS_IMPORTE
      WS_SALDO_ING = WS_SALDO_ING + WS_IMPORTE
      xl.Cells(f1, 3) = WS_IMPORTE
   Else
      WS_SALDO = WS_SALDO - WS_IMPORTE
      WS_SALDO_SAL = WS_SALDO_SAL + WS_IMPORTE
      xl.Cells(f1, 4) = WS_IMPORTE
   End If
   ws_conta = ws_conta + 1
   xl.Cells(f1, 1) = ws_conta
    If Trim(all_llave!ALL_autocon) = "" Then
        xl.Cells(f1, 2) = WS_LARGO + " " + all_llave!all_concepto '
    Else
        xl.Cells(f1, 2) = WS_LARGO + " " + all_llave!all_concepto  'WS_LARGO
    End If
   xl.Cells(f1, 5) = WS_SALDO
   'If Trim(all_llave!ALL_CONCEPTO) <> "" And Trim(WS_LARGO) <> Trim(all_llave!ALL_CONCEPTO) Then
   '   F1 = F1 + 1
   '   xl.Cells(F1, 2) = Trim(all_llave!ALL_CONCEPTO)
   'End If
   
OTRO:
  PUB_NUM_OPER = all_llave!ALL_NUMOPER
  all_llave.MoveNext
Loop
   If LK_FECHA_DIA = wsFECHA1 Then
      If WS_MONEDA = "S" Then
         PUB_TIPREG = 1000
      Else
         PUB_TIPREG = 1001
      End If
      SQ_OPER = 1
      PUB_NUMTAB = 0
      PUB_CODCIA = LK_CODCIA
      LEER_TAB_LLAVE
      If tab_llave.EOF Then
         tab_llave.AddNew
         tab_llave!TAB_NUMTAB = 0
         tab_llave!TAB_CODCIA = LK_CODCIA
         If WS_MONEDA = "S" Then
            tab_llave!TAB_TIPREG = 1000
         Else
            tab_llave!TAB_TIPREG = 1001
         End If
      Else
         tab_llave.Edit
      End If
      tab_llave!tab_NOMLARGO = PUB_NUM_OPER
      tab_llave!tab_nomcorto = Format(LK_FECHA_DIA, "dd/mm/yyyy")
      tab_llave!TAB_contable2 = WS_SALDO
      tab_llave.Update
   End If
   
VAMOS:
wran1 = "C5"
wran2 = "C" & f1
xl.Visible = True
wranF = "C" & f1 + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
wran1 = "D5"
wran2 = "D" & f1
wranF = "D" & f1 + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
wranF = "C" & 5 & ":C" & f1
xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 3
wranF = "D" & 5 & ":D" & f1
xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 3
wranF = "E" & 5 & ":E" & f1
xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 3

FrmImp2.ProgBar.Value = todos_cli
Screen.MousePointer = 0
DoEvents
FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
DoEvents
xcuenta = 1
xl.Cells(2, 2) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
xl.Cells(2, 5) = "'" & Format(wsFECHA1, "dd mmm yyyy")
xl.DisplayAlerts = False
xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
xl.Worksheets("Hoja1").Protect PUB_CLAVE

DoEvents
FrmImp2.lblproceso.Visible = False
FrmImp2.ProgBar.Visible = False
xl.APPLICATION.Visible = True
Set xl = Nothing
Screen.MousePointer = 0
FrmImp2.Pantalla.Enabled = True
FrmImp2.Pantalla.Caption = "Por &Pantalla"
FrmImp2.lblproceso.Visible = False

Exit Sub


WEXCEL:
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FrmImp2.lblproceso.Caption = "Abriendo , Archivo Caja.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\OFFICE\CAJA.xls", 0, True, 4, WPAS, WPAS
Return

Exit Sub

'LLENA_VALOR:
'For I = QJ To WDIF
'  If I >= 3 Then
'    xl.Cells(F1 + 7, I) = Format(LOC_VALOR, "0.000")
'  End If
'Next I
'Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
Exit Sub

End Sub

Public Sub REPO_CAJA_GEN_FECHA()
Dim ws_conta As Integer
Dim WS_NOMCLI As String * 20
Dim WS_NOMBAN As String * 20
Dim WS_SALDO As Currency
Dim WS_MONEDA As String * 1
Dim ww_moneda
Dim ws_mensaje
Dim WS_FBG, WS_LARGO, ws_nomche, ws_codclie
Dim WS_SALDO_ING As Currency
Dim WS_SALDO_SAL As Currency
Dim todos_cli
Dim wsFECHA1
Dim wsFECHA2 'mODIFICADO
'If Right(txtFecha.Text, 2) = "__" Then
 '    wsFECHA1 = Left(txtFecha.Text, 8)
'Else
 '    wsFECHA1 = Trim(txtFecha.Text)
'End If
'Modificado 20042004  --iNICIO
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
'fIN mODIFICADO
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If


FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL

xl.Worksheets(1).Activate
WS_SALDO = 0
xcuenta = 0
f1 = 5

WS_MONEDA = Left(cmdmoneda.Text, 1)

If WS_MONEDA = "S" Then
   xl.Cells(4, 2) = "MONEDA:" & " SOLES"
ElseIf WS_MONEDA = "D" Then
   xl.Cells(4, 2) = "MONEDA:" & " DOLLARES"
End If


PUB_FECHA = wsFECHA1
PUB_FECHA1 = wsFECHA2
SQ_OPER = 4
pu_codcia = LK_CODCIA
LEER_ALL_LLAVE
Set all_llave = all_llave3 'asigna el rs de Reporte de caja a ALL_LLAve 20042004
If all_llave.EOF Then GoTo VAMOS
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.max = all_llave.RowCount
FrmImp2.ProgBar.Value = 0
f1 = f1 + 1
ws_conta = 0
If WS_MONEDA = "S" Then
     WS_SALDO = all_llave!ALL_IMPORTE
Else
     WS_SALDO = all_llave!ALL_IMPORTE_DOLL
End If
'MsgBox all_llave!all_codtra
xl.Cells(f1, 3) = "Saldo Anterior:"
xl.Cells(f1, 6) = WS_SALDO
all_llave.MoveNext
Do Until all_llave.EOF
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   If all_llave!ALL_SIGNO_CAJA = 0 Then GoTo OTRO
   If all_llave!all_CODCIA <> LK_CODCIA Then GoTo OTRO
   If LK_EMP = "HER" And all_llave!ALL_CODTRA = 2727 And all_llave!ALL_SIGNO_CAR = 1 Then GoTo OTRO
   If all_llave!all_flag_ext = "E" Then GoTo OTRO
   If (all_llave!ALL_SIGNO_CAR = 0 And all_llave!ALL_tipmov = 0) Or (all_llave("all_codtra") = 2725) Then
      WS_IMPORTE = all_llave!ALL_IMPORTE
   Else
      WS_IMPORTE = all_llave!ALL_IMPORTE_AMORT
   End If
   'bloqueado por mic para diro
'   If all_llave!ALL_SIGNO_CAR <> 0 And all_llave!ALL_tipmov = 0 And LK_EMP = "HER" Then
'      WS_IMPORTE = all_llave!ALL_IMPORTE
'   End If
  
   If Trim(all_llave!ALL_moneda_ccm) <> " " And val(all_llave!all_codban) <> 0 Then
      ww_moneda = all_llave!ALL_moneda_ccm
   ElseIf Trim(all_llave!ALL_MONEDA_CLI) <> " " And val(all_llave!ALL_CODCLIE) <> 0 Then
      ww_moneda = all_llave!ALL_MONEDA_CLI
   ElseIf Trim(all_llave!ALL_MONEDA_CAJA) <> " " Then
      ww_moneda = all_llave!ALL_MONEDA_CAJA
   End If
   If ww_moneda <> WS_MONEDA Then GoTo OTRO
   WS_NOMCLI = ""
   WS_NOMBAN = ""
   ws_mensaje = ""
   WS_FBG = ""
   WS_LARGO = ""
   ws_nomche = ""
   ws_codclie = val(all_llave!ALL_CODCLIE)
   If all_llave!all_numfac <> 0 Then
       If all_llave!ALL_FBG = "F" Then
          WS_FBG = "Fact. " & all_llave!ALL_NUMSER & "-" & all_llave!all_numfac
       ElseIf all_llave!ALL_FBG = "B" Then
          WS_FBG = "Bolet." & all_llave!ALL_NUMSER & "-" & all_llave!all_numfac
       Else
          WS_FBG = "Guia. " & all_llave!ALL_NUMSER & "-" & all_llave!all_numfac
       End If
   End If
   If all_llave!all_chenum <> 0 And all_llave!ALL_SIGNO_CCM = -1 Then
       ws_nomche = "O/Pago: " & all_llave!all_chenum
   End If
      
   If ws_codclie <> 0 Then
         pu_cp = all_llave!ALL_CP
         pu_codcia = LK_CODCIA
         pu_codclie = all_llave!ALL_CODCLIE
         SQ_OPER = 1
         LEER_CLI_LLAVE
         WS_NOMCLI = Left(cli_llave!CLI_NOMBRE, 18) & ":"
   End If
   If val(all_llave!all_codban) <> 0 Then
         pu_codcia = LK_CODCIA
         PUB_CODBAN = all_llave!all_codban
         SQ_OPER = 1
         LEER_CCM_LLAVE
         WS_NOMBAN = ccm_llave!CCM_NOMBRE & ":"
   End If
   If all_llave!ALL_SIGNO_CAR <> 0 Then
       SQ_OPER = 1
       pu_cp = all_llave!ALL_CP
       pu_codclie = ws_codclie
       pu_codcia = LK_CODCIA
       PUB_SERDOC = Nulo_Valor0(all_llave!ALL_serdoc)
       PUB_NUMDOC = all_llave!ALL_NUMDOC
       PUB_TIPDOC = all_llave!ALL_TIPDOC
       LEER_CAR_LLAVE
       'WS_NOMCLI = Left(WS_NOMCLI, 30) & car_llave!car_FBG & "-." & car_llave!car_NUMSER & "-" & car_llave!car_NUMFAC
   End If
   WS_LARGO = Trim(WS_NOMCLI) & Trim(WS_FBG) & Trim(ws_mensaje) & Trim(WS_NOMBAN) & Trim(ws_nomche)
   If WS_LARGO = "" Then WS_LARGO = all_llave!all_concepto
      
   f1 = f1 + 1
   If all_llave!ALL_SIGNO_CAJA = 1 Then
      WS_SALDO = WS_SALDO + WS_IMPORTE
      WS_SALDO_ING = WS_SALDO_ING + WS_IMPORTE
      xl.Cells(f1, 4) = WS_IMPORTE
   Else
      WS_SALDO = WS_SALDO - WS_IMPORTE
      WS_SALDO_SAL = WS_SALDO_SAL + WS_IMPORTE
      xl.Cells(f1, 5) = WS_IMPORTE
   End If
   ws_conta = ws_conta + 1
   xl.Cells(f1, 1) = ws_conta
   xl.Cells(f1, 2) = all_llave!ALL_FECHA_DIA
    If Trim(all_llave!ALL_autocon) = "" Then
        xl.Cells(f1, 3) = all_llave!all_concepto 'WS_LARGO
    Else
        xl.Cells(f1, 3) = all_llave!ALL_autocon 'WS_LARGO
    End If
   xl.Cells(f1, 6) = WS_SALDO
   'If Trim(all_llave!ALL_CONCEPTO) <> "" And Trim(WS_LARGO) <> Trim(all_llave!ALL_CONCEPTO) Then
   '   F1 = F1 + 1
   '   xl.Cells(F1, 2) = Trim(all_llave!ALL_CONCEPTO)
   'End If
   
OTRO:
  PUB_NUM_OPER = all_llave!ALL_NUMOPER
  all_llave.MoveNext
Loop
   If LK_FECHA_DIA = wsFECHA1 Then
      If WS_MONEDA = "S" Then
         PUB_TIPREG = 1000
      Else
         PUB_TIPREG = 1001
      End If
      SQ_OPER = 1
      PUB_NUMTAB = 0
      PUB_CODCIA = LK_CODCIA
      LEER_TAB_LLAVE
      If tab_llave.EOF Then
         tab_llave.AddNew
         tab_llave!TAB_NUMTAB = 0
         tab_llave!TAB_CODCIA = LK_CODCIA
         If WS_MONEDA = "S" Then
            tab_llave!TAB_TIPREG = 1000
         Else
            tab_llave!TAB_TIPREG = 1001
         End If
      Else
         tab_llave.Edit
      End If
      tab_llave!tab_NOMLARGO = PUB_NUM_OPER
      tab_llave!tab_nomcorto = Format(LK_FECHA_DIA, "dd/mm/yyyy")
      tab_llave!TAB_contable2 = WS_SALDO
      tab_llave.Update
   End If
   
VAMOS:
wran1 = "D5"
wran2 = "D" & f1
xl.Visible = True
wranF = "D" & f1 + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
wran1 = "E5"
wran2 = "E" & f1
wranF = "E" & f1 + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
wranF = "D" & 5 & ":D" & f1
xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 3
wranF = "E" & 5 & ":E" & f1
xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 3
wranF = "F" & 5 & ":F" & f1
xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 3

FrmImp2.ProgBar.Value = todos_cli
Screen.MousePointer = 0
DoEvents
FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
DoEvents
xcuenta = 1
xl.Cells(2, 2) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
xl.Cells(2, 5) = "'" & Format(wsFECHA1, "dd mmm yyyy")
xl.DisplayAlerts = False
xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
xl.Worksheets("Hoja1").Protect PUB_CLAVE

DoEvents
FrmImp2.lblproceso.Visible = False
FrmImp2.ProgBar.Visible = False
xl.APPLICATION.Visible = True
Set xl = Nothing
Screen.MousePointer = 0
FrmImp2.Pantalla.Enabled = True
FrmImp2.Pantalla.Caption = "Por &Pantalla"
FrmImp2.lblproceso.Visible = False

Exit Sub


WEXCEL:
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FrmImp2.lblproceso.Caption = "Abriendo , Archivo Caja_Fecha.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\OFFICE\CAJA_FECHA.xls", 0, True, 4, WPAS, WPAS
Return

Exit Sub

'LLENA_VALOR:
'For I = QJ To WDIF
'  If I >= 3 Then
'    xl.Cells(F1 + 7, I) = Format(LOC_VALOR, "0.000")
'  End If
'Next I
'Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
Exit Sub

End Sub


Public Sub REG_COMPRA1()
Dim dni_cliente As String
'On Error GoTo FINTODO
Dim TD_F As String * 1
Dim TD_B As String * 1
Dim TD_N As String * 1
Dim TD_D As String * 1
Dim w_exo As Currency
Dim ws_codcia As String * 2
Dim FILAS As Integer
Dim nn, m_ind As Integer
Dim AWQ_NETO_ACT_FIJO As Currency
Dim AWQ_CTA_ACT_FIJO As String
Dim AWQ_BRUTO_ACT_FIJO As Currency
Dim WCONTROL As Integer
FILAS = 5
Dim WQ
Dim wfecha
Dim ws_clave
Dim LETRAS(24) As String * 1
Dim wRuta As String
Dim wmonto As Currency
Dim wcodclie As Currency
Dim valor_venta As Currency
Dim WQ_TIPMOV As Integer
Dim descto As Currency
Dim valor_igv As Currency
Dim valor_precio As Currency
Dim S_VALOR_VENTA As Currency
Dim s_descto As Currency
Dim s_valor_igv As Currency
Dim S_VALOR_PRECIO As Currency
'AGREGADO PARA RESUMEN MIC
Dim MicS_VALOR_VENTA1 As Currency
Dim MicS_VALOR_VENTA  As Currency
Dim Mics_descto  As Currency
Dim Mics_valor_igv  As Currency
Dim MicS_VALOR_PRECIO   As Currency
'===============================
Dim t_valor_venta As Currency
Dim t_valor_venta1 As Currency
Dim t_descto As Currency
Dim t_valor_igv As Currency
Dim t_valor_precio As Currency

Dim wnumfac As Currency
Dim WFLAG As String * 1
Dim wsFECHA1, wsFECHA2
Dim xcuenta As Integer
Dim wq_fecha, wq_codclie, wq_codven, wq_docu, wq_nombre, wq_bruto, wq_gastos, wq_desto, wq_flete, wq_fbg, wq_serie
Dim wq_tot_descto, WQ_IMPTO, WQ_TOTAL, wq_estado, wq_condi
Dim wq_ruc
Dim wflag_numfac
Dim WQ_BRUTO_D, WQ_GASTOS_D, WQ_FLETE_D, WQ_IMPTO_D, WQ_TOTAL_D As Currency
Dim wserie As String * 3
Dim AWQ_BRUTO As Currency
Dim AWQ_DESCTOS As Currency
Dim AWQ_GASTOS As Currency
Dim AWQ_FLETES As Currency
Dim AWQ_IMPTO As Currency
Dim AWQ_NETO As Currency
Dim AWQ_NETO_CRED  As Currency
Dim AWQ_NETO_CONT   As Currency
Dim AWQ_COSTO_VENTA As Currency
Dim AWQ_Descuento As Currency
Dim AWQ_Descuento_D As Currency
Dim WS_SIGNO As Integer
Dim ws_tc As Currency
Dim wsTexto As String

Dim var_tipo As String * 1
Dim S_VALOR_VENTA1 As Currency
Dim cont_bruto1 As Currency
Dim cont_bruto As Currency
Dim cont_igv As Currency
Dim cont_total As Currency
Dim cred_bruto1 As Currency
Dim cred_bruto As Currency
Dim cred_igv As Currency
Dim cred_total As Currency


dni_cliente = ""
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) <> cop_llave!cop_fecha_proceso Then cheasiento.Value = 0
If CDate(wsFECHA2) <> cop_llave!cop_fecha_proceso2 Then cheasiento.Value = 0
GoSub WEXCEL
If LK_EMP = "PIU" Then
  max.Text = 90000
End If
pub_cadena = ""
xcuenta = 0
TD_F = ""
TD_B = ""
TD_N = ""
TD_D = ""
wsTexto = "TIPO: "
For fila = 0 To 3
 txttp.ListIndex = fila
 If fila = 0 And txttp.Selected(fila) Then
   TD_F = "F"
   wsTexto = wsTexto + "- FACT."
 End If
 If fila = 1 And txttp.Selected(fila) Then
   TD_B = "B"
   wsTexto = wsTexto + "- BOLE."
 End If
 If fila = 2 And txttp.Selected(fila) Then
   TD_N = "N"
   wsTexto = wsTexto + "- NCRE."
 End If
 If fila = 3 And txttp.Selected(fila) Then
   TD_D = "D"
   wsTexto = wsTexto + "- NDEB."
 End If
Next fila

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = ""
WCONTROL = 0
f1 = 5  'Fila Inicial
t_valor_venta = 0
t_valor_venta1 = 0
t_descto = 0
t_valor_igv = 0
t_valor_precio = 0

'NCREDITO: ' empieza

'WCONTROL = WCONTROL + 1
wmonto = 0
wcodclie = 0
valor_venta = 0
descto = 0
valor_igv = 0
valor_precio = 0
wnumfac = 0
WFLAG = ""
xcuenta = 0
wflag_numfac = ""
wserie = ""
AWQ_BRUTO = 0
AWQ_DESCTOS = 0
AWQ_GASTOS = 0
AWQ_FLETES = 0
AWQ_IMPTO = 0
AWQ_NETO = 0
AWQ_NETO_CRED = 0
AWQ_NETO_CONT = 0
AWQ_COSTO_VENTA = 0
AWQ_NETO_ACT_FIJO = 0
AWQ_BRUTO_ACT_FIJO = 0



f1 = 4
FILAS = 0
FILAS = FILAS + 1
f1 = f1 + 1

NCREDITO:
cont_bruto1 = 0
cont_bruto = 0
cont_igv = 0
cont_total = 0
cred_bruto1 = 0
cred_bruto = 0
cred_igv = 0
cred_total = 0

S_VALOR_VENTA1 = 0
S_VALOR_VENTA = 0
s_descto = 0
s_valor_igv = 0
S_VALOR_PRECIO = 0

WCONTROL = WCONTROL + 1

If WCONTROL = 1 Then
  If TD_F = "F" Or TD_B = "B" Then
  Else
    GoTo OTRO_DOCUMENTO
  End If
  pub_cadena = "SELECT * FROM  FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND FAR_TIPMOV = 10 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = ? OR FAR_FBG = ? )  ORDER BY  FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC,FAR_NUMSEC" 'AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "'
  WS_SIGNO = 1
ElseIf WCONTROL = 2 Then
  If TD_N = "N" Then
  Else
    GoTo OTRO_DOCUMENTO
  End If
  pub_cadena = "SELECT *  FROM FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND FAR_TIPMOV = 97 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?  AND FAR_CP = 'C'  ORDER BY FAR_TIPMOV, FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC " ' FAR_TIPMOV, FAR_FBG DESC , FAR_FECHA_COMPRA,  FAR_NUMSER,FAR_NUMFAC"'AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "'
  WS_SIGNO = -1
ElseIf WCONTROL = 3 Then
  If TD_D = "D" Then
  Else
    GoTo OTRO_DOCUMENTO
  End If
  pub_cadena = "SELECT *  FROM FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? OR FAR_CODCIA= ? OR FAR_CODCIA= ? ) AND FAR_TIPMOV = 98 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?   AND FAR_CP = 'C'  ORDER BY FAR_TIPMOV, FAR_FBG DESC , FAR_FECHA_COMPRA,  FAR_NUMSER,FAR_NUMFAC" 'AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "'
  WS_SIGNO = 1
End If

If Trim(txtSerie.Text) <> "" Then
If WCONTROL = 1 Then
  pub_cadena = "SELECT * FROM  FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND FAR_TIPMOV = 20 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = ? OR FAR_FBG = ? )AND FAR_NUMSER = ?  ORDER BY FAR_TIPMOV, FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC" 'AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "'
  WS_SIGNO = 1
End If
End If

If Trim(fbtxt.Text) <> "" And Trim(txtSerie.Text) <> "" Then
If WCONTROL = 1 Then
  pub_cadena = "SELECT * FROM  FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND FAR_TIPMOV = 20 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = ? OR FAR_FBG = ? )AND FAR_NUMSER = ?  ORDER BY FAR_TIPMOV, FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC" 'AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "'
  WS_SIGNO = 1
End If
End If

If Right(max.Text, 1) = "D" Or Right(max.Text, 1) = "S" Then
If WCONTROL = 1 Then
  pub_cadena = "SELECT * FROM  FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND FAR_TIPMOV = 20 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = ? OR FAR_FBG = ? )AND FAR_MONEDA = ?   ORDER BY FAR_TIPMOV, FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC" 'AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "'
  WS_SIGNO = 1
End If
End If




Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = ""
PS_REP02(1) = ""
PS_REP02(2) = ""
PS_REP02(3) = ""
PS_REP02(4) = ""
PS_REP02(5) = wsFECHA1
PS_REP02(6) = wsFECHA2
If WCONTROL = 1 Then
 PS_REP02(7) = TD_F
 PS_REP02(8) = TD_B
End If


Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

PS_REP02(0) = ""
PS_REP02(1) = ""
PS_REP02(2) = ""
PS_REP02(3) = ""
PS_REP02(4) = ""
If checia.Visible And checia.Value = 1 Then
   If Trim(par_llave!par_art_cias) <> "" Then
     nn = 1
     For m_ind = 1 To 15
         ws_codcia = Mid(par_llave!par_art_cias, nn, 2)
         If Trim(ws_codcia) = "" Then Exit For
         PS_REP02(m_ind - 1) = ws_codcia
         nn = nn + 2
     Next m_ind
   End If
Else
PS_REP02(0) = LK_CODCIA
End If


' el PS_REP1(0) ESTA MAS ABAJO
PS_REP02(5) = wsFECHA1
PS_REP02(6) = wsFECHA2

If WCONTROL = 1 Then
PS_REP02(7) = TD_F
PS_REP02(8) = TD_B
''If fbtxt.Text = "F" Or fbtxt.Text = "B" Then
''   PS_REP02(7) = fbtxt.Text
''   PS_REP02(8) = fbtxt.Text
''End If
End If

If Trim(txtSerie.Text) <> "" Then
   PS_REP02(9) = txtSerie.Text
End If
If Right(max.Text, 1) = "S" Or Right(max.Text, 1) = "D" Then
   PS_REP02(9) = Right(max.Text, 1)
End If


DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep02.Requery
If llave_rep02.EOF Then
  GoTo OTRO_DOCUMENTO
End If
FrmImp2.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.max = llave_rep02.RowCount

WFLAG = ""
SQ_OPER = 1
pu_cp = "C"
pu_codcia = llave_rep02!FAR_CODCIA
wnumfac = llave_rep02!far_numfac
wfecha = llave_rep02!far_numser 'llave_rep02!far_fecha
wserie = llave_rep02!far_numser
wq_fecha = llave_rep02!FAR_fecha_compra
WQ_TIPMOV = llave_rep02!FAR_TIPMOV
wq_fbg = Trim(llave_rep02!far_fbg)

xcuenta = 0
WFLAG = "A"
wflag_numfac = "A"
AWQ_DESCTOS = 0
AWQ_GASTOS = 0
AWQ_FLETES = 0
w_exo = 0
If llave_rep02.EOF Then GoTo CANCELA
Do Until llave_rep02.EOF

   If Trim(txtSerie.Text) <> "" Then
      If Trim(txtSerie.Text) <> Trim(llave_rep02!far_numser) Then GoTo SALTARIN
   End If
   
   
'  If llave_rep02!FAR_numfac = 22891 And Val(llave_rep02!FAR_numser) = 4 Then Stop
  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  If Trim(wfecha) <> Trim(llave_rep02!far_numser) Then '
     If wflag_numfac = "A" Then
       GoSub IMPRI_FAC
     End If
     wflag_numfac = ""
     f1 = f1 + 1
     GoSub TOTAL_DIA
     '´t_valor_venta = t_valor_venta + S_VALOR_VENTA
     't_descto = t_descto + s_descto
     't_valor_igv = t_valor_igv + s_valor_igv
     't_valor_precio = t_valor_precio + S_VALOR_PRECIO
     wnumfac = llave_rep02!far_numfac
     wfecha = llave_rep02!far_numser
     wq_fbg = Trim(llave_rep02!far_fbg)
     WQ_TIPMOV = llave_rep02!FAR_TIPMOV
     wq_serie = llave_rep02!far_numser
     wserie = llave_rep02!far_numser
     S_VALOR_VENTA = 0
     s_descto = 0
     s_valor_igv = 0
     S_VALOR_PRECIO = 0
     w_exo = 0
  End If
  
  
  If wnumfac = val(llave_rep02!far_numfac) And val(wserie) = val(llave_rep02!far_numser) And wq_fbg = llave_rep02!far_fbg And WQ_TIPMOV = llave_rep02!FAR_TIPMOV Then
  Else
     GoSub IMPRI_FAC
     wflag_numfac = ""
  End If
  
    
  wnumfac = llave_rep02!far_numfac
  wq_fecha = Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy")
  wq_codclie = llave_rep02!far_codclie
  wq_codven = llave_rep02!FAR_CODVEN
  wq_fbg = Trim(llave_rep02!far_fbg)
  WQ_TIPMOV = llave_rep02!FAR_TIPMOV
  wq_serie = "'" & llave_rep02!far_numser
  wserie = llave_rep02!far_numser
  wq_docu = "'" & llave_rep02!far_numfac
  wq_nombre = ""
  ws_tc = 1
  If llave_rep02!FAR_MONEDA = "D" Then
    ws_tc = JALAR(llave_rep02!FAR_fecha_compra)
    If ws_tc <= 0 Then
        MsgBox "Falta Ingresar el Tipo de Cambio del día : " & Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy"), 48, Pub_Titulo
        GoTo CANCELA
    End If
    If llave_rep02!far_estado <> "E" Then
       WQ_BRUTO_D = Format((val(llave_rep02!FAR_BRUTO) - val(llave_rep02!FAR_TOT_DESCTO)) * WS_SIGNO, "0.000")
       WQ_GASTOS_D = val(llave_rep02!FAR_GASTOS) * WS_SIGNO
       WQ_FLETE_D = val(Nulo_Valor0(llave_rep02!FAR_TOT_FLETE)) * WS_SIGNO
       WQ_IMPTO_D = Format(val(llave_rep02!far_IMPTO) * WS_SIGNO, "0.000")
       WQ_TOTAL_D = val(WQ_BRUTO_D) + val(WQ_IMPTO_D)
       AWQ_Descuento_D = val(Nulo_Valor0(llave_rep02!FAR_TOT_DESCTO)) * WS_SIGNO
    End If
  End If
  wq_bruto = Format((val(llave_rep02!FAR_BRUTO) - val(llave_rep02!FAR_TOT_DESCTO)) * WS_SIGNO, "0.000") ' BLOQUEADO POR DIF DSTO MIC
  'wq_bruto = Format((Val(llave_rep02!FAR_BRUTO)) * WS_SIGNO, "0.000") 'AGRGADO REEMPLAZANDO LA LINEA DE ARRIBA MIC
  AWQ_Descuento = val(Nulo_Valor0(llave_rep02!FAR_TOT_DESCTO)) * WS_SIGNO
  wq_gastos = val(llave_rep02!FAR_GASTOS) * WS_SIGNO * ws_tc
  wq_flete = val(Nulo_Valor0(llave_rep02!FAR_TOT_FLETE)) * WS_SIGNO
  WQ_IMPTO = Format(val(llave_rep02!far_IMPTO) * WS_SIGNO, "0.000")
  WQ_TOTAL = val(wq_bruto) + val(WQ_IMPTO)  ' (Val(llave_rep02!far_bruto) + Val(llave_rep02!far_impto) - Val(llave_rep02!FAR_TOT_DESCTO) + Val(llave_rep02!FAR_GASTOS)) * WS_SIGNO * WS_TC
  If llave_rep02!FAR_EX_IGV = "A" Then
     w_exo = w_exo + val(llave_rep02!FAR_SUBTOTAL)
  End If

  If llave_rep02!far_estado = "E" Then
     wq_bruto = 0
     wq_gastos = 0
     wq_flete = 0
     WQ_IMPTO = 0
     WQ_TOTAL = 0
     w_exo = 0
     WQ_BRUTO_D = 0
     WQ_GASTOS_D = 0
     WQ_FLETE_D = 0
     WQ_IMPTO_D = 0
     WQ_TOTAL_D = 0
 End If
  
  
  
  If wq_bruto = 0 Then WQ_TOTAL = 0
  wq_estado = llave_rep02!far_estado
  If wq_estado <> "E" Then
    If Left(UCase(llave_rep02!far_subtra), 1) <> "A" Then
     AWQ_COSTO_VENTA = AWQ_COSTO_VENTA + ((llave_rep02!FAR_COSPRO * llave_rep02!far_cantidad)) * WS_SIGNO
    End If
  End If
  If llave_rep02!far_signo_car <> 0 And llave_rep02!FAR_DIAS <> 0 Then
     var_tipo = "1"
  Else
     var_tipo = "0"
  End If
  wflag_numfac = "A"
 WFLAG = "A"
pu_codcia = llave_rep02!FAR_CODCIA
wfecha = llave_rep02!far_numser
SALTARIN:
 llave_rep02.MoveNext
Loop
If wflag_numfac = "A" Then
    GoSub IMPRI_FAC
    wflag_numfac = ""
End If
If WFLAG = "A" Then
    f1 = f1 + 1
    FILAS = FILAS + 1
    GoSub TOTAL_DIA
End If
'  If cheasiento.Value = 1 Then
'    FrmImp2.lblProceso.Caption = "Procesando Pase de Contabilidad . . . "
'    DoEvents
'    GoSub PASE_CONTAB
'  End If
  xcuenta = c1 + 1
  'wranF = "A6:" & "K6"
  'xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3
  If WCONTROL = 1 Then
   If cheasiento.Value = 1 Then
'    FrmImp2.lblProceso.Caption = "Procesando Pase de Contabilidad . . . "
'    DoEvents
'    GoSub PASE_CONTAB
   End If
  End If
OTRO_DOCUMENTO:
If WCONTROL >= 3 Or Trim(fbtxt.Text) <> "" Or Trim(txtSerie.Text) <> "" Or Right(max.Text, 1) = "S" Or Right(max.Text, 1) = "D" Then
Else
  GoTo NCREDITO
End If

MOSTRAR:
   If cheasiento.Value = 1 Then
    FrmImp2.lblproceso.Caption = "Procesando Pase de Contabilidad . . . "
    DoEvents
    GoSub PASE_CONTAB
   End If


  f1 = f1 + 2
 ' xl.Cells(F1, 1) = "Total General = "
 ' xl.Worksheets(1).Rows(F1).RowHeight = 20
 ' xl.Cells(F1, 8) = t_valor_venta
 ' xl.Cells(F1, 9) = t_descto
 ' xl.Cells(F1, 11) = t_valor_igv
 ' xl.Cells(F1, 12) = t_valor_precio
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  'MOSTRANDO RESUMEN MIC
  xl.Cells(f1, 1) = "T O T A L    G E N E R A L     = "
  xl.Cells(f1, 8) = MicS_VALOR_VENTA1
  xl.Cells(f1, 10) = MicS_VALOR_VENTA
  xl.Cells(f1, 9) = Mics_descto
  xl.Cells(f1, 11) = Mics_valor_igv
  xl.Cells(f1, 12) = MicS_VALOR_PRECIO
  '============================
  If checia.Visible And checia.Value = 1 Then
    xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  Else
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  End If
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'" & wsTexto & " DEL " & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub

IMPRI_FAC:
     FILAS = FILAS + 1
     f1 = f1 + 1
     pu_codclie = wq_codclie
     LEER_CLI_LLAVE
     dni_cliente = " "
     wq_ruc = " "
     If Not cli_llave.EOF Then
        wq_codclie = cli_llave!cli_codclie 'xl.Cells(F1, 2)
        wq_nombre = Trim(cli_llave!CLI_NOMBRE)
        wq_ruc = Trim(cli_llave!cli_ruc_esposo)
        dni_cliente = Trim(cli_llave!cli_RUC_ESPOSA)
        If dni_cliente = "" Then dni_cliente = " "
        If wq_ruc = "" Then wq_ruc = " "
     End If
     If wq_fbg = "F" Then
        wq_condi = "01"
     ElseIf wq_fbg = "B" Then
        wq_condi = "03"
     ElseIf wq_fbg = "N" Then
        wq_condi = "07"
     ElseIf wq_fbg = "D" Then
        wq_condi = "08"
     End If
     xl.Cells(f1, 1) = "'" & Format(wq_fecha, "dd/mm/yy")
     
     xl.Cells(f1, 6) = wq_nombre '2
     If wq_fbg = "B" Then
      xl.Cells(f1, 5) = dni_cliente '3
     Else
      xl.Cells(f1, 5) = wq_ruc '3
     End If
     xl.Cells(f1, 2) = "'" & wq_condi   '4
     ' xl.Cells(F1, 5) = wq_fbg '5
     xl.Cells(f1, 3) = wq_serie '6
     xl.Cells(f1, 4) = wq_docu '7
     
     If wq_estado = "E" Then
         xl.Cells(f1, 6) = wq_nombre
         xl.Cells(f1, 9) = "[ A  N U L A D O ] "
     Else
         xl.Cells(f1, 6) = wq_nombre
     End If
     
     If wq_estado <> "E" Then
      If True Then 'If Left(cli_llave!CLI_CUENTA_CONTAB, 2) <> "12" Then
       xl.Cells(f1, 8) = Format(val(wq_bruto), "0.000")
       S_VALOR_VENTA1 = S_VALOR_VENTA1 + (wq_bruto - w_exo)
       If var_tipo = "0" Then
         cont_bruto1 = cont_bruto1 + (wq_bruto - w_exo)
       Else
         cred_bruto1 = cred_bruto1 + (wq_bruto - w_exo)
       End If
       t_valor_venta1 = t_valor_venta1 + (wq_bruto - w_exo)
       If val(ws_tc) <> 1 Then
         xl.Cells(f1, 9) = val(ws_tc)
         wq_bruto = Format(val(wq_bruto) * ws_tc, "0.000")
         WQ_IMPTO = Format(val(wq_bruto) * (LK_IGV / 100), "0.000")
         WQ_TOTAL = val(wq_bruto) + val(WQ_IMPTO)
        End If
       xl.Cells(f1, 10) = Format(wq_bruto, "0.000")
       xl.Cells(f1, 11) = val(WQ_IMPTO)
       xl.Cells(f1, 12) = val(WQ_TOTAL)
      Else
       xl.Cells(f1, 8) = Format(val(wq_bruto) - w_exo, "0.000")
       If val(ws_tc) <> 1 Then
         xl.Cells(f1, 9) = val(ws_tc)
       End If
       xl.Cells(f1, 11) = val(WQ_IMPTO)
       xl.Cells(f1, 12) = val(WQ_TOTAL)
       xl.Cells(f1, 15) = Trim(cli_llave!CLI_CUENTA_CONTAB)
       AWQ_IMPTO = AWQ_IMPTO + val(WQ_IMPTO)
       AWQ_NETO = AWQ_NETO + val(WQ_TOTAL)
       AWQ_BRUTO = AWQ_BRUTO + wq_bruto
       AWQ_DESCTOS = wq_desto
       AWQ_GASTOS = wq_gastos
       AWQ_FLETES = wq_flete
       If wq_condi = "CRED" Then
         xl.Cells(f1, 15) = -1
         AWQ_NETO_CRED = AWQ_NETO_CRED + val(WQ_TOTAL)
       Else
         xl.Cells(f1, 15) = 0
         AWQ_NETO_CONT = AWQ_NETO_CONT + val(WQ_TOTAL)
       End If
     End If
     End If
     If ws_tc > 1 And Right(max.Text, 1) = "D" Then
       xl.Cells(f1, 16) = Format(val(WQ_BRUTO_D) - w_exo, "0.000")
       xl.Cells(f1, 17) = val(WQ_IMPTO_D)
       xl.Cells(f1, 18) = val(WQ_TOTAL_D)
       xl.Cells(f1, 19) = " " & ws_tc
    End If
     
     
     S_VALOR_VENTA = S_VALOR_VENTA + (wq_bruto - w_exo)
     s_descto = s_descto + val(w_exo)
     s_valor_igv = s_valor_igv + val(WQ_IMPTO)
     S_VALOR_PRECIO = S_VALOR_PRECIO + val(WQ_TOTAL)
    
     t_valor_venta = t_valor_venta + (wq_bruto - w_exo)
     t_descto = t_descto + val(w_exo)
     t_valor_igv = t_valor_igv + val(WQ_IMPTO)
     t_valor_precio = t_valor_precio + val(WQ_TOTAL)
     
     If var_tipo = "0" Then
      cont_bruto = cont_bruto + (wq_bruto - w_exo)
      cont_igv = cont_igv + val(WQ_IMPTO)
      cont_total = cont_total + val(WQ_TOTAL)
     Else
      cred_bruto = cred_bruto + (wq_bruto - w_exo)
      cred_igv = cred_igv + val(WQ_IMPTO)
      cred_total = cred_total + val(WQ_TOTAL)
     End If
     
     
     
     If FILAS >= val(max.Text) Then
        f1 = f1 + 1
        xl.Cells(f1, 1) = "VAN ... "
        xl.Worksheets(1).Rows(f1).RowHeight = 20
        xl.Cells(f1, 8) = t_valor_venta1
        xl.Cells(f1, 10) = t_valor_venta
        Debug.Print t_valor_venta
        xl.Cells(f1, 9) = t_descto
        xl.Cells(f1, 11) = t_valor_igv
        xl.Cells(f1, 12) = t_valor_precio
        'F1 = F1 + 1
        wranF = "A" & f1
        xl.APPLICATION.Range(wranF).Select
        On Error Resume Next
        xl.APPLICATION.ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
        FILAS = 1
        f1 = f1 + 1
        xl.Cells(f1, 1) = "VIENEN ... "
        xl.Worksheets(1).Rows(f1).RowHeight = 20
        xl.Cells(f1, 8) = t_valor_venta1
        xl.Cells(f1, 10) = t_valor_venta
        xl.Cells(f1, 9) = t_descto
        xl.Cells(f1, 11) = t_valor_igv
        xl.Cells(f1, 12) = t_valor_precio
    End If

Return

CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ""
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\REGVENTA.xls", 0, True, 4
Return

TOTAL_DIA:
  f1 = f1 + 1
  xl.Cells(f1, 1) = "Total Credito = "
  'xl.Worksheets(1).Rows(F1).RowHeight = 20
  xl.Cells(f1, 8) = cred_bruto1
  xl.Cells(f1, 10) = cred_bruto
  xl.Cells(f1, 9) = 0
  xl.Cells(f1, 11) = cred_igv
  xl.Cells(f1, 12) = cred_total
  f1 = f1 + 1
  xl.Cells(f1, 1) = "Total Contado = "
  'xl.Worksheets(1).Rows(F1).RowHeight = 20
  xl.Cells(f1, 8) = cont_bruto1
  xl.Cells(f1, 10) = cont_bruto
  xl.Cells(f1, 9) = 0
  xl.Cells(f1, 11) = cont_igv
  xl.Cells(f1, 12) = cont_total
  f1 = f1 + 1
  xl.Cells(f1, 1) = "Total   = "
  xl.Worksheets(1).Rows(f1).RowHeight = 20
  xl.Cells(f1, 2) = ""
  xl.Cells(f1, 3) = ""
  xl.Cells(f1, 7) = ""
  xl.Cells(f1, 8) = S_VALOR_VENTA1
  xl.Cells(f1, 10) = S_VALOR_VENTA
  xl.Cells(f1, 9) = s_descto
  xl.Cells(f1, 11) = s_valor_igv
  xl.Cells(f1, 12) = S_VALOR_PRECIO
  
  'AGREGADO PARA RESUMEN TOTAL MIC
  MicS_VALOR_VENTA1 = MicS_VALOR_VENTA1 + S_VALOR_VENTA1
  MicS_VALOR_VENTA = MicS_VALOR_VENTA + S_VALOR_VENTA
  Mics_descto = Mics_descto + s_descto
  Mics_valor_igv = Mics_valor_igv + s_valor_igv
  MicS_VALOR_PRECIO = MicS_VALOR_PRECIO + S_VALOR_PRECIO
  '=====================================
  
  cont_bruto = 0
  cont_igv = 0
  cont_total = 0
  cred_bruto = 0
  cred_igv = 0
  cred_total = 0
Return

PASE_CONTAB:
Dim wcta As String
Dim wcta_clientes As Currency
Dim PS_CONTAB1 As rdoQuery
Dim contab_llave As rdoResultset
Dim ws_nro_voucher As Integer
Dim ws_nro_sec As Integer
Dim ws_glosa As String
Dim wsq_fecha As String
Dim wdh As String * 1
Dim wscodcia As String * 2
Dim wsq_fecha2
wscodcia = LK_CODCIA
ws_glosa = "Registro de Venta"
If LK_EMP_PTO = "A" Then
 wscodcia = "00"
 ws_glosa = "Registro de Venta - " & Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
End If
wsq_fecha = Format(cop_llave!cop_fecha_proceso, "yyyy/mm/dd")
wsq_fecha2 = Format(cop_llave!cop_fecha_proceso2, "yyyy/mm/dd")
pub_cadena = "DELETE COMOV  WHERE  COV_FLAG_AUTOMATICA = '3' AND COV_CODUSU = '" & LK_CODCIA & "' AND (COV_FECHA_VOUCHER >=  ' " & wsq_fecha & "' AND COV_FECHA_VOUCHER <=  ' " & wsq_fecha2 & "')"
CN.Execute pub_cadena, rdExecDirect

FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.max = 10
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
PSCOV_VOUCHER(0) = wscodcia
PSCOV_VOUCHER(1) = cop_llave!cop_fecha_proceso
PSCOV_VOUCHER(2) = cop_llave!cop_fecha_proceso2
cov_voucher.Requery
If cov_voucher.EOF Then
 ws_nro_voucher = 0
Else
 cov_voucher.MoveLast
 ws_nro_voucher = cov_voucher!COV_NRO_VOUCHER
End If
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
wran1 = "A" & 6 & ":Q" & f1
xl.APPLICATION.Worksheets("Hoja1").Range(wran1).Sort Key1:=xl.APPLICATION.Worksheets("Hoja1").Range("N6")
fila = 6
' xl.Application.Visible = True
wcta = Trim(xl.Cells(fila, 14))
wcta_clientes = 0
wdh = "D"
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
For fila = 6 To f1
  If val(xl.Cells(fila, 15)) = 0 Then GoTo OTRITO
  'If Trim(xl.Cells(fila, 5)) = "N" Then
  '   GoTo OTRITO
  'End If
  If val(xl.Cells(fila, 14)) = 0 Then GoTo OTRITO
   If wcta <> Trim(xl.Cells(fila, 14)) Then
     GoSub Graba
     wcta_clientes = 0
     wcta = Trim(xl.Cells(fila, 14))
     wcta_clientes = wcta_clientes + val(xl.Cells(fila, 11))
  Else
     wcta_clientes = wcta_clientes + val(xl.Cells(fila, 11))
  End If
OTRITO:
Next fila
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
GoSub Graba
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1

If AWQ_NETO_ACT_FIJO <> 0 Then
 SQ_OPER = 1
 PUB_SECUENCIA = 2
 PUB_CODTRA = 2401
 PUB_CODCIA = wscodcia
 LEER_CNT_LLAVE
 If cnt_llave.EOF Then
  MsgBox "Error de Dato de Transaccion , Consulte a su Proveedor.", 48, Pub_Titulo
  End
 End If
 If Trim(cnt_llave!CNT_CTA1) <> "" Then
  wcta = cnt_llave!CNT_CTA1  'Ventas a costo
  wdh = cnt_llave!CNT_DH1
  wcta_clientes = AWQ_BRUTO_ACT_FIJO
  GoSub Graba
 End If
  wcta = AWQ_CTA_ACT_FIJO
  wdh = "D"
  wcta_clientes = AWQ_NETO_ACT_FIJO
  GoSub Graba
End If

SQ_OPER = 1
PUB_SECUENCIA = 24
PUB_CODTRA = 2401
PUB_CODCIA = wscodcia
LEER_CNT_LLAVE
If cnt_llave.EOF Then
 MsgBox "Error de Dato de Transaccion , Consulte a su Proveedor.", 48, Pub_Titulo
 'End
  GoTo CANCELA
End If
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1

If Trim(cnt_llave!CNT_CTA1) <> "" Then
 wcta = cnt_llave!CNT_CTA1 'Ventas Brutas
 wdh = cnt_llave!CNT_DH1
 wcta_clientes = AWQ_BRUTO
 GoSub Graba
End If
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
If Trim(cnt_llave!CNT_CTA2) <> "" Then
 wcta = cnt_llave!CNT_CTA2 'Impuesto
 wdh = cnt_llave!CNT_DH2
 wcta_clientes = AWQ_IMPTO
 GoSub Graba
End If

If AWQ_NETO_CONT <> 0 Then
 SQ_OPER = 1
 PUB_SECUENCIA = 24
 PUB_CODTRA = 2401
 PUB_CODCIA = wscodcia
 LEER_CNT_LLAVE
 If cnt_llave.EOF Then
  MsgBox "Error de Dato de Transaccion , Consulte a su Proveedor.", 48, Pub_Titulo
  End
 End If
 If Trim(cnt_llave!CNT_CTA1) <> "" Then
  wcta = cnt_llave!CNT_CTA1  'Ventas a costo
  wdh = "D" ' ES HABER PERO AL DEBE
  wcta_clientes = AWQ_NETO_CONT
  GoSub Graba
 End If
End If
SQ_OPER = 1
PUB_SECUENCIA = 24
PUB_CODTRA = 2401
PUB_CODCIA = wscodcia
LEER_CNT_LLAVE
If cnt_llave.EOF Then
 MsgBox "Error de Dato de Transaccion , Consulte a su Proveedor.", 48, Pub_Titulo
 End
End If
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
ws_nro_voucher = ws_nro_voucher + 1
If Trim(cnt_llave!CNT_CTA3) <> "" Then
 wcta = cnt_llave!CNT_CTA3  'Ventas a costo
 wdh = cnt_llave!CNT_DH3
 wcta_clientes = AWQ_COSTO_VENTA
 GoSub Graba
End If
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
If Trim(cnt_llave!CNT_CTA4) <> "" Then
 wcta = cnt_llave!CNT_CTA4  'Ventas a costo
 wdh = cnt_llave!CNT_DH4
 wcta_clientes = AWQ_COSTO_VENTA
 GoSub Graba
End If

FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1

Return

Graba:
     If wcta_clientes = 0 Then Return
     ws_nro_sec = ws_nro_sec + 1
     cov_voucher.AddNew
     cov_voucher!COV_CODCIA = wscodcia
     cov_voucher!COV_FECHA_VOUCHER = cop_llave!cop_fecha_proceso2
     cov_voucher!COV_NRO_MOV = ws_nro_sec
     cov_voucher!COV_NRO_VOUCHER = ws_nro_voucher
     cov_voucher!COV_NUMTAB = 0
     cov_voucher!COV_CODCTA = wcta
     cov_voucher!COV_DH = wdh
     cov_voucher!COV_IMPORTE = wcta_clientes
     cov_voucher!COV_ESTADO = " "
     cov_voucher!COV_CODUSU = LK_CODCIA
     cov_voucher!cov_flag_automatica = "3"
     cov_voucher!COV_glosa = ws_glosa
     cov_voucher!COV_FECHA_doc = LK_FECHA_DIA
     cov_voucher.Update
Return


LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"

Return

FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
 GoTo CANCELA
 Resume Next
End Sub

Public Sub REG_VENTA()
Dim dni_cliente As String
'On Error GoTo FINTODO
Dim TD_F As String * 1
Dim TD_B As String * 1
Dim TD_N As String * 1
Dim TD_D As String * 1
Dim w_exo As Currency
Dim ws_codcia As String * 2
Dim FILAS As Integer
Dim nn, m_ind As Integer
Dim AWQ_NETO_ACT_FIJO As Currency
Dim AWQ_CTA_ACT_FIJO As String
Dim AWQ_BRUTO_ACT_FIJO As Currency
Dim WCONTROL As Integer
FILAS = 5
Dim WQ
Dim wfecha
Dim ws_clave
Dim LETRAS(24) As String * 1
Dim wRuta As String
Dim wmonto As Currency
Dim wcodclie As Currency
Dim valor_venta As Currency
Dim WQ_TIPMOV As Integer
Dim descto As Currency
Dim valor_igv As Currency
Dim valor_precio As Currency
Dim S_VALOR_VENTA As Currency
Dim s_descto As Currency
Dim s_valor_igv As Currency
Dim S_VALOR_PRECIO As Currency
'AGREGADO PARA RESUMEN MIC
Dim MicS_VALOR_VENTA1 As Currency
Dim MicS_VALOR_VENTA  As Currency
Dim Mics_descto  As Currency
Dim Mics_valor_igv  As Currency
Dim MicS_VALOR_PRECIO   As Currency
'===============================
Dim t_valor_venta As Currency
Dim t_valor_venta1 As Currency
Dim t_descto As Currency
Dim t_valor_igv As Currency
Dim t_valor_precio As Currency

Dim wnumfac As Currency
Dim WFLAG As String * 1
Dim wsFECHA1, wsFECHA2
Dim xcuenta As Integer
Dim wq_fecha, wq_codclie, wq_codven, wq_docu, wq_nombre, wq_bruto, wq_gastos, wq_desto, wq_flete, wq_fbg, wq_serie
Dim wq_tot_descto, WQ_IMPTO, WQ_TOTAL, wq_estado, wq_condi
Dim wq_ruc
Dim wflag_numfac
Dim WQ_BRUTO_D, WQ_GASTOS_D, WQ_FLETE_D, WQ_IMPTO_D, WQ_TOTAL_D As Currency
Dim wserie As String * 3
Dim AWQ_BRUTO As Currency
Dim AWQ_DESCTOS As Currency
Dim AWQ_GASTOS As Currency
Dim AWQ_FLETES As Currency
Dim AWQ_IMPTO As Currency
Dim AWQ_NETO As Currency
Dim AWQ_NETO_CRED  As Currency
Dim AWQ_NETO_CONT   As Currency
Dim AWQ_COSTO_VENTA As Currency
Dim AWQ_Descuento As Currency
Dim AWQ_Descuento_D As Currency
Dim WS_SIGNO As Integer
Dim ws_tc As Currency
Dim wsTexto As String

Dim var_tipo As String * 1
Dim S_VALOR_VENTA1 As Currency
Dim cont_bruto1 As Currency
Dim cont_bruto As Currency
Dim cont_igv As Currency
Dim cont_total As Currency
Dim cred_bruto1 As Currency
Dim cred_bruto As Currency
Dim cred_igv As Currency
Dim cred_total As Currency


dni_cliente = ""
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) <> cop_llave!cop_fecha_proceso Then cheasiento.Value = 0
If CDate(wsFECHA2) <> cop_llave!cop_fecha_proceso2 Then cheasiento.Value = 0
GoSub WEXCEL
If LK_EMP = "PIU" Then
  max.Text = 90000
End If
pub_cadena = ""
xcuenta = 0
TD_F = ""
TD_B = ""
TD_N = ""
TD_D = ""
wsTexto = "TIPO: "
For fila = 0 To 3
 txttp.ListIndex = fila
 If fila = 0 And txttp.Selected(fila) Then
   TD_F = "F"
   wsTexto = wsTexto + "- FACT."
 End If
 If fila = 1 And txttp.Selected(fila) Then
   TD_B = "B"
   wsTexto = wsTexto + "- BOLE."
 End If
 If fila = 2 And txttp.Selected(fila) Then
   TD_N = "N"
   wsTexto = wsTexto + "- NCRE."
 End If
 If fila = 3 And txttp.Selected(fila) Then
   TD_D = "D"
   wsTexto = wsTexto + "- NDEB."
 End If
Next fila

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = ""
WCONTROL = 0
f1 = 5  'Fila Inicial
t_valor_venta = 0
t_valor_venta1 = 0
t_descto = 0
t_valor_igv = 0
t_valor_precio = 0

'NCREDITO: ' empieza

'WCONTROL = WCONTROL + 1
wmonto = 0
wcodclie = 0
valor_venta = 0
descto = 0
valor_igv = 0
valor_precio = 0
wnumfac = 0
WFLAG = ""
xcuenta = 0
wflag_numfac = ""
wserie = ""
AWQ_BRUTO = 0
AWQ_DESCTOS = 0
AWQ_GASTOS = 0
AWQ_FLETES = 0
AWQ_IMPTO = 0
AWQ_NETO = 0
AWQ_NETO_CRED = 0
AWQ_NETO_CONT = 0
AWQ_COSTO_VENTA = 0
AWQ_NETO_ACT_FIJO = 0
AWQ_BRUTO_ACT_FIJO = 0



f1 = 4
FILAS = 0
FILAS = FILAS + 1
f1 = f1 + 1

NCREDITO:
cont_bruto1 = 0
cont_bruto = 0
cont_igv = 0
cont_total = 0
cred_bruto1 = 0
cred_bruto = 0
cred_igv = 0
cred_total = 0

S_VALOR_VENTA1 = 0
S_VALOR_VENTA = 0
s_descto = 0
s_valor_igv = 0
S_VALOR_PRECIO = 0

WCONTROL = WCONTROL + 1

If WCONTROL = 1 Then
  If TD_F = "F" Or TD_B = "B" Then
  Else
    GoTo OTRO_DOCUMENTO
  End If
  pub_cadena = "SELECT * FROM  FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND FAR_TIPMOV = 10 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = ? OR FAR_FBG = ? )  ORDER BY  FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC,FAR_NUMSEC" 'AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "'
  WS_SIGNO = 1
ElseIf WCONTROL = 2 Then
  If TD_N = "N" Then
  Else
    GoTo OTRO_DOCUMENTO
  End If
  pub_cadena = "SELECT *  FROM FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND FAR_TIPMOV = 97 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?  AND FAR_CP = 'C'  ORDER BY FAR_TIPMOV, FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC " ' FAR_TIPMOV, FAR_FBG DESC , FAR_FECHA_COMPRA,  FAR_NUMSER,FAR_NUMFAC"'AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "'
  WS_SIGNO = -1
ElseIf WCONTROL = 3 Then
  If TD_D = "D" Then
  Else
    GoTo OTRO_DOCUMENTO
  End If
  pub_cadena = "SELECT *  FROM FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? OR FAR_CODCIA= ? OR FAR_CODCIA= ? ) AND FAR_TIPMOV = 98 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?   AND FAR_CP = 'C'  ORDER BY FAR_TIPMOV, FAR_FBG DESC , FAR_FECHA_COMPRA,  FAR_NUMSER,FAR_NUMFAC" 'AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "'
  WS_SIGNO = 1
End If

If Trim(txtSerie.Text) <> "" Then
If WCONTROL = 1 Then
  pub_cadena = "SELECT * FROM  FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND FAR_TIPMOV = 10 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = ? OR FAR_FBG = ? )AND FAR_NUMSER = ?  ORDER BY FAR_TIPMOV, FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC" 'AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "'
  WS_SIGNO = 1
End If
End If

If Trim(fbtxt.Text) <> "" And Trim(txtSerie.Text) <> "" Then
If WCONTROL = 1 Then
  pub_cadena = "SELECT * FROM  FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND FAR_TIPMOV = 10 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = ? OR FAR_FBG = ? )AND FAR_NUMSER = ?  ORDER BY FAR_TIPMOV, FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC" 'AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "'
  WS_SIGNO = 1
End If
End If

If Right(max.Text, 1) = "D" Or Right(max.Text, 1) = "S" Then
If WCONTROL = 1 Then
  pub_cadena = "SELECT * FROM  FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND FAR_TIPMOV = 10 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = ? OR FAR_FBG = ? )AND FAR_MONEDA = ?   ORDER BY FAR_TIPMOV, FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC" 'AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "'
  WS_SIGNO = 1
End If
End If




Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = ""
PS_REP02(1) = ""
PS_REP02(2) = ""
PS_REP02(3) = ""
PS_REP02(4) = ""
PS_REP02(5) = wsFECHA1
PS_REP02(6) = wsFECHA2
If WCONTROL = 1 Then
 PS_REP02(7) = TD_F
 PS_REP02(8) = TD_B
End If


Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

PS_REP02(0) = ""
PS_REP02(1) = ""
PS_REP02(2) = ""
PS_REP02(3) = ""
PS_REP02(4) = ""
If checia.Visible And checia.Value = 1 Then
   If Trim(par_llave!par_art_cias) <> "" Then
     nn = 1
     For m_ind = 1 To 15
         ws_codcia = Mid(par_llave!par_art_cias, nn, 2)
         If Trim(ws_codcia) = "" Then Exit For
         PS_REP02(m_ind - 1) = ws_codcia
         nn = nn + 2
     Next m_ind
   End If
Else
PS_REP02(0) = LK_CODCIA
End If


' el PS_REP1(0) ESTA MAS ABAJO
PS_REP02(5) = wsFECHA1
PS_REP02(6) = wsFECHA2

If WCONTROL = 1 Then
PS_REP02(7) = TD_F
PS_REP02(8) = TD_B
''If fbtxt.Text = "F" Or fbtxt.Text = "B" Then
''   PS_REP02(7) = fbtxt.Text
''   PS_REP02(8) = fbtxt.Text
''End If
End If

If Trim(txtSerie.Text) <> "" Then
   PS_REP02(9) = txtSerie.Text
End If
If Right(max.Text, 1) = "S" Or Right(max.Text, 1) = "D" Then
   PS_REP02(9) = Right(max.Text, 1)
End If


DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep02.Requery
If llave_rep02.EOF Then
  GoTo OTRO_DOCUMENTO
End If
FrmImp2.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.max = llave_rep02.RowCount

WFLAG = ""
SQ_OPER = 1
pu_cp = "C"
pu_codcia = llave_rep02!FAR_CODCIA
wnumfac = llave_rep02!far_numfac
wfecha = llave_rep02!far_numser 'llave_rep02!far_fecha
wserie = llave_rep02!far_numser
wq_fecha = llave_rep02!FAR_fecha_compra
WQ_TIPMOV = llave_rep02!FAR_TIPMOV
wq_fbg = Trim(llave_rep02!far_fbg)

xcuenta = 0
WFLAG = "A"
wflag_numfac = "A"
AWQ_DESCTOS = 0
AWQ_GASTOS = 0
AWQ_FLETES = 0
w_exo = 0
If llave_rep02.EOF Then GoTo CANCELA
Do Until llave_rep02.EOF

   If Trim(txtSerie.Text) <> "" Then
      If Trim(txtSerie.Text) <> Trim(llave_rep02!far_numser) Then GoTo SALTARIN
   End If
   
   
'  If llave_rep02!FAR_numfac = 22891 And Val(llave_rep02!FAR_numser) = 4 Then Stop
  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  If Trim(wfecha) <> Trim(llave_rep02!far_numser) Then '
     If wflag_numfac = "A" Then
       GoSub IMPRI_FAC
     End If
     wflag_numfac = ""
     f1 = f1 + 1
     GoSub TOTAL_DIA
     '´t_valor_venta = t_valor_venta + S_VALOR_VENTA
     't_descto = t_descto + s_descto
     't_valor_igv = t_valor_igv + s_valor_igv
     't_valor_precio = t_valor_precio + S_VALOR_PRECIO
     wnumfac = llave_rep02!far_numfac
     wfecha = llave_rep02!far_numser
     wq_fbg = Trim(llave_rep02!far_fbg)
     WQ_TIPMOV = llave_rep02!FAR_TIPMOV
     wq_serie = llave_rep02!far_numser
     wserie = llave_rep02!far_numser
     S_VALOR_VENTA = 0
     s_descto = 0
     s_valor_igv = 0
     S_VALOR_PRECIO = 0
     w_exo = 0
  End If
  
  
  If wnumfac = val(llave_rep02!far_numfac) And val(wserie) = val(llave_rep02!far_numser) And wq_fbg = llave_rep02!far_fbg And WQ_TIPMOV = llave_rep02!FAR_TIPMOV Then
  Else
     GoSub IMPRI_FAC
     wflag_numfac = ""
  End If
  
    
  wnumfac = llave_rep02!far_numfac
  wq_fecha = Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy")
  wq_codclie = llave_rep02!far_codclie
  wq_codven = llave_rep02!FAR_CODVEN
  wq_fbg = Trim(llave_rep02!far_fbg)
  WQ_TIPMOV = llave_rep02!FAR_TIPMOV
  wq_serie = "'" & llave_rep02!far_numser
  wserie = llave_rep02!far_numser
  wq_docu = "'" & llave_rep02!far_numfac
  wq_nombre = ""
  ws_tc = 1
  If llave_rep02!FAR_MONEDA = "D" Then
    ws_tc = JALAR(llave_rep02!FAR_fecha_compra)
    If ws_tc <= 0 Then
        MsgBox "Falta Ingresar el Tipo de Cambio del día : " & Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy"), 48, Pub_Titulo
        GoTo CANCELA
    End If
    If llave_rep02!far_estado <> "E" Then
       WQ_BRUTO_D = Format((val(llave_rep02!FAR_BRUTO) - val(llave_rep02!FAR_TOT_DESCTO)) * WS_SIGNO, "0.000")
       WQ_GASTOS_D = val(llave_rep02!FAR_GASTOS) * WS_SIGNO
       WQ_FLETE_D = val(Nulo_Valor0(llave_rep02!FAR_TOT_FLETE)) * WS_SIGNO
       WQ_IMPTO_D = Format(val(llave_rep02!far_IMPTO) * WS_SIGNO, "0.000")
       WQ_TOTAL_D = val(WQ_BRUTO_D) + val(WQ_IMPTO_D)
       AWQ_Descuento_D = val(Nulo_Valor0(llave_rep02!FAR_TOT_DESCTO)) * WS_SIGNO
    End If
  End If
  wq_bruto = Format((val(llave_rep02!FAR_BRUTO) - val(llave_rep02!FAR_TOT_DESCTO) - val(llave_rep02!FAR_GASTOS)) * WS_SIGNO, "0.000") ' BLOQUEADO POR DIF DSTO MIC
  'wq_bruto = Format((Val(llave_rep02!FAR_BRUTO)) * WS_SIGNO, "0.000") 'AGRGADO REEMPLAZANDO LA LINEA DE ARRIBA MIC
  AWQ_Descuento = val(Nulo_Valor0(llave_rep02!FAR_TOT_DESCTO)) * WS_SIGNO
  wq_gastos = val(llave_rep02!FAR_GASTOS) * WS_SIGNO * ws_tc
  wq_flete = val(Nulo_Valor0(llave_rep02!FAR_TOT_FLETE)) * WS_SIGNO
  WQ_IMPTO = Format(val(llave_rep02!far_IMPTO) * WS_SIGNO, "0.000")
  WQ_TOTAL = val(wq_bruto) + val(WQ_IMPTO)  ' (Val(llave_rep02!far_bruto) + Val(llave_rep02!far_impto) - Val(llave_rep02!FAR_TOT_DESCTO) + Val(llave_rep02!FAR_GASTOS)) * WS_SIGNO * WS_TC
  If llave_rep02!FAR_EX_IGV = "A" Then
     w_exo = w_exo + val(llave_rep02!FAR_SUBTOTAL)
  End If

  If llave_rep02!far_estado = "E" Then
     wq_bruto = 0
     wq_gastos = 0
     wq_flete = 0
     WQ_IMPTO = 0
     WQ_TOTAL = 0
     w_exo = 0
     WQ_BRUTO_D = 0
     WQ_GASTOS_D = 0
     WQ_FLETE_D = 0
     WQ_IMPTO_D = 0
     WQ_TOTAL_D = 0
 End If
  
  
  
  If wq_bruto = 0 Then WQ_TOTAL = 0
  wq_estado = llave_rep02!far_estado
  If wq_estado <> "E" Then
    If Left(UCase(llave_rep02!far_subtra), 1) <> "A" Then
     AWQ_COSTO_VENTA = AWQ_COSTO_VENTA + ((llave_rep02!FAR_COSPRO * llave_rep02!far_cantidad)) * WS_SIGNO
    End If
  End If
  If llave_rep02!far_signo_car <> 0 And llave_rep02!FAR_DIAS <> 0 Then
     var_tipo = "1"
  Else
     var_tipo = "0"
  End If
  wflag_numfac = "A"
 WFLAG = "A"
pu_codcia = llave_rep02!FAR_CODCIA
wfecha = llave_rep02!far_numser
SALTARIN:
 llave_rep02.MoveNext
Loop
If wflag_numfac = "A" Then
    GoSub IMPRI_FAC
    wflag_numfac = ""
End If
If WFLAG = "A" Then
    f1 = f1 + 1
    FILAS = FILAS + 1
    GoSub TOTAL_DIA
End If
'  If cheasiento.Value = 1 Then
'    FrmImp2.lblProceso.Caption = "Procesando Pase de Contabilidad . . . "
'    DoEvents
'    GoSub PASE_CONTAB
'  End If
  xcuenta = c1 + 1
  'wranF = "A6:" & "K6"
  'xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3
  If WCONTROL = 1 Then
   If cheasiento.Value = 1 Then
'    FrmImp2.lblProceso.Caption = "Procesando Pase de Contabilidad . . . "
'    DoEvents
'    GoSub PASE_CONTAB
   End If
  End If
OTRO_DOCUMENTO:
If WCONTROL >= 3 Or Trim(fbtxt.Text) <> "" Or Trim(txtSerie.Text) <> "" Or Right(max.Text, 1) = "S" Or Right(max.Text, 1) = "D" Then
Else
  GoTo NCREDITO
End If

MOSTRAR:
   If cheasiento.Value = 1 Then
    FrmImp2.lblproceso.Caption = "Procesando Pase de Contabilidad . . . "
    DoEvents
    GoSub PASE_CONTAB
   End If


  f1 = f1 + 2
 ' xl.Cells(F1, 1) = "Total General = "
 ' xl.Worksheets(1).Rows(F1).RowHeight = 20
 ' xl.Cells(F1, 8) = t_valor_venta
 ' xl.Cells(F1, 9) = t_descto
 ' xl.Cells(F1, 11) = t_valor_igv
 ' xl.Cells(F1, 12) = t_valor_precio
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  'MOSTRANDO RESUMEN MIC
  xl.Cells(f1, 1) = "T O T A L    G E N E R A L     = "
  xl.Cells(f1, 8) = MicS_VALOR_VENTA1
  xl.Cells(f1, 10) = MicS_VALOR_VENTA
  xl.Cells(f1, 9) = Mics_descto
  xl.Cells(f1, 11) = Mics_valor_igv
  xl.Cells(f1, 12) = MicS_VALOR_PRECIO
  '============================
  If checia.Visible And checia.Value = 1 Then
    xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  Else
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  End If
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'" & wsTexto & " DEL " & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub

IMPRI_FAC:
     FILAS = FILAS + 1
     f1 = f1 + 1
     pu_codclie = wq_codclie
     LEER_CLI_LLAVE
     dni_cliente = " "
     wq_ruc = " "
     If Not cli_llave.EOF Then
        wq_codclie = cli_llave!cli_codclie 'xl.Cells(F1, 2)
        wq_nombre = Trim(cli_llave!CLI_NOMBRE)
        wq_ruc = Trim(cli_llave!cli_ruc_esposo)
        dni_cliente = Trim(cli_llave!cli_RUC_ESPOSA)
        If dni_cliente = "" Then dni_cliente = " "
        If wq_ruc = "" Then wq_ruc = " "
     End If
     If wq_fbg = "F" Then
        wq_condi = "01"
     ElseIf wq_fbg = "B" Then
        wq_condi = "03"
     ElseIf wq_fbg = "N" Then
        wq_condi = "07"
     ElseIf wq_fbg = "D" Then
        wq_condi = "08"
     End If
     xl.Cells(f1, 1) = "'" & Format(wq_fecha, "dd/mm/yy")
     
     xl.Cells(f1, 6) = wq_nombre '2
     If wq_fbg = "B" Then
      xl.Cells(f1, 5) = dni_cliente '3
     Else
      xl.Cells(f1, 5) = wq_ruc '3
     End If
     xl.Cells(f1, 2) = "'" & wq_condi   '4
     ' xl.Cells(F1, 5) = wq_fbg '5
     xl.Cells(f1, 3) = wq_serie '6
     xl.Cells(f1, 4) = wq_docu '7
     
     If wq_estado = "E" Then
         xl.Cells(f1, 5) = " "       '3  agregado GTS
         xl.Cells(f1, 6) = "[ A  N U L A D O ] "
         xl.Cells(f1, 9) = " "
     Else
         xl.Cells(f1, 6) = wq_nombre
     End If
     
     If wq_estado <> "E" Then
      If True Then 'If Left(cli_llave!CLI_CUENTA_CONTAB, 2) <> "12" Then
       xl.Cells(f1, 8) = Format(val(wq_bruto), "0.000")
       S_VALOR_VENTA1 = S_VALOR_VENTA1 + (wq_bruto - w_exo)
       If var_tipo = "0" Then
         cont_bruto1 = cont_bruto1 + (wq_bruto - w_exo)
       Else
         cred_bruto1 = cred_bruto1 + (wq_bruto - w_exo)
       End If
       t_valor_venta1 = t_valor_venta1 + (wq_bruto - w_exo)
       If val(ws_tc) <> 1 Then
         xl.Cells(f1, 9) = val(ws_tc)
         wq_bruto = Format(val(wq_bruto) * ws_tc, "0.000")
         WQ_IMPTO = Format(val(wq_bruto) * (LK_IGV / 100), "0.000")
         WQ_TOTAL = val(wq_bruto) + val(WQ_IMPTO)
        End If
       xl.Cells(f1, 10) = Format(wq_bruto, "0.000")
       xl.Cells(f1, 11) = val(WQ_IMPTO)
       xl.Cells(f1, 12) = val(WQ_TOTAL)
      Else
       xl.Cells(f1, 8) = Format(val(wq_bruto) - w_exo, "0.000")
       If val(ws_tc) <> 1 Then
         xl.Cells(f1, 9) = val(ws_tc)
       End If
       xl.Cells(f1, 11) = val(WQ_IMPTO)
       xl.Cells(f1, 12) = val(WQ_TOTAL)
       xl.Cells(f1, 15) = Trim(cli_llave!CLI_CUENTA_CONTAB)
       AWQ_IMPTO = AWQ_IMPTO + val(WQ_IMPTO)
       AWQ_NETO = AWQ_NETO + val(WQ_TOTAL)
       AWQ_BRUTO = AWQ_BRUTO + wq_bruto
       AWQ_DESCTOS = wq_desto
       AWQ_GASTOS = wq_gastos
       AWQ_FLETES = wq_flete
       If wq_condi = "CRED" Then
         xl.Cells(f1, 15) = -1
         AWQ_NETO_CRED = AWQ_NETO_CRED + val(WQ_TOTAL)
       Else
         xl.Cells(f1, 15) = 0
         AWQ_NETO_CONT = AWQ_NETO_CONT + val(WQ_TOTAL)
       End If
     End If
     End If
     If ws_tc > 1 And Right(max.Text, 1) = "D" Then
       xl.Cells(f1, 16) = Format(val(WQ_BRUTO_D) - w_exo, "0.000")
       xl.Cells(f1, 17) = val(WQ_IMPTO_D)
       xl.Cells(f1, 18) = val(WQ_TOTAL_D)
       xl.Cells(f1, 19) = " " & ws_tc
    End If
     
     
     S_VALOR_VENTA = S_VALOR_VENTA + (wq_bruto - w_exo)
     s_descto = s_descto + val(w_exo)
     s_valor_igv = s_valor_igv + val(WQ_IMPTO)
     S_VALOR_PRECIO = S_VALOR_PRECIO + val(WQ_TOTAL)
    
     t_valor_venta = t_valor_venta + (wq_bruto - w_exo)
     t_descto = t_descto + val(w_exo)
     t_valor_igv = t_valor_igv + val(WQ_IMPTO)
     t_valor_precio = t_valor_precio + val(WQ_TOTAL)
     
     If var_tipo = "0" Then
      cont_bruto = cont_bruto + (wq_bruto - w_exo)
      cont_igv = cont_igv + val(WQ_IMPTO)
      cont_total = cont_total + val(WQ_TOTAL)
     Else
      cred_bruto = cred_bruto + (wq_bruto - w_exo)
      cred_igv = cred_igv + val(WQ_IMPTO)
      cred_total = cred_total + val(WQ_TOTAL)
     End If
     
     
     
     If FILAS >= val(max.Text) Then
        f1 = f1 + 1
        xl.Cells(f1, 1) = "VAN ... "
        xl.Worksheets(1).Rows(f1).RowHeight = 20
        'xl.Cells(f1, 8) = t_valor_venta1     ' quitado GTS para no sumar columna H
        xl.Cells(f1, 10) = t_valor_venta
        Debug.Print t_valor_venta
        xl.Cells(f1, 9) = t_descto
        xl.Cells(f1, 11) = t_valor_igv
        xl.Cells(f1, 12) = t_valor_precio
        'F1 = F1 + 1
        wranF = "A" & f1
        xl.APPLICATION.Range(wranF).Select
        On Error Resume Next
        xl.APPLICATION.ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
        FILAS = 1
        f1 = f1 + 1
        xl.Cells(f1, 1) = "VIENEN ... "
        xl.Worksheets(1).Rows(f1).RowHeight = 20
        'xl.Cells(f1, 8) = t_valor_venta1        ' quitado GTS para no sumar columna H
        xl.Cells(f1, 10) = t_valor_venta
        xl.Cells(f1, 9) = t_descto
        xl.Cells(f1, 11) = t_valor_igv
        xl.Cells(f1, 12) = t_valor_precio
    End If

Return

CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ""
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\REGVENTA.xls", 0, True, 4
Return

TOTAL_DIA:
  'f1 = f1 + 1
  'xl.Cells(f1, 1) = "Total Credito = "
  'xl.Worksheets(1).Rows(F1).RowHeight = 20
  'xl.Cells(f1, 8) = cred_bruto1
  'xl.Cells(f1, 10) = cred_bruto
  'xl.Cells(f1, 9) = 0
  'xl.Cells(f1, 11) = cred_igv
  'xl.Cells(f1, 12) = cred_total
  'f1 = f1 + 1
  'xl.Cells(f1, 1) = "Total Contado = "
  'xl.Worksheets(1).Rows(F1).RowHeight = 20
  'xl.Cells(f1, 8) = cont_bruto1
  'xl.Cells(f1, 10) = cont_bruto
  'xl.Cells(f1, 9) = 0
  'xl.Cells(f1, 11) = cont_igv
  'xl.Cells(f1, 12) = cont_total
  'f1 = f1 + 1
  'xl.Cells(f1, 1) = "Total   = "
  'xl.Worksheets(1).Rows(f1).RowHeight = 20
  'xl.Cells(f1, 2) = ""
  'xl.Cells(f1, 3) = ""
  'xl.Cells(f1, 7) = ""
  'xl.Cells(f1, 8) = S_VALOR_VENTA1
  'xl.Cells(f1, 10) = S_VALOR_VENTA
  'xl.Cells(f1, 9) = s_descto
  'xl.Cells(f1, 11) = s_valor_igv
  'xl.Cells(f1, 12) = S_VALOR_PRECIO
  
  'AGREGADO PARA RESUMEN TOTAL MIC
  'MicS_VALOR_VENTA1 = MicS_VALOR_VENTA1 + S_VALOR_VENTA1
  MicS_VALOR_VENTA = MicS_VALOR_VENTA + S_VALOR_VENTA
  Mics_descto = Mics_descto + s_descto
  Mics_valor_igv = Mics_valor_igv + s_valor_igv
  MicS_VALOR_PRECIO = MicS_VALOR_PRECIO + S_VALOR_PRECIO
  '=====================================
  
  cont_bruto = 0
  cont_igv = 0
  cont_total = 0
  cred_bruto = 0
  cred_igv = 0
  cred_total = 0
Return

PASE_CONTAB:
Dim wcta As String
Dim wcta_clientes As Currency
Dim PS_CONTAB1 As rdoQuery
Dim contab_llave As rdoResultset
Dim ws_nro_voucher As Integer
Dim ws_nro_sec As Integer
Dim ws_glosa As String
Dim wsq_fecha As String
Dim wdh As String * 1
Dim wscodcia As String * 2
Dim wsq_fecha2
wscodcia = LK_CODCIA
ws_glosa = "Registro de Venta"
If LK_EMP_PTO = "A" Then
 wscodcia = "00"
 ws_glosa = "Registro de Venta - " & Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
End If
wsq_fecha = Format(cop_llave!cop_fecha_proceso, "yyyy/mm/dd")
wsq_fecha2 = Format(cop_llave!cop_fecha_proceso2, "yyyy/mm/dd")
pub_cadena = "DELETE COMOV  WHERE  COV_FLAG_AUTOMATICA = '3' AND COV_CODUSU = '" & LK_CODCIA & "' AND (COV_FECHA_VOUCHER >=  ' " & wsq_fecha & "' AND COV_FECHA_VOUCHER <=  ' " & wsq_fecha2 & "')"
CN.Execute pub_cadena, rdExecDirect

FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.max = 10
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
PSCOV_VOUCHER(0) = wscodcia
PSCOV_VOUCHER(1) = cop_llave!cop_fecha_proceso
PSCOV_VOUCHER(2) = cop_llave!cop_fecha_proceso2
cov_voucher.Requery
If cov_voucher.EOF Then
 ws_nro_voucher = 0
Else
 cov_voucher.MoveLast
 ws_nro_voucher = cov_voucher!COV_NRO_VOUCHER
End If
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
wran1 = "A" & 6 & ":Q" & f1
xl.APPLICATION.Worksheets("Hoja1").Range(wran1).Sort Key1:=xl.APPLICATION.Worksheets("Hoja1").Range("N6")
fila = 6
' xl.Application.Visible = True
wcta = Trim(xl.Cells(fila, 14))
wcta_clientes = 0
wdh = "D"
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
For fila = 6 To f1
  If val(xl.Cells(fila, 15)) = 0 Then GoTo OTRITO
  'If Trim(xl.Cells(fila, 5)) = "N" Then
  '   GoTo OTRITO
  'End If
  If val(xl.Cells(fila, 14)) = 0 Then GoTo OTRITO
   If wcta <> Trim(xl.Cells(fila, 14)) Then
     GoSub Graba
     wcta_clientes = 0
     wcta = Trim(xl.Cells(fila, 14))
     wcta_clientes = wcta_clientes + val(xl.Cells(fila, 11))
  Else
     wcta_clientes = wcta_clientes + val(xl.Cells(fila, 11))
  End If
OTRITO:
Next fila
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
GoSub Graba
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1

If AWQ_NETO_ACT_FIJO <> 0 Then
 SQ_OPER = 1
 PUB_SECUENCIA = 2
 PUB_CODTRA = 2401
 PUB_CODCIA = wscodcia
 LEER_CNT_LLAVE
 If cnt_llave.EOF Then
  MsgBox "Error de Dato de Transaccion , Consulte a su Proveedor.", 48, Pub_Titulo
  End
 End If
 If Trim(cnt_llave!CNT_CTA1) <> "" Then
  wcta = cnt_llave!CNT_CTA1  'Ventas a costo
  wdh = cnt_llave!CNT_DH1
  wcta_clientes = AWQ_BRUTO_ACT_FIJO
  GoSub Graba
 End If
  wcta = AWQ_CTA_ACT_FIJO
  wdh = "D"
  wcta_clientes = AWQ_NETO_ACT_FIJO
  GoSub Graba
End If

SQ_OPER = 1
PUB_SECUENCIA = 24
PUB_CODTRA = 2401
PUB_CODCIA = wscodcia
LEER_CNT_LLAVE
If cnt_llave.EOF Then
 MsgBox "Error de Dato de Transaccion , Consulte a su Proveedor.", 48, Pub_Titulo
 'End
  GoTo CANCELA
End If
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1

If Trim(cnt_llave!CNT_CTA1) <> "" Then
 wcta = cnt_llave!CNT_CTA1 'Ventas Brutas
 wdh = cnt_llave!CNT_DH1
 wcta_clientes = AWQ_BRUTO
 GoSub Graba
End If
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
If Trim(cnt_llave!CNT_CTA2) <> "" Then
 wcta = cnt_llave!CNT_CTA2 'Impuesto
 wdh = cnt_llave!CNT_DH2
 wcta_clientes = AWQ_IMPTO
 GoSub Graba
End If

If AWQ_NETO_CONT <> 0 Then
 SQ_OPER = 1
 PUB_SECUENCIA = 24
 PUB_CODTRA = 2401
 PUB_CODCIA = wscodcia
 LEER_CNT_LLAVE
 If cnt_llave.EOF Then
  MsgBox "Error de Dato de Transaccion , Consulte a su Proveedor.", 48, Pub_Titulo
  End
 End If
 If Trim(cnt_llave!CNT_CTA1) <> "" Then
  wcta = cnt_llave!CNT_CTA1  'Ventas a costo
  wdh = "D" ' ES HABER PERO AL DEBE
  wcta_clientes = AWQ_NETO_CONT
  GoSub Graba
 End If
End If
SQ_OPER = 1
PUB_SECUENCIA = 24
PUB_CODTRA = 2401
PUB_CODCIA = wscodcia
LEER_CNT_LLAVE
If cnt_llave.EOF Then
 MsgBox "Error de Dato de Transaccion , Consulte a su Proveedor.", 48, Pub_Titulo
 End
End If
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
ws_nro_voucher = ws_nro_voucher + 1
If Trim(cnt_llave!CNT_CTA3) <> "" Then
 wcta = cnt_llave!CNT_CTA3  'Ventas a costo
 wdh = cnt_llave!CNT_DH3
 wcta_clientes = AWQ_COSTO_VENTA
 GoSub Graba
End If
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
If Trim(cnt_llave!CNT_CTA4) <> "" Then
 wcta = cnt_llave!CNT_CTA4  'Ventas a costo
 wdh = cnt_llave!CNT_DH4
 wcta_clientes = AWQ_COSTO_VENTA
 GoSub Graba
End If

FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1

Return

Graba:
     If wcta_clientes = 0 Then Return
     ws_nro_sec = ws_nro_sec + 1
     cov_voucher.AddNew
     cov_voucher!COV_CODCIA = wscodcia
     cov_voucher!COV_FECHA_VOUCHER = cop_llave!cop_fecha_proceso2
     cov_voucher!COV_NRO_MOV = ws_nro_sec
     cov_voucher!COV_NRO_VOUCHER = ws_nro_voucher
     cov_voucher!COV_NUMTAB = 0
     cov_voucher!COV_CODCTA = wcta
     cov_voucher!COV_DH = wdh
     cov_voucher!COV_IMPORTE = wcta_clientes
     cov_voucher!COV_ESTADO = " "
     cov_voucher!COV_CODUSU = LK_CODCIA
     cov_voucher!cov_flag_automatica = "3"
     cov_voucher!COV_glosa = ws_glosa
     cov_voucher!COV_FECHA_doc = LK_FECHA_DIA
     cov_voucher.Update
Return


LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"

Return

FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
 GoTo CANCELA
 Resume Next
End Sub


Public Sub REG_COMPRA()
'On Error GoTo FINTODO
Dim AWQ_NETO_ACT_FIJO As Currency
Dim AWQ_CTA_ACT_FIJO As String
Dim AWQ_BRUTO_ACT_FIJO As Currency
Dim WCONTROL As Integer
Dim WQ
Dim wfecha
Dim ws_clave
Dim LETRAS(24) As String * 1
Dim wRuta As String
Dim wmonto As Currency
Dim wcodclie As Currency
Dim valor_venta As Currency
Dim descto As Currency
Dim valor_igv As Currency
Dim valor_precio As Currency
Dim S_VALOR_VENTA As Currency
Dim s_descto As Currency
Dim s_valor_igv As Currency
Dim S_VALOR_PRECIO As Currency

Dim t_valor_venta As Currency
Dim t_descto As Currency
Dim t_valor_igv As Currency
Dim t_valor_precio As Currency

Dim wnumfac As Currency
Dim WFLAG As String * 1
Dim wsFECHA1, wsFECHA2
Dim xcuenta As Integer
Dim wq_fecha, wq_codclie, wq_codven, wq_docu, wq_nombre, wq_bruto, wq_gastos, wq_desto, wq_flete, wq_fbg, wq_serie
Dim wq_tot_descto, WQ_IMPTO, WQ_TOTAL, wq_estado, wq_condi
Dim wq_ruc
Dim wflag_numfac
Dim wserie As String * 3
Dim AWQ_BRUTO As Currency
Dim AWQ_DESCTOS As Currency
Dim AWQ_GASTOS As Currency
Dim AWQ_FLETES As Currency
Dim AWQ_IMPTO As Currency
Dim AWQ_NETO As Currency
Dim AWQ_NETO_CRED  As Currency
Dim AWQ_NETO_CONT   As Currency
Dim AWQ_COSTO_VENTA As Currency
Dim WS_SIGNO As Integer
Dim wcodcta, wcodcta2, wtipo_mov
Dim wwwserie As String * 3
Dim www_fbg
Dim www_tipmov
Dim AWQ_TIPO_CAMBIO As Currency
Dim FLAG_TC As Integer
' VAR CONTABLES
Dim var_secuencia As Integer

' Agregado
Dim wq_fecha_compra, wq_guia

var_secuencia = 0
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) <> cop_llave!cop_fecha_proceso Then cheasiento.Value = 0
If CDate(wsFECHA2) <> cop_llave!cop_fecha_proceso2 Then cheasiento.Value = 0
GoSub WEXCEL
pub_cadena = ""
xcuenta = 0

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
WCONTROL = 0
f1 = 5  'Fila Inicial
t_valor_venta = 0
t_descto = 0
t_valor_igv = 0
t_valor_precio = 0

'NCREDITO: ' empieza

'WCONTROL = WCONTROL + 1
wmonto = 0
wcodclie = 0
valor_venta = 0
descto = 0
valor_igv = 0
valor_precio = 0
wnumfac = 0
WFLAG = ""
xcuenta = 0
wflag_numfac = ""
wserie = ""
wwwserie = ""
AWQ_BRUTO = 0
AWQ_DESCTOS = 0
AWQ_GASTOS = 0
AWQ_FLETES = 0
AWQ_IMPTO = 0
AWQ_NETO = 0
AWQ_NETO_CRED = 0
AWQ_NETO_CONT = 0
AWQ_COSTO_VENTA = 0
AWQ_NETO_ACT_FIJO = 0
AWQ_BRUTO_ACT_FIJO = 0

NCREDITO:
S_VALOR_VENTA = 0
s_descto = 0
s_valor_igv = 0
S_VALOR_PRECIO = 0

WCONTROL = WCONTROL + 1
If WCONTROL = 1 Then ' Compras(20) y Provisiones (99)
  If val(txt_cli.Text) <> 0 Then
   'pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCLIE = " & Trim(txt_cli.Text) & " AND FAR_CODCIA = ? AND (FAR_TIPMOV = 20 OR FAR_TIPMOV = 99)  AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND FAR_CP = 'P' AND FAR_ESTADO <> 'E' ORDER BY FAR_FBG,FAR_FECHA_COMPRA, FAR_NUMOPER"
   pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCLIE = " & Trim(txt_cli.Text) & " AND FAR_CODCIA = ? AND (FAR_TIPMOV = 20 OR FAR_TIPMOV = 99)  AND FAR_FECHA_PRO >= ? AND FAR_FECHA_PRO <= ? AND FAR_CP = 'P'  ORDER BY FAR_FBG,FAR_NUMFAC,FAR_FECHA_PRO, FAR_NUMOPER"
  Else
   pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND (FAR_TIPMOV = 20 OR FAR_TIPMOV = 99)  AND FAR_FECHA_PRO >= ? AND FAR_FECHA_PRO <= ? AND FAR_CP = 'P'  ORDER BY FAR_FBG,FAR_NUMFAC,FAR_FECHA_PRO, FAR_NUMOPER"
  End If
  WS_SIGNO = 1
ElseIf WCONTROL = 2 Then ' Notas de Crédito
  If val(txt_cli.Text) <> 0 Then
    'pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCLIE = " & Trim(txt_cli.Text) & " AND FAR_CODCIA = ? AND FAR_TIPMOV = 97 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND FAR_FBG = 'C' AND FAR_ESTADO <> 'E' AND FAR_CP = 'P'  ORDER BY FAR_FECHA_COMPRA, FAR_NUMOPER"
    pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCLIE = " & Trim(txt_cli.Text) & " AND FAR_CODCIA = ? AND FAR_TIPMOV = 97 AND FAR_FECHA_PRO >= ? AND FAR_FECHA_PRO <= ? AND FAR_FBG = 'C'  AND FAR_CP = 'P'  ORDER BY FAR_NUMFAC,FAR_FECHA_PRO, FAR_NUMOPER"
  Else
    pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = 97 AND FAR_FECHA_PRO >= ? AND FAR_FECHA_PRO <= ? AND FAR_FBG = 'C'  AND FAR_CP = 'P'  ORDER BY FAR_NUMFAC,FAR_FECHA_PRO, FAR_NUMOPER"
  End If
  WS_SIGNO = -1
ElseIf WCONTROL = 3 Then
  'pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = 22298 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND FAR_FBG = 'A' and FAR_ESTADO <> 'E' AND FAR_CP = 'P' ORDER BY FAR_FECHA_COMPRA, FAR_NUMOPER"
  ' Agregado 12052004 (Ody)
  If val(txt_cli.Text) <> 0 Then
    pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCLIE = " & Trim(txt_cli.Text) & " AND FAR_CODCIA = ? AND FAR_TIPMOV = 98 AND FAR_FECHA_PRO >= ? AND FAR_FECHA_PRO <= ? AND FAR_FBG = 'A' AND FAR_ESTADO <> 'E' AND FAR_CP = 'P'  ORDER BY FAR_NUMFAC,FAR_FECHA_PRO, FAR_NUMOPER"
  Else
    pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = 98 AND FAR_FECHA_PRO >= ? AND FAR_FECHA_PRO <= ? AND FAR_FBG = 'A' AND FAR_ESTADO <> 'E' AND FAR_CP = 'P'  ORDER BY FAR_NUMFAC,FAR_FECHA_PRO, FAR_NUMOPER"
  End If
  ' Fin Agregado
  WS_SIGNO = 1
End If
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = LK_FECHA_DIA
PS_REP02(2) = LK_FECHA_DIA
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


' el PS_REP1(0) ESTA MAS ABAJO
PS_REP02(0) = LK_CODCIA
'PS_REP02(1) = 10
PS_REP02(1) = wsFECHA1
PS_REP02(2) = wsFECHA2
DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep02.Requery
If llave_rep02.EOF Then ' Verifica si existen registros de facturas
  GoTo OTRO_DOCUMENTO
End If
FrmImp2.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.max = llave_rep02.RowCount

WFLAG = ""
SQ_OPER = 1
pu_cp = "P"
pu_codcia = LK_CODCIA
wnumfac = llave_rep02!far_numfac
wfecha = llave_rep02!far_fbg
wserie = llave_rep02!far_numser
wwwserie = llave_rep02!far_numser
wq_fecha = "01/01/1900"
xcuenta = 0
WFLAG = "A"
wflag_numfac = "A"
AWQ_DESCTOS = 0
AWQ_GASTOS = 0
AWQ_FLETES = 0
www_fbg = Trim(llave_rep02!far_fbg)

' Agregado : Declaración de variables
' Fecha : 13/04/2004

Dim strMonedaAct As String * 1
Dim strMoneda As String * 1
Dim dblTC As Double

strMoneda = llave_rep02!FAR_MONEDA

Do Until llave_rep02.EOF
    ' Agregado : Recuperar la moneda
    strMonedaAct = llave_rep02!FAR_MONEDA

'  If llave_rep02!FAR_numfac = 401 Then Stop
  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  If wfecha <> llave_rep02!far_fbg Then '
     If wflag_numfac = "A" Then
       GoSub IMPRI_FAC
     End If
     wflag_numfac = ""
     f1 = f1 + 1
     GoSub TOTAL_DIA
     t_valor_venta = t_valor_venta + S_VALOR_VENTA
     t_descto = t_descto + s_descto
     t_valor_igv = t_valor_igv + s_valor_igv
     t_valor_precio = t_valor_precio + S_VALOR_PRECIO
     wnumfac = llave_rep02!far_numfac
     wfecha = llave_rep02!far_fbg
     S_VALOR_VENTA = 0
     s_descto = 0
     s_valor_igv = 0
     S_VALOR_PRECIO = 0
  End If

  If wnumfac <> llave_rep02!far_numfac Then
     GoSub IMPRI_FAC
     strMoneda = strMonedaAct
     wflag_numfac = ""
     wnumfac = llave_rep02!far_numfac
  ElseIf val(wwwserie) <> val(llave_rep02!far_numser) And CDate(wq_fecha) = CDate(llave_rep02!FAR_fecha_compra) Then ' Se cambio far_fecha_compra por far_fecha_pro
    If Trim(www_fbg) <> Trim(llave_rep02!far_fbg) And wflag_numfac <> "" Then
      GoSub IMPRI_FAC
      wflag_numfac = ""
      wnumfac = llave_rep02!far_numfac
    End If
  ElseIf val(wwwserie) = val(llave_rep02!far_numser) And CDate(wq_fecha) = CDate(llave_rep02!FAR_fecha_compra) Then ' Se cambio far_fecha_compra por far_fecha_pro
    If Trim(www_fbg) <> Trim(llave_rep02!far_fbg) And wflag_numfac <> "" Then
      GoSub IMPRI_FAC
      wflag_numfac = ""
      wnumfac = llave_rep02!far_numfac
    End If
  End If
'  xl.Application.Visible = True
  'wq_fecha = Format(llave_rep02!FAR_fecha_compra, "dd/mm/yy")
  wq_fecha = Format(llave_rep02!FAR_fecha_compra, "dd/mm/yy")
  wq_fecha_compra = Format(llave_rep02!FAR_fecha_compra, "dd/mm/yy") ' Agregado 14/04/04
  wq_codclie = llave_rep02!far_codclie
  wq_codven = llave_rep02!FAR_CODVEN
  www_fbg = Trim(llave_rep02!far_fbg)
  wq_fbg = "'" & llave_rep02!far_cod_sunat ' Trim(llave_rep02!far_fbg)
  wq_docu = "'" & llave_rep02!far_numfac
  wq_serie = "'" & llave_rep02!far_numguia
  AWQ_TIPO_CAMBIO = 1
  FLAG_TC = 1


  If LK_MONEDA = "S" And llave_rep02!FAR_MONEDA = "D" Then
         FLAG_TC = 9999
  ElseIf LK_MONEDA = "D" And llave_rep02!FAR_MONEDA = "S" Then
         FLAG_TC = 8888  '1 / Far_Cost!FAR_TIPO_CAMBIO
  Else
   If par_llave!PAR_MONEDA_CON = "S" And llave_rep02!FAR_MONEDA = "D" Then
        FLAG_TC = 9999
   ElseIf par_llave!PAR_MONEDA_CON = "D" And llave_rep02!FAR_MONEDA = "D" Then
        FLAG_TC = 8888
    End If
  End If
  If FLAG_TC = 8888 Or FLAG_TC = 9999 Then
     PUB_CAL_INI = llave_rep02!FAR_fecha_compra 'llave_rep02!FAR_fecha_pro 'llave_rep02!FAR_fecha_compra 'Modificado
     PUB_CAL_FIN = llave_rep02!FAR_fecha_compra  'llave_rep02!FAR_fecha_pro 'llave_rep02!FAR_fecha_compra ' Modificado
     PUB_CODCIA = LK_CODCIA
     LEER_CAL_LLAVE
     If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 Then
       AWQ_TIPO_CAMBIO = 0
     Else
       AWQ_TIPO_CAMBIO = 1 'cal_llave!cal_tipo_cambio
     End If
     If AWQ_TIPO_CAMBIO <= 0 Then
         MsgBox "Definir Tipo de Cambios para el Periodo Actual. Dia : " & llave_rep02!FAR_fecha_compra & " (en el Calendario del Sistema)", 48, Pub_Titulo
         xl.DisplayAlerts = False
         xl.Cells(f1 + 1, 1) = "Falta Tipo de Cambio.... "
         GoTo CANCELA
         Exit Sub
     End If
     If FLAG_TC = 8888 Then AWQ_TIPO_CAMBIO = 1 / AWQ_TIPO_CAMBIO
  Else
    AWQ_TIPO_CAMBIO = 1
  End If

  wwwserie = llave_rep02!far_numser
  wq_serie = "'" & llave_rep02!FAR_NUMSER_C
  wserie = llave_rep02!FAR_NUMSER_C
  wq_docu = "'" & llave_rep02!FAR_NUMFAC_C
  wq_guia = "'" & llave_rep02!far_numguia
  www_tipmov = llave_rep02!FAR_TIPMOV
  wq_nombre = ""
  wq_bruto = (val(llave_rep02!FAR_BRUTO) - val(llave_rep02!FAR_TOT_DESCTO)) * WS_SIGNO
  wq_bruto = Format(wq_bruto * AWQ_TIPO_CAMBIO, "0.000")
  wq_desto = (val(llave_rep02!FAR_TOT_DESCTO) * WS_SIGNO) * AWQ_TIPO_CAMBIO
  wq_gastos = (val(llave_rep02!FAR_GASTOS) * WS_SIGNO) * AWQ_TIPO_CAMBIO
  wq_flete = (val(Nulo_Valor0(llave_rep02!FAR_TOT_FLETE)) * WS_SIGNO) * AWQ_TIPO_CAMBIO
  'wq_tot_descto =
  WQ_IMPTO = (Format(val(llave_rep02!far_IMPTO), "0.000") * WS_SIGNO) * AWQ_TIPO_CAMBIO
  WQ_TOTAL = (val(llave_rep02!FAR_BRUTO) + val(llave_rep02!far_IMPTO) - val(llave_rep02!FAR_TOT_DESCTO) + val(llave_rep02!FAR_GASTOS)) * WS_SIGNO
  WQ_TOTAL = wq_bruto + wq_gastos + WQ_IMPTO '(wq_total * AWQ_TIPO_CAMBIO)
  If wq_bruto = 0 Then WQ_TOTAL = 0
  wq_estado = llave_rep02!far_estado
  If wq_estado <> "E" Then
    If Left(UCase(llave_rep02!far_subtra), 1) <> "A" Then
     AWQ_COSTO_VENTA = AWQ_COSTO_VENTA + ((llave_rep02!FAR_COSPRO * llave_rep02!far_cantidad)) * WS_SIGNO
    End If
  End If
  wq_condi = llave_rep02!far_numguia
  var_secuencia = Nulo_Valor0(llave_rep02!FAR_NUM_LOTE)
  wflag_numfac = "A"
 WFLAG = "A"
 llave_rep02.MoveNext
Loop

If wflag_numfac = "A" Then
    GoSub IMPRI_FAC
    wflag_numfac = ""
End If
If WFLAG = "A" Then
    f1 = f1 + 1
    GoSub TOTAL_DIA
    t_valor_venta = t_valor_venta + S_VALOR_VENTA
    t_descto = t_descto + s_descto
    t_valor_igv = t_valor_igv + s_valor_igv
    t_valor_precio = t_valor_precio + S_VALOR_PRECIO
End If
'  If cheasiento.Value = 1 Then
'    FrmImp2.lblProceso.Caption = "Procesando Pase de Contabilidad . . . "
'    DoEvents
'    GoSub PASE_CONTAB
'  End If
  xcuenta = c1 + 1
  wranF = "A6:" & "K6"
  xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3
  If WCONTROL = 1 Then
   If cheasiento.Value = 1 Then
'    FrmImp2.lblProceso.Caption = "Procesando Pase de Contabilidad . . . "
'    DoEvents
'    GoSub PASE_CONTAB
   End If
  End If
OTRO_DOCUMENTO:
If WCONTROL >= 3 Then
Else
  GoTo NCREDITO
End If

MOSTRAR:
   If cheasiento.Value = 1 Then
    FrmImp2.lblproceso.Caption = "Procesando Pase de Contabilidad . . . "
    DoEvents
    GoSub PASE_CONTAB
   End If


  f1 = f1 + 2
  xl.Cells(f1, 1) = "Total General = "
  xl.Cells(f1, 12) = t_valor_venta
  xl.Cells(f1, 13) = t_descto
  xl.Cells(f1, 14) = t_valor_igv
  xl.Cells(f1, 15) = t_valor_precio


  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'" & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  'xl.Worksheets(1).Protect ws_clave   'quita proteccion GTS
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub

IMPRI_FAC:
     f1 = f1 + 1
     pu_codclie = wq_codclie
     LEER_CLI_LLAVE
     If Not cli_llave.EOF Then
         wq_codclie = cli_llave!cli_codclie 'xl.Cells(F1, 2)
         wq_nombre = Trim(cli_llave!CLI_NOMBRE)
         wq_ruc = Trim(cli_llave!cli_ruc_esposo)
         wcodcta = Trim(cli_llave!CLI_CUENTA_CONTAB)
         wcodcta2 = Trim(cli_llave!CLI_CUENTA_CONTAB2)
     End If
     xl.Cells(f1, 1) = "'" & Format(wq_fecha, "dd/mm/yyyy")
     xl.Cells(f1, 2) = wq_codclie
     xl.Cells(f1, 3) = wq_nombre
     xl.Cells(f1, 4) = wq_ruc
     xl.Cells(f1, 5) = wq_condi
     xl.Cells(f1, 6) = wq_fbg
     xl.Cells(f1, 7) = wq_guia
     xl.Cells(f1, 8) = wq_serie
     xl.Cells(f1, 9) = wq_docu
     xl.Cells(f1, 16) = www_tipmov
     xl.Cells(f1, 17) = var_secuencia
     If Trim(www_fbg) = "" Then
       xl.Cells(f1, 18) = "L"
     Else
       xl.Cells(f1, 18) = www_fbg
     End If
     If wq_estado = "E" Then
         xl.Cells(f1, 3) = "[ANULADO] "
     Else
         xl.Cells(f1, 3) = wq_nombre
     End If
     If wq_estado <> "E" Then
      If Left(cli_llave!CLI_CUENTA_CONTAB, 2) <> "12" Then
'       xl.Cells(f1, 9) = wq_bruto
'       xl.Cells(f1, 10) = Val(wq_desto)
'       xl.Cells(f1, 11) = Val(WQ_IMPTO)
'       xl.Cells(f1, 12) = Val(WQ_TOTAL)
'       xl.Cells(f1, 15) = Trim(wcodcta)
'       xl.Cells(f1, 14) = Trim(wcodcta2)

        ' Agregado : Obtener el tipo de Cambio
        dblTC = 1
        If strMoneda = "D" Then 'llave_rep02!FAR_MONEDA = "D" Then
            dblTC = JALAR(CDate(wq_fecha_compra))
            If dblTC <= 0 Then
                MsgBox "Falta Ingresar el Tipo de Cambio del día : " & Format(wq_fecha_compra, "dd/mm/yyyy"), 48, Pub_Titulo 'Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy")
                GoTo CANCELA
            End If
        End If
        
        ' Llenado de celdas
        xl.Cells(f1, 10) = wq_bruto
        xl.Cells(f1, 13) = val(wq_desto)
        xl.Cells(f1, 11) = IIf(dblTC = 1, "", FormatNumber(dblTC, 3)) 'Val(WQ_IMPTO)
        xl.Cells(f1, 12) = wq_bruto * dblTC 'Val(WQ_TOTAL)

        xl.Cells(f1, 14) = val(WQ_IMPTO) * dblTC
        
        xl.Cells(f1, 15) = (wq_bruto + val(WQ_IMPTO)) * dblTC ' Total
        
        xl.Cells(f1, 18) = Trim(wcodcta)
        xl.Cells(f1, 19) = Trim(wcodcta2)


       'xl.Cells(F1, 14) = Trim(cli_llave!CLI_CUENTA_CONTAB) ' Ya comentado
'       AWQ_BRUTO_ACT_FIJO = AWQ_BRUTO_ACT_FIJO + wq_bruto
'       AWQ_NETO_ACT_FIJO = AWQ_NETO_ACT_FIJO + Val(WQ_TOTAL)
'       AWQ_CTA_ACT_FIJO = Trim(cli_llave!CLI_CUENTA_CONTAB)
'       s_valor_igv = s_valor_igv + Val(WQ_IMPTO)
'       AWQ_IMPTO = AWQ_IMPTO + Val(WQ_IMPTO)
'       ' ACUMULA OTROS ********
'       S_VALOR_VENTA = S_VALOR_VENTA + wq_bruto
'       s_descto = s_descto + Val(wq_desto)
'       S_VALOR_PRECIO = S_VALOR_PRECIO + Val(WQ_TOTAL)
       
       
        ' Agregado :
        AWQ_BRUTO_ACT_FIJO = AWQ_BRUTO_ACT_FIJO + wq_bruto * dblTC
        AWQ_NETO_ACT_FIJO = AWQ_NETO_ACT_FIJO + (wq_bruto + val(WQ_IMPTO)) * dblTC
        AWQ_CTA_ACT_FIJO = Trim(cli_llave!CLI_CUENTA_CONTAB)
        s_valor_igv = s_valor_igv + (val(WQ_IMPTO) * dblTC)
        AWQ_IMPTO = AWQ_IMPTO + (val(WQ_IMPTO) * dblTC)
        ' ACUMULA OTROS ********
        S_VALOR_VENTA = S_VALOR_VENTA + wq_bruto * dblTC
        s_descto = s_descto + val(wq_desto)
        S_VALOR_PRECIO = S_VALOR_PRECIO + (wq_bruto + val(WQ_IMPTO)) * dblTC
       
      Else
       S_VALOR_VENTA = S_VALOR_VENTA + wq_bruto
       s_descto = s_descto + val(wq_desto)
       s_valor_igv = s_valor_igv + val(WQ_IMPTO)
       S_VALOR_PRECIO = S_VALOR_PRECIO + val(WQ_TOTAL)
       xl.Cells(f1, 8) = wq_bruto
       xl.Cells(f1, 9) = val(wq_desto)
       xl.Cells(f1, 10) = val(WQ_IMPTO)
       xl.Cells(f1, 11) = val(WQ_TOTAL)
       xl.Cells(f1, 14) = Trim(cli_llave!CLI_CUENTA_CONTAB)
       AWQ_IMPTO = AWQ_IMPTO + val(WQ_IMPTO)
       AWQ_NETO = AWQ_NETO + val(WQ_TOTAL)
       AWQ_BRUTO = AWQ_BRUTO + wq_bruto
       AWQ_DESCTOS = wq_desto
       AWQ_GASTOS = wq_gastos
       AWQ_FLETES = wq_flete
       If wq_condi = "CRED" Then
         xl.Cells(f1, 15) = -1
         AWQ_NETO_CRED = AWQ_NETO_CRED + val(WQ_TOTAL)
       Else
         xl.Cells(f1, 15) = 0
         AWQ_NETO_CONT = AWQ_NETO_CONT + val(WQ_TOTAL)
       End If
     End If
     End If
Return

CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\REGCOMPRA1.xls", 0, True, 4, WPAS, WPAS
  'xl.Workbooks.Open "Y:\ADMIN\STANDAR\REGCOMPRA.xls", 0, True, 4, WPAS, WPAS
Return

TOTAL_DIA:
  If wfecha = " " Then
    xl.Cells(f1, 1) = "Total Facturas x Mercaderia    = "
  ElseIf wfecha = "K" Then
    xl.Cells(f1, 1) = "Total Facturas Varias    = "
  ElseIf wfecha = "C" Then
    xl.Cells(f1, 1) = "Total N.Creditos  = "
  ElseIf wfecha = "A" Then
    xl.Cells(f1, 1) = "Total N.Debito    = "
  End If
  xl.Cells(f1, 3) = ""
  xl.Cells(f1, 4) = ""
  xl.Cells(f1, 8) = ""
  xl.Cells(f1, 9) = ""
  xl.Cells(f1, 12) = S_VALOR_VENTA
  xl.Cells(f1, 13) = s_descto
  xl.Cells(f1, 14) = s_valor_igv
  xl.Cells(f1, 15) = S_VALOR_PRECIO
Return


PASE_CONTAB:
Dim wscodcia As String * 2
Dim wcta As String
Dim wcta_clientes As Currency
Dim PS_CONTAB1 As rdoQuery
Dim contab_llave As rdoResultset
Dim ws_nro_voucher As Integer
Dim ws_nro_sec As Integer
Dim ws_glosa As String
Dim wsq_fecha As String
Dim wdh As String * 1
Dim wcta_clientes2 As Currency
Dim wcta_clientes3 As Currency
Dim wcta2 As String
Dim wcta_bruto As Currency
Dim wcta_total_bruto As Currency
Dim wcta_igv As Currency
Dim wcta_neto As Currency
Dim wfg1 As String * 1
Dim wwwf As String * 1
Dim flag_pasaigv As String * 1


Dim wsq_fecha2
wsq_fecha = Format(cop_llave!cop_fecha_proceso, "yyyy/mm/dd")
wsq_fecha2 = Format(cop_llave!cop_fecha_proceso2, "yyyy/mm/dd")

pub_cadena = "DELETE COMOV  WHERE COV_FLAG_AUTOMATICA = '4'  AND COV_CODUSU = '" & LK_CODCIA & "' AND (COV_FECHA_VOUCHER >=  ' " & wsq_fecha & "' AND COV_FECHA_VOUCHER <=  ' " & wsq_fecha2 & "') "
CN.Execute pub_cadena, rdExecDirect

FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.max = 7
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1

If LK_EMP_PTO = "A" Then
PSCOV_VOUCHER(0) = "00"
Else
PSCOV_VOUCHER(0) = LK_CODCIA
End If
PSCOV_VOUCHER(1) = cop_llave!cop_fecha_proceso
PSCOV_VOUCHER(2) = cop_llave!cop_fecha_proceso2
cov_voucher.Requery
If cov_voucher.EOF Then
 ws_nro_voucher = 0
Else
 cov_voucher.MoveLast
 ws_nro_voucher = cov_voucher!COV_NRO_VOUCHER
End If
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
wran1 = "A" & 6 & ":R" & f1
'xl.Application.Visible = True
xl.APPLICATION.Worksheets("Hoja1").Range(wran1).Sort Key1:=xl.APPLICATION.Worksheets("Hoja1").Range("R6"), Order1:=xlDescending, Key1:=xl.APPLICATION.Worksheets("Hoja1").Range("R6")
fila = 6
wcta = Trim(xl.Cells(fila, 15))
wcta2 = Trim(xl.Cells(fila, 14))
wtipo_mov = Trim(xl.Cells(fila, 16))
wcta_clientes = 0
wcta_clientes2 = 0
wcta_clientes3 = 0
wcta_igv = 0
wcta_bruto = 0
wcta_neto = 0
wcta_total_bruto = 0
If LK_EMP_PTO = "A" Then
 ws_glosa = "Registro de Compra - " & Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
Else
  ws_glosa = "Registro de Compra"
End If

wdh = "H"
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
wfg1 = "X"
wwwf = ""
flag_pasaigv = ""
For fila = 6 To f1
  If Trim(xl.Cells(fila, 15)) = "" Then
    GoTo OTRITO
  Else
    If wfg1 = "X" Then
      wfg1 = ""
      wcta = Trim(xl.Cells(fila, 15))
'      xl.Application.Visible = True
      wcta2 = Trim(xl.Cells(fila, 14))
    End If
  End If
  If Trim(xl.Cells(fila, 16)) = "" Then GoTo OTRITO
  If Trim(xl.Cells(fila, 15)) = "" Then GoTo OTRITO
  If xl.Cells(fila, 16) = 98 Or xl.Cells(fila, 16) = 97 Then
      If wwwf = "A" Then
          wdh = "H"
          GoSub Graba
          wdh = "D"
          wcta = wcta2
          wcta_clientes = wcta_bruto
          GoSub Graba
          'JALA  COD. TRANS
          SQ_OPER = 1
          PUB_SECUENCIA = 1
          PUB_CODTRA = 1401
          PUB_CODCIA = LK_CODCIA
          LEER_CNT_LLAVE
          If cnt_llave.EOF Then
           MsgBox "Error de Dato de Transaccion , Consulte a su Proveedor.", 48, Pub_Titulo
           End
          End If
          FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
          ' PARA LA CUENTA DE  IGV DE REG. DE COMPRAS
          FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
          If Trim(cnt_llave!CNT_CTA1) <> "" Then
              wcta = cnt_llave!CNT_CTA1 'Impuesto
              wdh = cnt_llave!CNT_DH1
              wcta_clientes = wcta_igv
              GoSub Graba
         End If
         flag_pasaigv = "A"
       End If
        ws_nro_voucher = ws_nro_voucher + 1
        If val(xl.Cells(fila, 16)) = 97 Then
            ws_glosa = "Nota de Credito/Prov: " + Trim(xl.Cells(fila, 7)) + " - " + xl.Cells(fila, 8)
        ElseIf val(xl.Cells(fila, 16)) = 98 Then
            ws_glosa = "Nota de Debito/Prov: " + Trim(xl.Cells(fila, 7)) + " - " + xl.Cells(fila, 8)
        End If
        SQ_OPER = 1
        PUB_SECUENCIA = val(xl.Cells(fila, 17))
        PUB_CODTRA = 2410
        PUB_CODCIA = LK_CODCIA
        LEER_CNT_LLAVE
        If cnt_llave.EOF Then
          MsgBox "Error de Dato de Transaccion , Consulte a su Proveedor.", 48, Pub_Titulo
        End If
        If Trim(cnt_llave!CNT_CTA1) <> "" Then
          wcta = cnt_llave!CNT_CTA1  'Cuenta del Proveedor
          If Trim(wcta) = "CLIENTES" Then
            wcta = Trim(xl.Cells(fila, 15))
          End If
          wdh = cnt_llave!CNT_DH1
          wcta_clientes = Abs(val(xl.Cells(fila, 12)))
          GoSub Graba
        End If
        If Trim(cnt_llave!CNT_CTA2) <> "" Then
          wcta = cnt_llave!CNT_CTA2  'cuenta de IGV fijo
          wdh = cnt_llave!CNT_DH2
          wcta_clientes = Abs(val(xl.Cells(fila, 11)))
          GoSub Graba
        End If
        If Trim(cnt_llave!CNT_CTA3) <> "" Then
          wcta = cnt_llave!CNT_CTA3  'cuenta de bruto fijo
          wdh = cnt_llave!CNT_DH3
          wcta_clientes = Abs(val(xl.Cells(fila, 9)))
          GoSub Graba
        End If
        ws_nro_voucher = ws_nro_voucher + 1
        wwwf = ""
        GoTo OTRITO
  End If
  If wcta <> Trim(xl.Cells(fila, 15)) Then
     wdh = "H"
     GoSub Graba
     wdh = "D"
     wcta = wcta2
     wcta_clientes = wcta_bruto
     GoSub Graba
     wcta_clientes = 0
     wcta_bruto = 0
     wcta = Trim(xl.Cells(fila, 15))
     wcta2 = Trim(xl.Cells(fila, 14))
     wcta_clientes = wcta_clientes + val(xl.Cells(fila, 12))
     wcta_igv = wcta_igv + val(xl.Cells(fila, 11))
     wcta_bruto = wcta_bruto + val(xl.Cells(fila, 9))
     If val(wtipo_mov) = 20 And Left(Trim(wcta2), 2) <> "33" Then
       wcta_total_bruto = wcta_total_bruto + val(xl.Cells(fila, 9))
     End If
     wwwf = "A"
  Else
     wcta_clientes = wcta_clientes + val(xl.Cells(fila, 12))
     wcta_igv = wcta_igv + val(xl.Cells(fila, 11))
     wcta_bruto = wcta_bruto + val(xl.Cells(fila, 9))
     wtipo_mov = Trim(xl.Cells(fila, 16))
     'If Val(wtipo_mov) <> 98 And Val(wtipo_mov) <> 97 And Val(wtipo_mov) = 20 And Left(Trim(wcta2), 2) <> "33" Then
     wcta_total_bruto = wcta_total_bruto + val(xl.Cells(fila, 9))
     'Else
     '  ws_nro_voucher = ws_nro_voucher + 1
     'End If
     wwwf = "A"
  End If
OTRITO:
Next fila

If wwwf = "A" Then
   wdh = "H"
   GoSub Graba
   wdh = "D"
   wcta = wcta2
   wcta_clientes = wcta_bruto
   GoSub Graba
   'JALA  COD. TRANS
   SQ_OPER = 1
   PUB_SECUENCIA = 1
   PUB_CODTRA = 1401
   PUB_CODCIA = LK_CODCIA
   LEER_CNT_LLAVE
   If cnt_llave.EOF Then
    MsgBox "Error de Dato de Transaccion , Consulte a su Proveedor.", 48, Pub_Titulo
    End
   End If
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   ' PARA LA CUENTA DE  IGV DE REG. DE COMPRAS
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   If Trim(cnt_llave!CNT_CTA1) <> "" Then
       wcta = cnt_llave!CNT_CTA1 'Impuesto
       wdh = cnt_llave!CNT_DH1
       wcta_clientes = wcta_igv
       GoSub Graba
  End If
End If

'ws_nro_voucher = ws_nro_voucher + 1
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
'If wwwf = "A" Then
'  wdh = "H"
'  GoSub GRABA
'  wdh = "D"
'  wcta = wcta2
'  wcta_clientes = wcta_bruto
'  GoSub GRABA
'End If
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
'ws_nro_voucher = ws_nro_voucher + 1
'If Trim(cnt_llave!CNT_CTA2) <> "" Then
' wcta = cnt_llave!CNT_CTA2 'Valor Compra
' wdh = cnt_llave!CNT_DH2
' wcta_clientes = wcta_total_bruto
' GoSub GRABA
'End If
'If Trim(cnt_llave!CNT_CTA3) <> "" Then
' wcta = cnt_llave!CNT_CTA3 'Valor compra
' wdh = cnt_llave!CNT_DH3
' wcta_clientes = wcta_total_bruto
' GoSub GRABA
'End If


Return

Graba:
'    If Trim(wcta) = "40101" Then Stop
     ws_nro_sec = ws_nro_sec + 1
     cov_voucher.AddNew
     If LK_EMP_PTO = "A" Then
       cov_voucher!COV_CODCIA = "00"
     Else
       cov_voucher!COV_CODCIA = LK_CODCIA
     End If
     cov_voucher!COV_FECHA_VOUCHER = cop_llave!cop_fecha_proceso2
     cov_voucher!COV_NRO_MOV = ws_nro_sec
     cov_voucher!COV_NRO_VOUCHER = ws_nro_voucher
     cov_voucher!COV_NUMTAB = 0
     cov_voucher!COV_CODCTA = wcta
     cov_voucher!COV_DH = wdh
     cov_voucher!COV_IMPORTE = wcta_clientes
     cov_voucher!COV_ESTADO = " "
     cov_voucher!COV_CODUSU = LK_CODCIA
     cov_voucher!cov_flag_automatica = "4"
     cov_voucher!COV_glosa = ws_glosa
     cov_voucher!COV_FECHA_doc = LK_FECHA_DIA
     cov_voucher.Update

Return



LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"

Return

FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
 GoTo CANCELA
 Resume Next

End Sub


Public Sub VENTAS_DEUDA()
'On Error GoTo FINTODO
Dim cant_doc As Currency
Dim WR_TC As Currency
Dim WW_VENTAS As Currency
Dim WS_FILA, WS_COL As Integer
Dim WW_VENTA_ANT As Currency
Dim WT_VENTA As Currency
Dim WS_MES, WS_ANO As Integer
Dim wsFECHA1 As Date
Dim TAB_MESES(24) As String
Dim TAB_COLS(24) As Integer
Dim JJ As Integer
Dim WW_VENTA As Currency
Dim wfecha
Dim wnumfac As Long
Dim WFBG As String * 1
Dim wserie As Integer
Dim wq_fecha As Date
Dim J As Integer
Dim WWFECHA1 As Date
Dim WW_FECHA2 As Date
Dim FFF As Integer
Dim WS_MES_ANT As Integer
Dim ws_clave
Dim LETRAS(55) As String * 2
Dim wRuta As String
Dim wmonto As Currency
Dim WSZONA As Integer
WSZONA = 0
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If IsDate(wsFECHA1) = False Then
MsgBox "Fecha Invalidad.", 48, Pub_Titulo
Exit Sub
End If
If DateDiff("m", wsFECHA1, LK_FECHA_DIA) > 12 Then
   MsgBox "No mas de 12 meses..."
   Exit Sub
End If
If DateDiff("m", wsFECHA1, LK_FECHA_DIA) < 0 Then
   MsgBox "Fecha no Valida..."
   Exit Sub
End If
If Trim(wsFECHA1) = "" Then
 Exit Sub
End If

If Trim(solozonas.Text) = "" Then
  MsgBox "Seleccione una Zona.", 48, Pub_Titulo
  solozonas.SetFocus
  Exit Sub
End If
WSZONA = val(Right(Trim(solozonas.Text), 6))
WS_MES_ANT = DatePart("m", txtCampo1.Text)
Dim i As Integer

If Not IsDate(wsFECHA1) Then
  MsgBox "Fecha Invalida ..", 48, Pub_Titulo
  Exit Sub
End If


Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents

pub_cadena = "SELECT CLI_NOMBRE, CLI_CODCIA, CLI_CODCLIE  FROM CLIENTES WHERE CLI_CODCIA = ?  AND CLI_ZONA_NEW = ?  ORDER BY  CLI_NOMBRE"
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

pub_cadena = "SELECT CAR_IMPORTE, CAR_CODCLIE, CAR_FECHA_INGR, CAR_MONEDA, CAR_FECHA_SUNAT  FROM CARTERA WHERE CAR_CODCLIE = ? AND  CAR_CODCIA=?  AND CAR_IMPORTE <> 0   AND CAR_CP = 'C' AND (CAR_TIPDOC = 'FA' OR  CAR_TIPDOC = 'CC')    ORDER BY  CAR_FECHA_INGR"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

pub_cadena = "SELECT FAR_MONEDA, FAR_FECHA_COMPRA, FAR_BRUTO, FAR_IMPTO, FAR_TOT_DESCTO , FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_ESTADO, FAR_FECHA  FROM FACART WHERE FAR_CODCLIE = ? AND  FAR_CODCIA=?  AND FAR_FECHA >= ? AND FAR_ESTADO <> 'E' AND FAR_ESTADO2 <> 'C'  AND FAR_CP = 'C'  ORDER BY  FAR_FECHA, FAR_FBG, FAR_NUMSER, FAR_NUMFAC "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)


FrmImp2.lblproceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
FrmImp2.lblproceso.Caption = "Procesando ... "
DoEvents

GoSub WEXCEL
GoSub LETRAS
i = 3
WW_FECHA2 = wsFECHA1
WWFECHA1 = LK_FECHA_DIA
J = 0
Do Until WWFECHA1 < wsFECHA1
  xl.Cells(4, i) = TAB_MESES(DatePart("m", wsFECHA1))
  
  WS_MES = DatePart("m", wsFECHA1)
  If WS_MES < WS_MES_ANT Then WS_MES = WS_MES + 12
  J = J + 3
  TAB_COLS(WS_MES) = J
 
  wsFECHA1 = DateAdd("m", 1, wsFECHA1)
  i = i + 3
Loop
J = i
xl.Cells(3, J) = "TOTAL "
xl.Cells(4, J) = "VENTA"
wranF = "A4" & ":" & Trim(LETRAS(J)) & "4"
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3

wranF = "A5" & ":" & Trim(LETRAS(J)) & "5"
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3


pub_cadena = ""
xcuenta = 0

FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0

PS_REP03(0) = LK_CODCIA
PS_REP03(1) = WSZONA
llave_rep03.Requery
If llave_rep03.EOF = True Then
   MsgBox "No Existen Cliente en Esta Zona... ", 48, Pub_Titulo
   Pantalla.Enabled = True
   cerrar.Enabled = True
   FrmImp2.lblproceso.Visible = False
   FrmImp2.ProgBar.Visible = False
   Exit Sub
End If
FrmImp2.ProgBar.max = llave_rep03.RowCount
WS_FILA = 4
Do Until llave_rep03.EOF
WT_VENTA = 0
PS_REP01(0) = llave_rep03!cli_codclie
PS_REP01(1) = LK_CODCIA
PS_REP01(2) = WW_FECHA2
llave_rep01.Requery
PS_REP02(0) = llave_rep03!cli_codclie
PS_REP02(1) = LK_CODCIA
llave_rep02.Requery
WS_COL = 3
If llave_rep01.EOF And llave_rep02.EOF Then GoTo OTRO

WS_FILA = WS_FILA + 1
xl.Cells(WS_FILA, 1) = Left(llave_rep03!CLI_NOMBRE, 25)
xl.Worksheets(1).Rows(WS_FILA).RowHeight = 17
'xl.Application.Visible = False
If llave_rep01.EOF Then GoTo SALTA
wnumfac = llave_rep01!far_numfac
WFBG = Trim(llave_rep01!far_fbg)
wserie = val(llave_rep01!far_numser)
wq_fecha = Format(llave_rep01!FAR_fecha, "dd/mm/yyyy")
WS_MES = DatePart("m", wq_fecha)
WS_ANO = DatePart("yyyy", wq_fecha)
If llave_rep01!FAR_MONEDA = "D" Then
   WR_TC = JALAR(llave_rep01!FAR_fecha_compra)
   If WR_TC <= 0 Then
      MsgBox "Falta Llenar Tipo de Cambio del día :" & Format(llave_rep01!FAR_fecha_compra, "dd/mm/yyyy"), 48, Pub_Titulo
      GoTo CANCELA
   End If
   WW_VENTA = WW_VENTA + ((val(llave_rep01!FAR_BRUTO) + val(llave_rep01!far_IMPTO) - val(llave_rep01!FAR_TOT_DESCTO)) * WR_TC)
Else
   WW_VENTA = val(llave_rep01!FAR_BRUTO) + val(llave_rep01!far_IMPTO) - val(llave_rep01!FAR_TOT_DESCTO) + WW_VENTA
End If
Do Until llave_rep01.EOF
  If WS_ANO = DatePart("yyyy", llave_rep01!FAR_fecha) And WS_MES = DatePart("m", llave_rep01!FAR_fecha) Then
  Else
         If WS_MES < WS_MES_ANT Then WS_MES = WS_MES + 12
         WS_COL = TAB_COLS(WS_MES)
         xl.Cells(WS_FILA, WS_COL) = WW_VENTA + val(xl.Cells(WS_FILA, WS_COL))
         WT_VENTA = WT_VENTA + WW_VENTA
         WW_VENTA = 0
  End If
  wq_fecha = Format(llave_rep01!FAR_fecha, "dd/mm/yyyy")
  If val(wserie) = val(llave_rep01!far_numser) And Trim(WFBG) = Trim(llave_rep01!far_fbg) And wnumfac = llave_rep01!far_numfac Then
  Else
    If llave_rep01!FAR_MONEDA = "D" Then
       WR_TC = JALAR(llave_rep01!FAR_fecha_compra)
       If WR_TC <= 0 Then
             MsgBox "Falta Llenar Tipo de Cambio del día :" & Format(llave_rep01!FAR_fecha_compra, "dd/mm/yyyy"), 48, Pub_Titulo
             GoTo CANCELA
       End If
       WW_VENTA = WW_VENTA + ((val(llave_rep01!FAR_BRUTO) + val(llave_rep01!far_IMPTO) - val(llave_rep01!FAR_TOT_DESCTO)) * WR_TC)
    Else
       WW_VENTA = val(llave_rep01!FAR_BRUTO) + val(llave_rep01!far_IMPTO) - val(llave_rep01!FAR_TOT_DESCTO) + WW_VENTA
    End If
  End If
  
  wnumfac = llave_rep01!far_numfac
  WFBG = Trim(llave_rep01!far_fbg)
  wserie = val(llave_rep01!far_numser)
  WS_MES = DatePart("m", wq_fecha)
  WS_ANO = DatePart("yyyy", wq_fecha)
  llave_rep01.MoveNext
Loop
         If WS_MES < WS_MES_ANT Then WS_MES = WS_MES + 12
         WS_COL = TAB_COLS(WS_MES)
         xl.Cells(WS_FILA, WS_COL) = WW_VENTA + val(xl.Cells(WS_FILA, WS_COL))
         WT_VENTA = WT_VENTA + WW_VENTA

         xl.Cells(WS_FILA, J) = WT_VENTA

SALTA:
cant_doc = 0
WW_VENTA = 0
WW_VENTA_ANT = 0
Do Until llave_rep02.EOF

  If llave_rep02!CAR_FECHA_INGR < WW_FECHA2 Then
      If llave_rep02!CAR_MONEDA = "D" Then
        WR_TC = JALAR(llave_rep02!CAR_FECHA_SUNAT)
        If WR_TC <= 0 Then
            MsgBox "Falta Llenar Tipo de Cambio del día :" & Format(llave_rep01!FAR_fecha_compra, "dd/mm/yyyy"), 48, Pub_Titulo
            GoTo CANCELA
        End If
        WW_VENTA_ANT = (val(llave_rep02!car_importe) * WR_TC) + WW_VENTA_ANT
      Else
        WW_VENTA_ANT = val(llave_rep02!car_importe) + WW_VENTA_ANT
      End If
      GoTo SIGUEME
  End If
  If WS_ANO = DatePart("yyyy", llave_rep02!CAR_FECHA_INGR) And WS_MES = DatePart("m", llave_rep02!CAR_FECHA_INGR) Then
  Else
         If WS_MES < WS_MES_ANT Then WS_MES = WS_MES + 12
         WS_COL = TAB_COLS(WS_MES) '+ 1
         xl.Cells(WS_FILA, WS_COL) = WW_VENTA + val(xl.Cells(WS_FILA, WS_COL))
         If val(xl.Cells(WS_FILA, WS_COL)) <> 0 Then
          xl.Cells(WS_FILA, WS_COL + 2) = "'" & Format(cant_doc, "0")
         End If
         WT_VENTA = WT_VENTA + WW_VENTA
         WW_VENTA = 0
         cant_doc = 0
  End If
  wq_fecha = Format(llave_rep02!CAR_FECHA_INGR, "dd/mm/yyyy")
  If llave_rep02!CAR_MONEDA = "D" Then
    WR_TC = JALAR(llave_rep02!CAR_FECHA_SUNAT)
    WW_VENTA = WW_VENTA + (val(llave_rep02!car_importe) * WR_TC)
  Else
    WW_VENTA = val(llave_rep02!car_importe) + WW_VENTA
  End If
  
  WS_MES = DatePart("m", wq_fecha)
  WS_ANO = DatePart("yyyy", wq_fecha)
  cant_doc = cant_doc + 1
SIGUEME:
  llave_rep02.MoveNext
Loop
         If WS_MES < WS_MES_ANT Then WS_MES = WS_MES + 12
         WS_COL = TAB_COLS(WS_MES) '+ 1
         xl.Cells(WS_FILA, WS_COL) = WW_VENTA + val(xl.Cells(WS_FILA, WS_COL))
         If val(xl.Cells(WS_FILA, WS_COL)) <> 0 Then
          xl.Cells(WS_FILA, WS_COL + 2) = "'" & Format(cant_doc, "0")
         End If
         WT_VENTA = WT_VENTA + WW_VENTA
         xl.Cells(WS_FILA, J) = WT_VENTA
         WW_VENTA = 0
         If WW_VENTA_ANT <> 0 Then xl.Cells(WS_FILA, 2) = WW_VENTA_ANT
'         xl.Application.Visible = True
otro2:

OTRO:
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1

llave_rep03.MoveNext
Loop
'xl.Application.Visible = True
  pub_mensaje = "DESEA CON ASTERISCOS  . ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbYes Then
     FrmImp2.ProgBar.Min = 0
     FrmImp2.ProgBar.max = 20
     FrmImp2.ProgBar.Value = 0
     For JJ = 4 To 60 Step 3
        FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
        If JJ > J Then GoTo PAS
         i = 5
         Do Until i > WS_FILA
            If val(xl.Cells(i, JJ - 1)) = 0 Then xl.Cells(i, JJ) = ""
            If val(xl.Cells(i, JJ - 1)) > 0 Then
             xl.Cells(i, JJ) = "* " & xl.Cells(i, JJ + 1)
            End If
            i = i + 1
         Loop
PAS:
     Next JJ
  End If


wranF = "A" & WS_FILA + 1 & ":" & Trim(LETRAS(J)) & WS_FILA + 1
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3

fila = 4
i = 3
Do Until i > J
  'xl.Worksheets(1).Rows(4).RowHeight = 15
  wran1 = Trim(LETRAS(fila))
  xl.Worksheets(1).Columns(wran1).ColumnWidth = 3
  wran1 = Trim(LETRAS(fila + 1))
  xl.Worksheets(1).Columns(wran1).ColumnWidth = 0
  wran1 = Trim(LETRAS(i)) & 5
  wran2 = Trim(LETRAS(i)) & WS_FILA
  wranF = Trim(LETRAS(i)) & WS_FILA + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  fila = fila + 3
  i = i + 3
  
Loop
'xl.Application.Visible = True
wranF = "A4" & ":" & Trim(LETRAS(i - 3)) & WS_FILA + 1
xl.Range(wranF).Borders.LineStyle = 1

xl.Cells(WS_FILA + 1, 1) = "TOTALES  S/. ="
xl.Worksheets(1).Rows(WS_FILA + 1).RowHeight = 18
FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
xl.Cells(2, 1) = Format(LK_FECHA_DIA, "dddd, dd mmmm yyyy")
xl.Cells(3, 1) = "R E L A C I O N   D E  V E N T A S   M E N S U A L E S  - ZONA : " & Left(solozonas.Text, 30)
xl.DisplayAlerts = False
xl.Worksheets(1).Protect ws_clave
xl.APPLICATION.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = False
FrmImp2.ProgBar.Visible = False
Set xl = Nothing
Screen.MousePointer = 0
FrmImp2.Pantalla.Enabled = True
FrmImp2.cerrar.Enabled = True
FrmImp2.Pantalla.Caption = "Por &Pantalla"
FrmImp2.lblproceso.Visible = False
Exit Sub

 

CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo .xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open (PUB_RUTA_OTRO) & "VENTAS_DEUDA.xls", 0, True, 4, WPAS, WPAS
Return





LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"
LETRAS(25) = "Y"
LETRAS(26) = "Z"
LETRAS(27) = "AA"
LETRAS(28) = "AB"
LETRAS(29) = "AC"
LETRAS(30) = "AD"
LETRAS(31) = "AE"
LETRAS(32) = "AF"
LETRAS(33) = "AG"
LETRAS(34) = "AH"
LETRAS(35) = "AI"
LETRAS(36) = "AJ"
LETRAS(37) = "AK"
LETRAS(38) = "AL"
LETRAS(39) = "AM"
LETRAS(40) = "AN"
LETRAS(41) = "AO"
LETRAS(42) = "AP"
LETRAS(43) = "AQ"
LETRAS(44) = "AR"
LETRAS(45) = "AS"
LETRAS(46) = "AT"
LETRAS(47) = "AU"
LETRAS(48) = "AV"
LETRAS(49) = "AW"
LETRAS(50) = "AX"



TAB_MESES(1) = "ENERO"
TAB_MESES(2) = "FEBRERO"
TAB_MESES(3) = "MARZO"
TAB_MESES(4) = "ABRIL"
TAB_MESES(5) = "MAYO"
TAB_MESES(6) = "JUNIO"
TAB_MESES(7) = "JULIO"
TAB_MESES(8) = "AGOSTO"
TAB_MESES(9) = "SETIEMBRE"
TAB_MESES(10) = "OCTUBRE"
TAB_MESES(11) = "NOVIEMBRE"
TAB_MESES(12) = "DICIEMBRE"


Return

FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
 GoTo CANCELA
 Resume Next
End Sub

Public Sub PRO_CO()
Dim ws_file As String
Dim WVALOR1
Dim WVALOR2
Dim WS_STANES As Integer
Dim WS_SALDO_CLI   As Currency
Dim WS_FILA_ULTIMA   As Integer
Dim WS_LLEVA_SALDO  As Currency
Dim PU_MONEDA As String * 1
Dim wtipdoc As String

If Not IsDate(txtCampo1.Text) Then
   MsgBox "FECHA INVALIDAD .."
  Exit Sub
End If
PUB_CP = "C"
If txtCampo1.Text = #1/1/1998# Then
  wtipdoc = InputBox("Tipo de documento : ", "", " ")
  
pub_cadena = "SELECT CAA_ESTADO  FROM CARACU WHERE CAA_CP=? AND CAA_CODCLIE = ? AND CAA_CODCIA=? AND CAA_TIPDOC=? AND CAA_ESTADO <> 'E' ORDER BY CAA_FECHA"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = "C"
PS_REP01(2) = LK_CODCIA

  
pub_cadena = "SELECT CLI_RUC_EMPRESA ,CLI_CODCLIE, CLI_CODCIA, CLI_ESTADO FROM CLIENTES WHERE CLI_CP = 'C'  AND CLI_CODCIA = ?  ORDER BY CLI_CODCLIE"
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP03(0) = LK_CODCIA
llave_rep03.Requery
Do Until llave_rep03.EOF
   PUB_IMPORTE = 0
   PS_REP01(1) = llave_rep03!cli_codclie
   PS_REP01(3) = wtipdoc
   llave_rep01.Requery
   If Not llave_rep01.EOF Then
        PUB_IMPORTE = 1
   End If
   If PUB_IMPORTE = 1 Then
       llave_rep03.Edit
       llave_rep03!CLI_estado = "X"
       llave_rep03.Update
   End If
   llave_rep03.MoveNext
Loop
  MsgBox "PROCESO TERMINADO...SALIR Y REINGRESAR..."
  Exit Sub
Else
pub_cadena = "SELECT CLI_RUC_EMPRESA ,CLI_CODCLIE, CLI_CODCIA, CLI_ESTADO FROM CLIENTES WHERE CLI_CP = 'C'  AND CLI_CODCIA = ? AND CLI_ESTADO = 'A'  ORDER BY CLI_CODCLIE"
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurValues)
End If

pub_cadena = "SELECT CAA_NUM_OPER FROM CARACU WHERE CAA_CP = 'C'  AND CAA_CODCIA = ? ORDER BY CAA_NUM_OPER DESC"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT CAR_CODCLIE FROM CARTERA WHERE CAR_CP = 'C'  AND CAR_CODCIA = ? AND CAR_CODCLIE = ?  ORDER BY CAR_CODCLIE"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PS_REP03(0) = LK_CODCIA
llave_rep03.Requery





If llave_rep03.EOF Then
   MsgBox " No hay Clientes Activados ...,Verificar ", 48, Pub_Titulo
   Exit Sub
End If

PUB_TIPDOC = Left(Combo1.Text, 2)
PU_MONEDA = Trim(Right(Combo1.Text, 1))
If PU_MONEDA <> "S" Then
   PU_MONEDA = "D"
End If

WS_LLEVA_SALDO = 0

WS_STANES = 1  'Val(InputBox("CUANTOS STANDS TIENE :", "INGRESAR", "1"))
If PUB_TIPDOC = "CO" Then
   WS_LLEVA_SALDO = 15
End If

WVALOR2 = InputBox("MONTO DE LA CUOTA  : ", "CUOTA", WS_LLEVA_SALDO)
If WVALOR2 = "" Then Exit Sub
If Not IsNumeric(WVALOR2) Then Exit Sub

WS_SALDO_CLI = WVALOR2

FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
FrmImp2.lblproceso.Visible = True
DoEvents
caa_histo.Requery

FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.max = llave_rep03.RowCount
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Visible = True
DoEvents
PS_REP01(0) = LK_CODCIA
llave_rep01.Requery
PUB_NUM_OPER_XXX = llave_rep01!CAA_NUM_OPER + 1
Do Until llave_rep03.EOF
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   DoEvents
  
SQ_OPER = 3
pu_codcia = LK_CODCIA
PUB_SERDOC = 0
pu_cp = PUB_CP
LEER_CAR_LLAVE
If car_menor.EOF = True Then
   PUB_NUMDOC = 1
Else
   PUB_NUMDOC = car_menor!car_NUMDOC + 1
End If
WS_STANES = val(llave_rep03!CLI_RUC_EMPRESA)
If WS_STANES = 0 Then WS_STANES = 1

If chestands.Value <> 1 Then
   WS_STANES = 1
End If

PUB_CODCLIE = llave_rep03!cli_codclie
PUB_CODVEN = 1
PUB_FECHA_VCTO = txtCampo1.Text
PUB_IMPORTE_AMORT = WS_SALDO_CLI * WS_STANES
PUB_NUMFAC_C = 1
GoSub GRABAR_TODO
llave_rep03.MoveNext
Loop
MsgBox "PROCESO TERMINADO SATISFACTORIAMENTE..", 48, Pub_Titulo
GoTo CANCELA

Exit Sub

GRABAR_TODO:
car_llave.AddNew
car_llave!CAR_CODCLIE = PUB_CODCLIE
car_llave!car_codcia = LK_CODCIA
car_llave!car_numguia = 0
car_llave!car_TIPDOC = PUB_TIPDOC
car_llave!CAR_cp = PUB_CP
car_llave!car_SERDOC = PUB_SERDOC
car_llave!car_NUMDOC = PUB_NUMDOC
car_llave!CAR_FECHA_INGR = LK_FECHA_DIA
car_llave!car_fecha_vcto = PUB_FECHA_VCTO
car_llave!car_fecha_vcto_orig = PUB_FECHA_VCTO
car_llave!CAR_SITUACION = 0
car_llave!CAR_COMISION = 0
car_llave!CAR_NUM_REN = 0
car_llave!car_concepto = "Cuota :" & Trim(Mid(Combo1.Text, 4, 25)) & ":" & Format(PUB_FECHA_VCTO, "mmm-yyyy")
car_llave!car_nombre_banco = " "
car_llave!car_NUM_CHEQUE = 0
car_llave!car_SIGNO_CAJA = 0
car_llave!car_numguia = 0

car_llave!CAR_IMP_INI = PUB_IMPORTE_AMORT 'xl.Cells(FILAX, 3)
car_llave!car_importe = PUB_IMPORTE_AMORT
car_llave!car_codtra = 0
car_llave!car_PRECIO = 0
car_llave!CAR_signo_car = 1
car_llave!car_NUMSER = 0
car_llave!car_NUMFAC = 0
car_llave!CAR_TIPMOV = 0
car_llave!car_FBG = ""
car_llave!CAR_codven = 0
car_llave!CAR_NUMSER_C = 0
car_llave!CAR_NUMFAC_C = PUB_NUMFAC_C ' xl.Cells(FILAX, 2)
car_llave!CAR_codban = 0
car_llave!CAR_MONEDA = PU_MONEDA

car_llave.Update


   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   pu_codclie = llave_rep03!cli_codclie
   LEER_CLI_LLAVE


cli_llave.Edit
cli_llave!cli_SALDO = Nulo_Valor0(cli_llave!cli_SALDO) + PUB_IMPORTE_AMORT
cli_llave.Update

caa_histo.AddNew
caa_histo!CAA_CODCLIE = PUB_CODCLIE
caa_histo!caa_codcia = LK_CODCIA
caa_histo!CAA_TIPDOC = PUB_TIPDOC
caa_histo!CAA_CP = PUB_CP
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
caa_histo!CAA_NUM_OPER = PUB_NUM_OPER_XXX
caa_histo!caa_serdoc = PUB_SERDOC
caa_histo!CAA_NUMDOC = PUB_NUMDOC
caa_histo!CAA_FECHA = LK_FECHA_DIA
caa_histo!CAA_FECHA_VCTO = PUB_FECHA_VCTO
caa_histo!caa_situacion = 0
caa_histo!caa_concepto = "Cuota :" & Trim(Mid(Combo1.Text, 4, 25)) & ":" & Format(PUB_FECHA_VCTO, "mmm-yyyy")
caa_histo!CAA_IMPORTE = PUB_IMPORTE_AMORT
caa_histo!CAA_TOTAL = PUB_IMPORTE_AMORT
caa_histo!CAA_SALDO = Nulo_Valor0(cli_llave!cli_SALDO)
caa_histo!caa_SALDO_car = PUB_IMPORTE_AMORT
caa_histo!CAA_SIGNO_CAJA = 0
caa_histo!CAA_SIGNO_CAJA_REAL = 0
caa_histo!CAA_SIGNO_CAR = 1
caa_histo!CAA_TIPMOV = 10
caa_histo!CAA_hora = Now
caa_histo!CAA_CODUSU = LK_CODUSU
caa_histo!CAA_ESTADO = "N"
caa_histo!CAA_NUMPLAN = 0
caa_histo!CAa_NUM_CHEQUE = ""
caa_histo!caa_numser_c = 0
caa_histo!caa_numfac_c = PUB_NUMFAC_C ' xl.Cells(FILAX, 2)xl.Cells(FILAX, 2)
caa_histo!CAa_numser = 0
caa_histo!CAa_numfac = 0
caa_histo!CAa_NOMBRE = cli_llave!CLI_NOMBRE
caa_histo!CAa_numGUIA = 0
caa_histo!CAA_FBG = ""
caa_histo!CAA_CODVEN = 0
caa_histo.Update

Return

CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  Screen.MousePointer = 0
  Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo :SALDO_CAR" & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Resume Next
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.APPLICATION.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
 
End Sub


Public Sub REPO_CAJA()
Dim ws_conta As Integer
Dim WS_NOMCLI As String * 20
Dim WS_NOMBAN As String * 20
Dim WS_SALDO As Currency
Dim WS_MONEDA As String * 1
Dim ww_moneda
Dim ws_mensaje
Dim WS_FBG, WS_LARGO, ws_nomche, ws_codclie
Dim WS_SALDO_ING As Currency
Dim WS_SALDO_SAL As Currency
Dim todos_cli

Dim wsFECHA1
Dim ws_codcia As String
Dim xcuenta As Integer
Dim docus As String
Dim docus2 As String
Dim wmonto As Currency
Dim WINGRESOS As Currency
Dim TOT_INGRESOS As Currency
Dim WSALDO_CAJA As Currency


If Right(txtFecha.Text, 2) = "__" Then
     wsFECHA1 = Left(txtFecha.Text, 8)
Else
     wsFECHA1 = Trim(txtFecha.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If


FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL

xl.Worksheets(1).Activate
WS_SALDO = 0
xcuenta = 0
f1 = 5

WS_MONEDA = Left(cmdmoneda.Text, 1)

If WS_MONEDA = "S" Then
   xl.Cells(4, 2) = "MONEDA:" & " SOLES"
ElseIf WS_MONEDA = "D" Then
   xl.Cells(4, 2) = "MONEDA:" & " DOLLARES"
End If

WSALDO_CAJA = 0
'PUB_FECHA = wsFECHA1
'SQ_OPER = 1
'pu_codcia = LK_CODCIA
'LEER_ALL_LLAVE
'If all_llave.EOF Then GoTo VAMOS
'FrmImp2.ProgBar.Min = 0
'FrmImp2.ProgBar.Max = all_llave.RowCount
'FrmImp2.ProgBar.Value = 0
f1 = f1 + 1
ws_conta = 0
'If WS_MONEDA = "S" Then
'     WS_SALDO = all_llave!ALL_IMPORTE
'Else
'     WS_SALDO = all_llave!ALL_IMPORTE_DOLL
'End If

VAMOS:
'MsgBox all_llave!all_codtra
WS_SALDO = 0
WSALDO_CAJA = 0
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
Else
  PUB_CODCIA = LK_CODCIA
End If


SQ_OPER = 1
LEER_PAR_LLAVE
WS_SALDO = par_llave!PAR_SALDO_CAJA_ayer
WSALDO_CAJA = WSALDO_CAJA + WS_SALDO
xl.Cells(f1, 2) = "Saldo Anterior:"
xl.Cells(f1, 5) = WS_SALDO
'all_llave.MoveNext
pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ? AND ALL_FBG = ?  AND ALL_TIPMOV = 10 AND ALL_SIGNO_CAJA = 1  AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_DIA, ALL_FBG, ALL_NUMSER, ALL_NUMFAC"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT PAR_NOMBRE FROM PARGEN WHERE PAR_CODCIA = ? "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)

xcuenta = 1
f1 = 6
f1 = f1 + 1
xl.Cells(f1, 1) = "MAS  : INGRESOS"
f1 = f1 + 1
xl.Cells(f1, 1) = 1
xl.Cells(f1, 2) = "VENTAS AL CONTADO:"
WINGRESOS = 0
TOT_INGRESOS = 0
For fila = 1 To 30
   ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
   If Trim(ws_codcia) = "00" Then GoTo OTRA_CIA
   If Trim(ws_codcia) = "" Then Exit For
    WINGRESOS = 0
    PS_REP01(0) = ws_codcia
    PS_REP01(1) = wsFECHA1
    PS_REP02(0) = ws_codcia
    llave_rep02.Requery
    f1 = f1 + 1
    xl.Cells(f1, 2) = Trim(llave_rep02!PAR_NOMBRE)
    ' SOLO PARA FACTURAS
    PS_REP01(2) = "F"
    llave_rep01.Requery
    wmonto = 0
    If Not llave_rep01.EOF Then
       docus = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE_AMORT
        docus2 = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
        llave_rep01.MoveNext
      Loop
      f1 = f1 + 1
      ' xl.Cells(f1, 3) = "F/. " & docus & " AL " & docus2
      xl.Cells(f1, 3) = "FACTURAS"
      xl.Cells(f1, 4) = wmonto
      WINGRESOS = WINGRESOS + wmonto
    End If
    'SOLO PARA BOLETAS
    PS_REP01(2) = "B"
    llave_rep01.Requery
    wmonto = 0
    If Not llave_rep01.EOF Then
       docus = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE_AMORT
        docus2 = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
        llave_rep01.MoveNext
      Loop
      f1 = f1 + 1
      ' xl.Cells(f1, 3) = "B/. " & docus & " AL " & docus2
      xl.Cells(f1, 3) = "BOLETAS"
      xl.Cells(f1, 4) = wmonto
      WINGRESOS = WINGRESOS + wmonto
    End If
    ' SOLO PARA GUIAS
    PS_REP01(2) = "G"
    llave_rep01.Requery
    wmonto = 0
    If Not llave_rep01.EOF Then
       docus = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE_AMORT
        docus2 = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
        llave_rep01.MoveNext
      Loop
      f1 = f1 + 1
      'xl.Cells(f1, 3) = "G/. " & docus & " AL " & docus2
      xl.Cells(f1, 3) = "GUIAS"
      xl.Cells(f1, 4) = wmonto
      WINGRESOS = WINGRESOS + wmonto
    End If
    ' SOLO PARA VENTAS ADMINISTRADORES P
    PS_REP01(2) = "P"
    llave_rep01.Requery
    wmonto = 0
    If Not llave_rep01.EOF Then
       docus = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE_AMORT
        docus2 = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
        llave_rep01.MoveNext
      Loop
      f1 = f1 + 1
      xl.Cells(f1, 3) = "P/. " & docus & " AL " & docus2
      xl.Cells(f1, 3) = "ADMINISTRACION"
      xl.Cells(f1, 4) = wmonto
      WINGRESOS = WINGRESOS + wmonto
    End If
      f1 = f1 + 1
      xl.Cells(f1, 4) = WINGRESOS
      TOT_INGRESOS = TOT_INGRESOS + WINGRESOS
OTRA_CIA:
   xcuenta = xcuenta + 2
 Next fila
 f1 = f1 + 1
 xl.Cells(f1, 5) = TOT_INGRESOS
 WSALDO_CAJA = WSALDO_CAJA + TOT_INGRESOS
 
' SOLO COBRANZA DE FA,LE,AD
 f1 = f1 + 1
 pub_cadena = "SELECT ALL_IMPORTE_AMORT, ALL_CODCLIE, ALL_CODCIA FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ?  AND ALL_TIPDOC = ? AND ALL_SIGNO_CAJA = 1 AND ALL_SIGNO_CAR = -1 AND ALL_IMPORTE_AMORT <> 0  AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122) ORDER BY  ALL_FECHA_DIA, ALL_NUMOPER"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

 f1 = f1 + 1
 xl.Cells(f1, 1) = 2
 xl.Cells(f1, 2) = "COBRANZA DE FACTURAS"
 WINGRESOS = 0
 TOT_INGRESOS = 0
 
 SQ_OPER = 1
 pu_cp = "C"
 xcuenta = 1
 For fila = 1 To 30
   ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
   If Trim(ws_codcia) = "00" Then GoTo OTRA_CIA2
   If Trim(ws_codcia) = "" Then Exit For
    f1 = f1 + 1
    PS_REP02(0) = ws_codcia
    llave_rep02.Requery
    xl.Cells(f1, 2) = Trim(llave_rep02!PAR_NOMBRE)
    WINGRESOS = 0
    PS_REP01(0) = ws_codcia
    PS_REP01(1) = wsFECHA1
    PS_REP01(2) = "FA"
    llave_rep01.Requery
    f1 = f1 + 1
    xl.Cells(f1, 2) = ".- FACTURAS "
    wmonto = 0
    If Not llave_rep01.EOF Then
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE_AMORT
        pu_codclie = llave_rep01!ALL_CODCLIE
        pu_codcia = llave_rep01!all_CODCIA
        LEER_CLI_LLAVE
        f1 = f1 + 1
        xl.Cells(f1, 3) = Trim(cli_llave!CLI_NOMBRE)
        xl.Cells(f1, 4) = val(llave_rep01!ALL_IMPORTE_AMORT)
        llave_rep01.MoveNext
      Loop
      WINGRESOS = WINGRESOS + wmonto
    End If
    f1 = f1 + 1
    xl.Cells(f1, 4) = WINGRESOS
    TOT_INGRESOS = TOT_INGRESOS + WINGRESOS
    
    WINGRESOS = 0
    PS_REP01(0) = ws_codcia
    PS_REP01(1) = wsFECHA1
    PS_REP01(2) = "AD"
    llave_rep01.Requery
    f1 = f1 + 1
    xl.Cells(f1, 2) = ".- COBRANZA DE ADMINISTRACION "
    wmonto = 0
    If Not llave_rep01.EOF Then
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE_AMORT
        pu_codclie = llave_rep01!ALL_CODCLIE
        pu_codcia = llave_rep01!all_CODCIA
        LEER_CLI_LLAVE
        f1 = f1 + 1
        xl.Cells(f1, 3) = Trim(cli_llave!CLI_NOMBRE)
        xl.Cells(f1, 4) = val(llave_rep01!ALL_IMPORTE_AMORT)
        llave_rep01.MoveNext
      Loop
      WINGRESOS = WINGRESOS + wmonto
    End If
    f1 = f1 + 1
    xl.Cells(f1, 4) = WINGRESOS
    
    TOT_INGRESOS = TOT_INGRESOS + WINGRESOS
OTRA_CIA2:
   xcuenta = xcuenta + 2
 Next fila
f1 = f1 + 1
xl.Cells(f1, 5) = TOT_INGRESOS
WSALDO_CAJA = WSALDO_CAJA + TOT_INGRESOS

' SOLO COBRANZA DE ORDINARIAS Y JUDICIALES
 f1 = f1 + 1
 
 pub_cadena = "SELECT ALL_TIPDOC, ALL_SITUACION, ALL_IMPORTE_AMORT, ALL_CODCIA FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ?  AND (ALL_TIPDOC = ? OR ALL_TIPDOC = ? OR ALL_TIPDOC = ? OR ALL_TIPDOC = ? ) AND  ALL_SIGNO_CAJA = 1 AND ALL_SIGNO_CAR = -1  AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_DIA, ALL_NUMOPER"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

 f1 = f1 + 1
 xl.Cells(f1, 1) = 3
 xl.Cells(f1, 2) = "COBRANZA ORDINARIA"
 
 WINGRESOS = 0
 TOT_INGRESOS = 0

xcuenta = 1
For fila = 1 To 30
   ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
   If Trim(ws_codcia) = "00" Then GoTo OTRA_CIA3
   If Trim(ws_codcia) = "" Then Exit For
    WINGRESOS = 0
    PS_REP01(0) = ws_codcia
    PS_REP01(1) = wsFECHA1
    PS_REP01(2) = "LE"
    PS_REP01(3) = "PA"
    PS_REP01(4) = "GJ"
    PS_REP01(5) = "IN"
    PS_REP02(0) = ws_codcia
    llave_rep02.Requery
    f1 = f1 + 1
    xl.Cells(f1, 2) = Trim(llave_rep02!PAR_NOMBRE)
    llave_rep01.Requery
    wmonto = 0
    If Not llave_rep01.EOF Then
      Do Until llave_rep01.EOF
      If llave_rep01!ALL_TIPDOC = "LE" And llave_rep01!ALL_SITUACION <> "P" Then
      Else
        wmonto = wmonto + llave_rep01!ALL_IMPORTE_AMORT
       End If
        llave_rep01.MoveNext
      Loop
      WINGRESOS = WINGRESOS + wmonto
    End If
    f1 = f1 + 1
    xl.Cells(f1, 4) = val(wmonto)
    TOT_INGRESOS = TOT_INGRESOS + WINGRESOS
OTRA_CIA3:
   xcuenta = xcuenta + 2
 Next fila
f1 = f1 + 1
xl.Cells(f1, 5) = TOT_INGRESOS
WSALDO_CAJA = WSALDO_CAJA + TOT_INGRESOS

f1 = f1 + 1

'COBRANZA JUDICIAL
 pub_cadena = "SELECT ALL_TIPDOC, ALL_SITUACION, ALL_IMPORTE_AMORT, ALL_CODCIA FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ?  AND (ALL_TIPDOC = ? OR ALL_TIPDOC = ? OR ALL_TIPDOC = ? ) AND  (ALL_SITUACION = ? OR ALL_SITUACION = ? ) AND ALL_SIGNO_CAJA = 1 AND ALL_SIGNO_CAR = -1  AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_DIA, ALL_NUMOPER"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

f1 = f1 + 1
 xl.Cells(f1, 1) = 3
 xl.Cells(f1, 2) = "COBRANZA JUDICIAL"
 
 WINGRESOS = 0
 TOT_INGRESOS = 0

xcuenta = 1
For fila = 1 To 30
   ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
   If Trim(ws_codcia) = "00" Then GoTo OTRA_CIAJ
   If Trim(ws_codcia) = "" Then Exit For
    WINGRESOS = 0
    PS_REP01(0) = ws_codcia
    PS_REP01(1) = wsFECHA1
    PS_REP01(2) = "LE"
    PS_REP01(3) = "GJ"
    PS_REP01(4) = "GJ"
    PS_REP01(5) = "J"
    PS_REP01(6) = "J"
    PS_REP02(0) = ws_codcia
    llave_rep02.Requery
    f1 = f1 + 1
    xl.Cells(f1, 2) = Trim(llave_rep02!PAR_NOMBRE)
    llave_rep01.Requery
    wmonto = 0
    If Not llave_rep01.EOF Then
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE_AMORT
        llave_rep01.MoveNext
      Loop
      WINGRESOS = WINGRESOS + wmonto
    End If
    f1 = f1 + 1
    xl.Cells(f1, 4) = val(wmonto)
    TOT_INGRESOS = TOT_INGRESOS + WINGRESOS
OTRA_CIAJ:
   xcuenta = xcuenta + 2
 Next fila
f1 = f1 + 1
xl.Cells(f1, 5) = TOT_INGRESOS
WSALDO_CAJA = WSALDO_CAJA + TOT_INGRESOS
 
' SOLO INGRESOS VARIOS
f1 = f1 + 1
 
 pub_cadena = "SELECT ALL_IMPORTE FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ?  AND ALL_CODTRA = 5350 AND ALL_SIGNO_CAJA = 1  AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_DIA, ALL_NUMOPER"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

 f1 = f1 + 1
 xl.Cells(f1, 1) = 4
 xl.Cells(f1, 2) = "INGRESOS VARIOS"
 
 WINGRESOS = 0
 TOT_INGRESOS = 0
xcuenta = 1
For fila = 1 To 30
   ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
   If Trim(ws_codcia) = "00" Then GoTo OTRA_CIA4
   If Trim(ws_codcia) = "" Then Exit For
    WINGRESOS = 0
    PS_REP01(0) = ws_codcia
    PS_REP01(1) = wsFECHA1
    PS_REP02(0) = ws_codcia
    llave_rep02.Requery
    f1 = f1 + 1
    xl.Cells(f1, 2) = Trim(llave_rep02!PAR_NOMBRE)
    llave_rep01.Requery
    wmonto = 0
    If Not llave_rep01.EOF Then
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE
        llave_rep01.MoveNext
      Loop
      WINGRESOS = WINGRESOS + wmonto
    End If
    f1 = f1 + 1
    xl.Cells(f1, 4) = val(wmonto)
    TOT_INGRESOS = TOT_INGRESOS + WINGRESOS
OTRA_CIA4:
   xcuenta = xcuenta + 2
 Next fila
f1 = f1 + 1
xl.Cells(f1, 5) = TOT_INGRESOS
WSALDO_CAJA = WSALDO_CAJA + TOT_INGRESOS
' SOLO depositos a bancos
f1 = f1 + 1
 
 pub_cadena = "SELECT ALL_IMPORTE FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ?  AND (ALL_CODTRA = 5310 ) AND ALL_SIGNO_CAJA = -1  AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_DIA, ALL_NUMOPER"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

 f1 = f1 + 1
 xl.Cells(f1, 1) = 5
 xl.Cells(f1, 2) = "MENOS : EGRESOS "
 f1 = f1 + 1
 xl.Cells(f1, 2) = "DEPOSITOS "
 
 WINGRESOS = 0
 TOT_INGRESOS = 0
xcuenta = 1
For fila = 1 To 30
   ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
   If Trim(ws_codcia) <> "00" Then GoTo OTRA_CIA5
   If Trim(ws_codcia) = "" Then Exit For
    WINGRESOS = 0
    PS_REP01(0) = ws_codcia
    PS_REP01(1) = wsFECHA1
    PS_REP02(0) = ws_codcia
    llave_rep02.Requery
    f1 = f1 + 1
    xl.Cells(f1, 2) = Trim(llave_rep02!PAR_NOMBRE)
    llave_rep01.Requery
    wmonto = 0
    If Not llave_rep01.EOF Then
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE
        llave_rep01.MoveNext
      Loop
      WINGRESOS = WINGRESOS + wmonto
    End If
    f1 = f1 + 1
    xl.Cells(f1, 4) = val(wmonto)
    TOT_INGRESOS = TOT_INGRESOS + WINGRESOS
OTRA_CIA5:
   xcuenta = xcuenta + 2
 Next fila
f1 = f1 + 1
xl.Cells(f1, 5) = TOT_INGRESOS
WSALDO_CAJA = WSALDO_CAJA - TOT_INGRESOS '(DEPOSITOS)
f1 = f1 + 1
f1 = f1 + 1
xl.Cells(f1, 3) = "SALDO DE CAJA = "
xl.Cells(f1, 5) = WSALDO_CAJA
If LK_EMP = "DPV" Then
  PUB_CODCIA = "00"
  SQ_OPER = 1
  LEER_PAR_LLAVE
  par_llave.Edit
  par_llave!PAR_SALDO_CAJA_HOY = WSALDO_CAJA
  par_llave.Update
  PUB_CODCIA = LK_CODCIA
  SQ_OPER = 1
  LEER_PAR_LLAVE
End If


DoEvents
FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
DoEvents
xcuenta = 1
xl.Cells(2, 2) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
xl.Cells(2, 5) = "'" & Format(wsFECHA1, "dd mmm yyyy")
xl.DisplayAlerts = False
xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
xl.Worksheets("Hoja1").Protect PUB_CLAVE
xl.APPLICATION.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = False
FrmImp2.ProgBar.Visible = False
Set xl = Nothing
Screen.MousePointer = 0
FrmImp2.Pantalla.Enabled = True
FrmImp2.Pantalla.Caption = "Por &Pantalla"
FrmImp2.lblproceso.Visible = False

Exit Sub


WEXCEL:
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FrmImp2.lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\OFFICE\CAJA.xls", 0, True, 4, WPAS, WPAS
Return

Exit Sub

'LLENA_VALOR:
'For I = QJ To WDIF
'  If I >= 3 Then
'    xl.Cells(F1 + 7, I) = Format(LOC_VALOR, "0.000")
'  End If
'Next I
'Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
Exit Sub

End Sub
Public Sub ULTIMAS_VENTAS()
Dim VAR_ULT As String
'On Error GoTo FINTODO
Dim cod
Dim Wche  As Integer
Dim Modo1 As String
Dim AWQ_NETO_ACT_FIJO As Currency
Dim J As Integer
Dim ww_por1 As Currency
Dim WW_FECHA As Date
Dim ww_por2 As Currency
Dim TAB_MESES(12)
Dim AWQ_CTA_ACT_FIJO As String
Dim WW_BRUTO As Currency
Dim WW_TOT_DESCTO As Currency
Dim WW_IMPTO As Currency

Dim WWFECHA1 As Date
Dim WW_FECHA2 As Date
Dim WCONTROL As Integer
Dim WW_VENTA As Currency
Dim wcodcia As String * 2
Dim WFBG As String * 1
Dim WW_VENTAS As Currency
Dim WS_FILA, WS_COL As Integer
Dim WS_MES, WS_ANO As Integer
Dim wfecha
Dim FFF As Integer
Dim WS_MES_ANT As Integer
Dim ws_clave
Dim LETRAS(24) As String * 1
Dim wRuta As String
Dim wmonto As Currency
Dim wcodclie As Currency
Dim valor_venta As Currency
Dim descto As Currency
Dim valor_igv As Currency
Dim valor_precio As Currency
Dim S_VALOR_VENTA As Currency
Dim s_descto As Currency
Dim s_valor_igv As Currency
Dim S_VALOR_PRECIO As Currency

Dim t_valor_venta As Currency
Dim tT_VALOR_VENTA As Currency
Dim t_valor_igv As Currency
Dim t_valor_precio As Currency

Dim wnumfac As Long
Dim WFLAG As String * 1
Dim wsFECHA1, wsFECHA2
Dim xcuenta As Integer
Dim wq_fecha, wq_codclie, wq_codven, wq_docu, wq_nombre, wq_bruto, wq_gastos, wq_desto, wq_flete, wq_fbg, wq_serie
Dim wq_tot_descto, WQ_IMPTO, WQ_TOTAL, wq_estado, wq_condi
Dim wq_NUMFAC
Dim wflag_numfac
Dim wserie As Integer
Dim AWQ_BRUTO As Currency
Dim AWQ_DESCTOS As Currency
Dim AWQ_GASTOS As Currency
Dim AWQ_FLETES As Currency
Dim AWQ_IMPTO As Currency
Dim AWQ_NETO As Currency
Dim AWQ_NETO_CRED  As Currency
Dim AWQ_NETO_CONT   As Currency
Dim AWQ_COSTO_VENTA As Currency
Dim WS_SIGNO As Integer

Dim TOTAL_SOLES As Currency
Dim TOTAL_UM As Double
Dim TOTAL_LITROS  As Double

Dim TOTAL_CLIENTE_SOLES As Currency
Dim TOTAL_CLIENTE_UM As Double
Dim TOTAL_CLIENTE_LITROS  As Double

VAR_ULT = InputBox("Cuantas Ultimas Compras desea Mostrar : ", "Ultimas Compras", "3")
If VAR_ULT = "" Then
 Exit Sub
End If


cod = 35
Pantalla.Enabled = False
cerrar.Enabled = False
MM:

GoSub WEXCEL
i = 7
pub_cadena = ""
xcuenta = 0

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
WCONTROL = 0
f1 = 5  'Fila Inicial
t_valor_venta = 0
WS_FILA = 4
wmonto = 0
wcodclie = 0
valor_venta = 0
descto = 0
valor_igv = 0
valor_precio = 0
wnumfac = 0
WFLAG = ""
xcuenta = 0
wflag_numfac = ""
wserie = 0
AWQ_BRUTO = 0
AWQ_DESCTOS = 0
AWQ_GASTOS = 0
AWQ_FLETES = 0
AWQ_IMPTO = 0
AWQ_NETO = 0
AWQ_NETO_CRED = 0
AWQ_NETO_CONT = 0
AWQ_COSTO_VENTA = 0

S_VALOR_VENTA = 0
s_descto = 0
s_valor_igv = 0
S_VALOR_PRECIO = 0


Wche = 0
Modo1 = "TAB_NUMTAB = "
For fila = 0 To zonas.ListCount - 1
  zonas.ListIndex = fila
  If zonas.Selected(fila) Then
    Wche = 1
    Modo1 = Modo1 + str(val(Right(zonas.Text, 5))) + " OR TAB_NUMTAB = "
  End If
Next fila
If Wche <> 0 Then
 If Trim(Right(Modo1, 2)) = "=" Then
   Modo1 = Left(Modo1, Len(Modo1) - 16)
 End If
 pub_cadena = "SELECT TAB_NUMTAB, TAB_NOMLARGO FROM TABLAS WHERE TAB_CODCIA = '00' AND TAB_TIPREG = " & cod & " AND (" & Modo1 & ") ORDER BY TAB_NUMTAB "
Else
 pub_cadena = "SELECT TAB_NUMTAB, TAB_NOMLARGO FROM TABLAS WHERE TAB_CODCIA = '00' AND TAB_TIPREG = " & cod & "   ORDER BY TAB_NUMTAB "
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
llave_rep01.Requery
If llave_rep01.EOF Then
   MsgBox "No ha determinado Zonas.", 48, Pub_Titulo
   Exit Sub
End If

pub_cadena = "SELECT CLI_NOMBRE, CLI_CODCLIE, CLI_CASA_DIREC, CLI_CASA_NUM  FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_ZONA_NEW  = ? "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


'DistinctRow

pub_cadena = "SELECT FAR_FBG, FAR_NUMFAC, FAR_BRUTO, FAR_IMPTO,SUM(FAR_CANTIDAD/FAR_EQUIV) AS UM ,SUM((FAR_CANTIDAD/FAR_EQUIV)*FAR_LITRO) AS LITROS , " & _
 " FAR_TOT_DESCTO, FAR_TIPMOV, FAR_CODCIA, FAR_CP, FAR_ESTADO, FAR_FECHA  " & _
 " FROM FACART " & _
 " WHERE FAR_TIPMOV = 10 AND  FAR_CODCIA=?  AND FAR_CODCLIE = ?   AND FAR_CP = 'C' AND FAR_ESTADO <> 'E' AND FAR_MONEDA = '" & Left(cmdmoneda.Text, 1) & "' " & _
 " GROUP BY FAR_FBG, FAR_NUMFAC, FAR_BRUTO, FAR_IMPTO , FAR_TOT_DESCTO, FAR_TIPMOV, FAR_CODCIA, FAR_CP, FAR_ESTADO, FAR_FECHA " & _
 " ORDER BY  FAR_FECHA DESC "
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
PS_REP03(0) = 0
PS_REP03(1) = 0
PS_REP03.MaxRows = val(VAR_ULT)
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenForwardOnly)

DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
FrmImp2.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0

WFLAG = ""
SQ_OPER = 1
pu_cp = "C"
pu_codcia = LK_CODCIA
xcuenta = 0
WFLAG = "A"
wflag_numfac = "A"
AWQ_DESCTOS = 0
AWQ_GASTOS = 0
AWQ_FLETES = 0

Do Until llave_rep01.EOF
 PS_REP02(0) = LK_CODCIA
 PS_REP02(1) = llave_rep01!TAB_NUMTAB
 WS_FILA = WS_FILA + 1
 xl.Cells(WS_FILA, 1) = llave_rep01!tab_NOMLARGO
 xl.Cells(WS_FILA, 1).Font.Bold = True
    'xl.Range(wranF).Font.Bold = True
    'xl.Range(wranF).Font.Name = "Arial"
    'xl.Range(wranF).Font.Size = 11
               
 llave_rep02.Requery
 If llave_rep02.RowCount <> 0 Then FrmImp2.ProgBar.max = llave_rep02.RowCount
 Do Until llave_rep02.EOF
    PS_REP03(1) = llave_rep02!cli_codclie
    PS_REP03(0) = LK_CODCIA
    llave_rep03.Requery
    If llave_rep03.EOF Then GoTo OTRO
    WS_FILA = WS_FILA + 1
    xl.Cells(WS_FILA, 1) = "'   " & llave_rep02!CLI_NOMBRE
    xl.Cells(WS_FILA, 2) = llave_rep02!CLI_CASA_DIREC
    xl.Cells(WS_FILA, 3) = llave_rep02!CLI_CASA_NUM
    Do Until llave_rep03.EOF
      WW_BRUTO = val(llave_rep03!FAR_BRUTO)
      WW_IMPTO = val(llave_rep03!far_IMPTO)
      WW_TOT_DESCTO = val(llave_rep03!FAR_TOT_DESCTO)
      WW_FECHA = llave_rep03!FAR_fecha
      xl.Cells(WS_FILA, 4) = WW_FECHA
      
      xl.Cells(WS_FILA, 5) = llave_rep03!far_numfac
      xl.Cells(WS_FILA, 6) = WW_BRUTO + WW_IMPTO + WW_TOT_DESCTO
      xl.Cells(WS_FILA, 7) = Format(llave_rep03!UM, "##,##0.000")
      xl.Cells(WS_FILA, 8) = Format(llave_rep03!LITROS, "##,##0.000")
      
      TOTAL_SOLES = TOTAL_SOLES + (WW_BRUTO + WW_IMPTO + WW_TOT_DESCTO)
      TOTAL_UM = TOTAL_UM + Format(llave_rep03!UM, "##,##0.000")
      TOTAL_LITROS = TOTAL_LITROS + Format(llave_rep03!LITROS, "##,##0.000")
      
      llave_rep03.MoveNext
      
      WS_FILA = WS_FILA + 1
      
    Loop
    xl.Cells(WS_FILA, 2) = "TOTAL CLIENTE"
    xl.Cells(WS_FILA, 6) = TOTAL_SOLES
    xl.Cells(WS_FILA, 7) = TOTAL_UM
    xl.Cells(WS_FILA, 8) = TOTAL_LITROS
    
    TOTAL_CLIENTE_SOLES = TOTAL_CLIENTE_SOLES + TOTAL_SOLES
    TOTAL_CLIENTE_UM = TOTAL_CLIENTE_UM + TOTAL_UM
    TOTAL_CLIENTE_LITROS = TOTAL_CLIENTE_LITROS + TOTAL_LITROS
    
    WS_FILA = WS_FILA + 1
    FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
OTRO:
  llave_rep02.MoveNext
 Loop
  WS_FILA = WS_FILA + 1
  xl.Cells(WS_FILA, 2) = "TOTAL GENERAL"
  xl.Cells(WS_FILA, 6) = TOTAL_CLIENTE_SOLES
  xl.Cells(WS_FILA, 7) = TOTAL_CLIENTE_UM
  xl.Cells(WS_FILA, 8) = TOTAL_CLIENTE_LITROS
  llave_rep01.MoveNext
Loop

xl.Cells(1, 6) = LK_FECHA_DIA
xl.Cells(2, 6) = Time
        
xl.Visible = True
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  'xl.Cells(3, 5) = LK_FECHA_DIA
  If Left(cmdmoneda.Text, 1) = "S" Then
   xl.Cells(4, 5) = "MONTO S/."
  Else
   xl.Cells(4, 5) = "MONTO US$"
  End If
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub

 

CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\STANDAR\ULTIMAS.xls", 0, True, 4, WPAS, WPAS
Return

FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
 GoTo CANCELA
 Resume Next
End Sub

Public Sub REPO_CAJA_DET()
FrmImp2.Pantalla.Enabled = False
Dim WIMPORTE As Currency
Dim wnumfac As Currency
Dim wusu As String * 1
Dim ww_fbg
Dim ww_numser
Dim ww_numfac
Dim ww_numini
Dim ww_serini
Dim ww_numfin
Dim WMONTO_SOLES As Currency
Dim WMONTO_DOLAR As Currency
Dim wsFECHA1
Dim WDOLAR_CRED As Currency
Dim WSOLES_CRED As Currency
Dim TT_TOTAL_SOLES As Currency
Dim TT_TOTAL_DOLAR  As Currency
Dim WSOLES_DOLAR As Currency
Dim TT_DOLAR_SOLES As Currency
Dim WSOLES_DOLAR_CRED As Currency
Dim WSOLES_DOLAR_TOT As Currency
Dim SOLES_TOTAL_COBRA As Currency
Dim DOLAR_TOTAL_COBRA As Currency
Dim TOTAL_COBRA_SD As Currency
Dim WS_TIPO_C As Currency

If Right(txtFecha.Text, 2) = "__" Then
     wsFECHA1 = Left(txtFecha.Text, 8)
Else
     wsFECHA1 = Trim(txtFecha.Text)
End If
If Not IsDate(wsFECHA1) Then
 FrmImp2.Pantalla.Enabled = True
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If


FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
f1 = 5
pub_mensaje = "Imprimir según su Usuario...?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbYes Then
  wusu = "A"
Else
  wusu = " "
End If

If wusu = "A" Then
   pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? " & _
   "OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10 " & _
   "AND ALL_CODUSU = ? AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 " & _
   "OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
Else
   pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND " & _
   "ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10   AND all_flag_ext <> 'E' AND " & _
   "(ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
End If

Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = 0
If wusu = "A" Then PS_REP01(3) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

'pub_cadena = "SELECT PAR_NOMBRE FROM PARGEN WHERE PAR_CODCIA = ? "
'Set PS_REP02 = CN.CreateQuery("", pub_cadena)
'Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
  PS_REP01(0) = "01"
  PS_REP01(1) = "02"
Else
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If


llave_rep01.Requery

If llave_rep01.EOF Then
  MsgBox "No Existe Movimientos", 48, Pub_Titulo
  GoTo CANCELA
  Exit Sub
End If
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.max = llave_rep01.RowCount
f1 = 6
ww_fbg = llave_rep01!ALL_FBG
ww_numser = llave_rep01!ALL_NUMSER
ww_serini = llave_rep01!ALL_NUMSER
ww_numfac = llave_rep01!all_numfac
ww_numini = llave_rep01!all_numfac
ww_numfin = llave_rep01!all_numfac
wnumfac = llave_rep01!all_numfac
Do Until llave_rep01.EOF
    FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
    If ww_fbg <> llave_rep01!ALL_FBG Then
        GoSub IMP_LINEA
        ww_fbg = llave_rep01!ALL_FBG
        ww_numser = llave_rep01!ALL_NUMSER
        ww_serini = llave_rep01!ALL_NUMSER
        ww_numfac = llave_rep01!all_numfac
        ww_numini = llave_rep01!all_numfac
        wnumfac = llave_rep01!all_numfac
    End If
    If ww_numser <> llave_rep01!ALL_NUMSER Then
        GoSub IMP_LINEA
        ww_numser = llave_rep01!ALL_NUMSER
        ww_serini = llave_rep01!ALL_NUMSER
        ww_numfac = llave_rep01!all_numfac
        ww_numini = llave_rep01!all_numfac
        wnumfac = llave_rep01!all_numfac
    End If
    If wnumfac <> llave_rep01!all_numfac Then
      
   '   MsgBox "Falta el Nro. :" & wnumfac
   '   wnumfac = llave_rep01!all_numfac
    End If
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
        WMONTO_SOLES = WMONTO_SOLES + val(llave_rep01!ALL_IMPORTE_AMORT)
'        MsgBox Val(llave_rep01!ALL_IMPORTE_AMORT)
        WSOLES_DOLAR = WSOLES_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
        If llave_rep01!ALL_SIGNO_CAJA = 0 Then
          WSOLES_CRED = WSOLES_CRED + val(llave_rep01!ALL_IMPORTE_AMORT)
          WSOLES_DOLAR_CRED = WSOLES_DOLAR_CRED + val(llave_rep01!ALL_IMPORTE_AMORT)
        End If
    Else
        WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
        WMONTO_DOLAR = WMONTO_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
        WSOLES_DOLAR = WSOLES_DOLAR + redondea(val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
        If llave_rep01!ALL_SIGNO_CAJA = 0 Then
          WDOLAR_CRED = WDOLAR_CRED + val(llave_rep01!ALL_IMPORTE_AMORT)
          WSOLES_DOLAR_CRED = WSOLES_DOLAR_CRED + redondea(val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
        End If
        
    End If
    ww_numfin = llave_rep01!all_numfac
    wnumfac = wnumfac + 1
    llave_rep01.MoveNext
Loop
GoSub IMP_LINEA
 TT_TOTAL_SOLES = 0
 TT_TOTAL_DOLAR = 0
 TT_DOLAR_SOLES = 0
 f1 = f1 + 1
 xl.Cells(f1 + 1, 4) = "TOTAL VENTA = "
 xl.Cells(f1 + 1, 7) = "S/."
 wran1 = "H" & 6
 wran2 = "H" & f1
 wranF = "H" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_SOLES = TT_TOTAL_SOLES + val(xl.Range(wranF))
 xl.Cells(f1 + 1, 9) = "US$."
 wran1 = "J" & 6
 wran2 = "J" & f1
 wranF = "J" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_DOLAR = TT_TOTAL_DOLAR + val(xl.Range(wranF))
 xl.Cells(f1 + 1, 11) = "S/."
 wran1 = "L" & 6
 wran2 = "L" & f1
 wranF = "L" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_DOLAR_SOLES = TT_DOLAR_SOLES + val(xl.Range(wranF))
 
 wranF = "A" & f1 + 1 & ":L" & f1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 wranF = "A" & f1 + 2 & ":L" & f1 + 2
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
f1 = f1 + 4
xl.Cells(f1, 1) = "(-) CREDITOS "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = WSOLES_CRED
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = WDOLAR_CRED
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = WSOLES_DOLAR_CRED

wranF = "A" & f1 + 1 & ":L" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
f1 = f1 + 3
xl.Cells(f1, 1) = "TOTAL EFECTIVO = "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = TT_TOTAL_SOLES - WSOLES_CRED
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = TT_TOTAL_DOLAR - WDOLAR_CRED
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = TT_DOLAR_SOLES - WSOLES_DOLAR_CRED

wranF = "A" & f1 + 1 & ":L" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1

' SELECT PARA COBRANZAS
pub_cadena = "SELECT ALL_MONEDA_CAJA , ALL_FECHA_SUNAT, ALL_MONEDA_CLI, ALL_IMPORTE_AMORT, ALL_TIPO_CAMBIO  FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV <> 10 AND ALL_SIGNO_CAJA = 1 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
 PS_REP01(0) = "01"
 PS_REP01(1) = "02"
Else
 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
llave_rep01.Requery
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Min = 0
If Not llave_rep01.EOF Then FrmImp2.ProgBar.max = llave_rep01.RowCount
Do Until llave_rep01.EOF
 FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
 If llave_rep01!ALL_MONEDA_CAJA = "S" Then
    WIMPORTE = llave_rep01!ALL_IMPORTE_AMORT
    If llave_rep01!ALL_MONEDA_CLI = "D" Then
       WIMPORTE = redondea(WIMPORTE * llave_rep01!ALL_TIPO_CAMBIO)
    End If
    SOLES_TOTAL_COBRA = SOLES_TOTAL_COBRA + WIMPORTE
    TOTAL_COBRA_SD = TOTAL_COBRA_SD + WIMPORTE
 Else
    WIMPORTE = llave_rep01!ALL_IMPORTE_AMORT
    If llave_rep01!ALL_MONEDA_CLI = "S" Then
       WIMPORTE = redondea(WIMPORTE / llave_rep01!ALL_TIPO_CAMBIO)
    End If
    WS_TIPO_C = llave_rep01!ALL_TIPO_CAMBIO
    DOLAR_TOTAL_COBRA = DOLAR_TOTAL_COBRA + WIMPORTE
    TOTAL_COBRA_SD = TOTAL_COBRA_SD + val(Format((val(WIMPORTE) * WS_TIPO_C), "0.000")) 'Val(llave_rep01!ALL_TIPO_CAMBIO))
 End If
  
llave_rep01.MoveNext
Loop


f1 = f1 + 3
xl.Cells(f1, 1) = "'(+)COBRANZAS "
f1 = f1 + 1
xl.Cells(f1, 1) = "'Ventas al Contado"
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = TT_TOTAL_SOLES - WSOLES_CRED
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = TT_TOTAL_DOLAR - WDOLAR_CRED
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = TT_DOLAR_SOLES - WSOLES_DOLAR_CRED
f1 = f1 + 1
xl.Cells(f1, 1) = "'Cobranza del Dia"
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = SOLES_TOTAL_COBRA
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = DOLAR_TOTAL_COBRA
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = TOTAL_COBRA_SD

f1 = f1 + 2
xl.Cells(f1, 1) = "'TOTAL DIA"
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = (TT_TOTAL_SOLES - WSOLES_CRED) + SOLES_TOTAL_COBRA
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = (TT_TOTAL_DOLAR - WDOLAR_CRED) + DOLAR_TOTAL_COBRA
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = (TT_DOLAR_SOLES - WSOLES_DOLAR_CRED) + TOTAL_COBRA_SD


 

 
 

 

    
DoEvents
FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
DoEvents
xcuenta = 1
xl.APPLICATION.Visible = True
xl.Cells(2, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
If checia.Value = 1 Then
  xl.Cells(2, 1) = "Almacenes 01 - 02 "
End If
xl.Cells(1, 1) = "IMPRESION: " & Now
xl.Cells(4, 10) = "FECHA:"
xl.Cells(4, 11) = "'" & Format(wsFECHA1, "dd mmm yyyy")
xl.DisplayAlerts = False
xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
xl.Worksheets("Hoja1").Protect PUB_CLAVE
xl.APPLICATION.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = False
FrmImp2.ProgBar.Visible = False
Set xl = Nothing
Screen.MousePointer = 0
FrmImp2.Pantalla.Enabled = True
FrmImp2.Pantalla.Caption = "Por &Pantalla"
FrmImp2.lblproceso.Visible = False

Exit Sub

IMP_LINEA:
f1 = f1 + 1
If ww_fbg = "F" Then
 xl.Cells(f1, 1) = "FACTURAS"
ElseIf ww_fbg = "B" Then
 xl.Cells(f1, 1) = "BOLETAS"
ElseIf ww_fbg = "G" Then
 xl.Cells(f1, 1) = "GUIAS"
Else
 xl.Cells(f1, 1) = "PENDIENTES"
End If
xl.Cells(f1, 2) = "Nº "
xl.Cells(f1, 3) = ww_serini
xl.Cells(f1, 4) = ww_numini
xl.Cells(f1, 5) = "'al Nº"
xl.Cells(f1, 6) = ww_numfin
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = Format(WMONTO_SOLES, "0.000")
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = Format(WMONTO_DOLAR, "0.000")
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = Format(WSOLES_DOLAR, "0.000")

WMONTO_SOLES = 0
WMONTO_DOLAR = 0
WSOLES_DOLAR = 0
Return



WEXCEL:
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FrmImp2.lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
  DoEvents
  WPAS = "131296"
  'WPAS = PUB_RUTA_OTRO + "CAJA_DET.xls"
  'DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open PUB_RUTA_OTRO & "\CAJA_DET.xls", 0, True, 4, WPAS, WPAS

  'xl.Workbooks.Open , "C:\ADMIN\HERTISA\CAJA_DET.XLS", 0, True, 4, PUB_CLAVE, PUB_CLAVE
Return

Exit Sub

'LLENA_VALOR:
'For I = QJ To WDIF
'  If I >= 3 Then
'    xl.Cells(F1 + 7, I) = Format(LOC_VALOR, "0.000")
'  End If
'Next I
'Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
Exit Sub

End Sub

Public Function JALAR(wfecha As Date) As Currency
PUB_CAL_INI = wfecha
PUB_CAL_FIN = wfecha
pu_codcia = LK_CODCIA
PUB_CODCIA = LK_CODCIA
SQ_OPER = 1
LEER_CAL_LLAVE
If cal_llave.EOF Then
  JALAR = 0
  Exit Function
End If
If IsNull(cal_llave!cal_tipo_cambio) Then
  JALAR = 0
  Exit Function
End If
JALAR = cal_llave!cal_tipo_cambio

End Function
Public Sub SECUENCIA()
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset
Dim wsFECHA1, wsFECHA2 As Date
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If

GoSub WEXCEL
pub_cadena = ""
xcuenta = 0

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
f1 = 5  'Fila Inicial

pub_cadena = "SELECT DISTINCT FAR_FBG, FAR_NUMSER, FAR_NUMFAC FROM FACART WHERE  FAR_CODCIA = ?  AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?  AND FAR_TIPMOV = 10  ORDER BY FAR_FBG, FAR_NUMSER, FAR_NUMFAC "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = 0
PS_REP02(2) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = wsFECHA1
PS_REP02(2) = wsFECHA2




DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep02.Requery
If llave_rep02.EOF Then
  GoTo CANCELA
End If
FrmImp2.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.max = llave_rep02.RowCount

PUB_FBG = llave_rep02!far_fbg
PUB_NUMFAC = llave_rep02!far_numfac
PUB_NUMSER = llave_rep02!far_numser

Do Until llave_rep02.EOF
   ''If llave_rep02.AbsolutePosition = 2 Then Stop

repite:
      If PUB_FBG = llave_rep02!far_fbg And PUB_NUMSER = val(llave_rep02!far_numser) And PUB_NUMFAC <> llave_rep02!far_numfac Then
         PUB_NUMFAC = PUB_NUMFAC + 1
      End If
      If PUB_FBG = llave_rep02!far_fbg And PUB_NUMSER = val(llave_rep02!far_numser) Then
         If PUB_NUMFAC <> llave_rep02!far_numfac Then
             f1 = f1 + 1
             xl.Cells(f1, 1) = PUB_FBG & "/. " & PUB_NUMSER & " - " & PUB_NUMFAC
             If Abs(PUB_NUMFAC - val(llave_rep02!far_numfac)) < 5 Then
                GoTo repite
             Else
             xl.Cells(f1, 1) = "FUERA DE ORDEN " & PUB_FBG & "/. " & PUB_NUMSER & " - " & llave_rep02!far_numfac
             End If
         End If
      Else
         PUB_NUMFAC = llave_rep02!far_numfac
         PUB_NUMSER = llave_rep02!far_numser
         PUB_FBG = llave_rep02!far_fbg
         'GoTo repite
      End If

  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  
 llave_rep02.MoveNext
Loop


MOSTRAR:



  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'" & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub


CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ""
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\RESU_VENTA_DIA.xls", 0, True, 4, WPAS, WPAS
Return





FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
 GoTo CANCELA
 Resume Next
End Sub

Public Sub RESU_VENTA_DIA()
Dim qver_onlyCont As Integer
Dim ts_suma  As Currency
Dim ts_suma_bruto  As Currency
Dim ts_suma_igv  As Currency
Dim ts_codcta As String
Dim Fres As Integer
Dim w_exo As Currency
Dim ws_codcia As String * 2
Dim FILAS As Integer
Dim nn, m_ind As Integer
Dim AWQ_NETO_ACT_FIJO As Currency
Dim AWQ_CTA_ACT_FIJO As String
Dim AWQ_BRUTO_ACT_FIJO As Currency
Dim WCONTROL As Integer
FILAS = 5
Dim WQ
Dim wfecha
Dim ws_clave
Dim LETRAS(24) As String * 1
Dim wRuta As String
Dim wmonto As Currency
Dim wcodclie As Currency
Dim valor_venta As Currency
Dim descto As Currency
Dim valor_igv As Currency
Dim valor_precio As Currency
Dim S_VALOR_VENTA As Currency
Dim s_descto As Currency
Dim s_valor_igv As Currency
Dim S_VALOR_PRECIO As Currency

Dim t_valor_venta As Currency
Dim t_descto As Currency
Dim t_valor_igv As Currency
Dim t_valor_precio As Currency

Dim wnumfac As Currency
Dim WFLAG As String * 1
Dim wsFECHA1, wsFECHA2
Dim xcuenta As Integer
Dim wq_fecha, wq_codclie, wq_codven, wq_docu, wq_nombre, wq_bruto, wq_gastos, wq_desto, wq_flete, wq_fbg, wq_serie
Dim wq_tot_descto, WQ_IMPTO, WQ_TOTAL, wq_estado, wq_condi
Dim wq_ruc
Dim wflag_numfac
Dim wserie As String * 3
Dim AWQ_BRUTO As Currency
Dim AWQ_DESCTOS As Currency
Dim AWQ_GASTOS As Currency
Dim AWQ_FLETES As Currency
Dim AWQ_IMPTO As Currency
Dim AWQ_NETO As Currency
Dim AWQ_NETO_CRED  As Currency
Dim AWQ_NETO_CONT   As Currency
Dim AWQ_COSTO_VENTA As Currency
Dim WS_SIGNO As Integer
Dim ws_tc As Currency
Dim wsTexto  As String
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) <> cop_llave!cop_fecha_proceso Then cheasiento.Value = 0
If CDate(wsFECHA2) <> cop_llave!cop_fecha_proceso2 Then cheasiento.Value = 0
If chepasa.Value = 1 Then
  pub_mensaje = "<Advertencia> El pase de la información es por cada Compañia. Continuar...?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
      Azul2 txtCampo1, txtCampo1
      GoTo CANCELA
  End If
  qver_onlyCont = CHE_BLOQ_MES(2)
  If qver_onlyCont = 1 Then
      MsgBox "Periodo Activo esta Cerrado. No procede.", 48, Pub_Titulo
      chepasa.Value = 0
      GoTo CANCELA
  End If
  If (cop_llave!cop_fecha_proceso = CDate(wsFECHA1)) And (cop_llave!cop_fecha_proceso2 = CDate(wsFECHA2)) Then
  Else
      MsgBox "Usted. a marcado la opción: Pasar la Información al Periodo Contable. " & Chr(13) & Chr(13) & "Las Fechas ingresadas son distintas a la del Periodo Contable Activo. Verificar...", 48, Pub_Titulo
      Azul2 txtCampo1, txtCampo1
      GoTo CANCELA
  End If
  If qver_onlyCont = 9 Then ' hay Información en OnlyCont. Confirmar.
      pub_mensaje = "Usted. a marcado la opción: Pasar la Información al Periodo Contable. " & Chr(13) & Chr(13) & "Existe Voucher en el Periodo Contable Activo. " & Chr(13) & Chr(13) & "< Desea adicionar este Nuevo Asiento de Voucher de todas Maneras >...?"
      Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
      If Pub_Respuesta = vbNo Then
        Azul2 txtCampo1, txtCampo1
        GoTo CANCELA
      End If
  End If
End If



GoSub WEXCEL
pub_cadena = ""
xcuenta = 0

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
WCONTROL = 0
f1 = 5  'Fila Inicial
t_valor_venta = 0
t_descto = 0
t_valor_igv = 0
t_valor_precio = 0

'NCREDITO: ' empieza

'WCONTROL = WCONTROL + 1
wmonto = 0
wcodclie = 0
valor_venta = 0
descto = 0
valor_igv = 0
valor_precio = 0
wnumfac = 0
WFLAG = ""
xcuenta = 0
wflag_numfac = ""
wserie = ""
Dim SS_VALOR_VENTA As Currency
Dim SS_VALOR_PRECIO As Currency
Dim SS_VALOR_IGV As Currency



Dim TD_F As String * 1
Dim TD_B As String * 1
Dim TD_N As String * 1
Dim TD_D As String * 1


TD_F = " "
TD_B = " "
TD_N = " "
TD_D = " "
wsTexto = "TIPO: "
For fila = 0 To 3
 txttp.ListIndex = fila
 If fila = 0 And txttp.Selected(fila) Then
   TD_F = "F"
   wsTexto = wsTexto + "- FACT."
 End If
 If fila = 1 And txttp.Selected(fila) Then
   TD_B = "B"
   wsTexto = wsTexto + "- BOL."
 End If
 If fila = 2 And txttp.Selected(fila) Then
   TD_N = "N"
   wsTexto = wsTexto + "- NCRED."
 End If
 If fila = 3 And txttp.Selected(fila) Then
   TD_D = "D"
   wsTexto = wsTexto + "- NDEB."
 End If
Next fila

pub_cadena = "SELECT CLI_CUENTA_CONTAB FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_CODCLIE = ? AND CLI_CP = 'C' "
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
PS_REP03(0) = 0
PS_REP03(1) = 0
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


pub_cadena = "SELECT FAR_CODCLIE FROM FACART WHERE  FAR_CODCIA = ? AND FAR_FBG = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ?  AND  FAR_TIPMOV = ? AND FAR_ESTADO <> 'E' "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = 0
PS_REP01(3) = 0
PS_REP01(4) = 0
PS_REP01.MaxRows = 1
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

Fres = 0

NCREDITO:
S_VALOR_VENTA = 0
s_descto = 0
s_valor_igv = 0
S_VALOR_PRECIO = 0

SS_VALOR_VENTA = 0
SS_VALOR_IGV = 0
SS_VALOR_PRECIO = 0

WCONTROL = WCONTROL + 1
pub_cadena = "SELECT DISTINCT FAR_TIPMOV, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_FECHA_COMPRA, FAR_CODCIA, FAR_MONEDA, FAR_BRUTO, FAR_IMPTO, FAR_TOT_DESCTO, FAR_EX_IGV  FROM FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? ) AND (FAR_TIPMOV = 10 OR FAR_TIPMOV = 97 OR FAR_TIPMOV = 98) AND FAR_ESTADO<>'E' AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ?)  ORDER BY FAR_FECHA_COMPRA "
pub_cadena = "SELECT DISTINCT FAR_CODCIA, FAR_TIPMOV, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_MONEDA, FAR_BRUTO, FAR_IMPTO, FAR_TOT_DESCTO, FAR_EX_IGV, FAR_FECHA_COMPRA, FAR_CODCIA FROM FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? ) AND (FAR_TIPMOV = 10 OR FAR_TIPMOV = 97 OR FAR_TIPMOV = 98) AND FAR_ESTADO<>'E' AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ?)  ORDER BY FAR_FECHA_COMPRA "
If fbtxt.Text <> "" And txtSerie.Text = "" Then
   pub_cadena = "SELECT DISTINCT FAR_CODCIA, FAR_TIPMOV, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_FECHA_COMPRA, FAR_CODCIA, FAR_MONEDA, FAR_BRUTO, FAR_IMPTO, FAR_TOT_DESCTO, FAR_EX_IGV  FROM FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND ( FAR_TIPMOV = 10 OR FAR_TIPMOV = 97 OR FAR_TIPMOV = 98 ) AND FAR_ESTADO<>'E' AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?  AND (FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ?) ORDER BY FAR_FECHA_COMPRA "
   PS_REP02(7) = ""
ElseIf fbtxt.Text <> "" And txtSerie.Text <> "" Then
   pub_cadena = "SELECT DISTINCT FAR_CODCIA, FAR_TIPMOV, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_FECHA_COMPRA, FAR_CODCIA, FAR_MONEDA, FAR_BRUTO, FAR_IMPTO, FAR_TOT_DESCTO, FAR_EX_IGV  FROM FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND ( FAR_TIPMOV = 10 OR FAR_TIPMOV = 97 OR FAR_TIPMOV = 98 ) AND FAR_ESTADO<>'E' AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?  AND (FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ?) AND FAR_NUMSER = ? ORDER BY FAR_FECHA_COMPRA "
   PS_REP02(7) = ""
   PS_REP02(8) = ""
ElseIf fbtxt.Text = "" And txtSerie.Text <> "" Then
   pub_cadena = "SELECT DISTINCT FAR_CODCIA, FAR_TIPMOV, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_FECHA_COMPRA, FAR_CODCIA, FAR_MONEDA, FAR_BRUTO, FAR_IMPTO, FAR_TOT_DESCTO, FAR_EX_IGV  FROM FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND ( FAR_TIPMOV = 10 OR FAR_TIPMOV = 97 OR FAR_TIPMOV = 98  ) AND FAR_ESTADO<>'E' AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?  AND FAR_NUMSER = ? AND (FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ?) ORDER BY FAR_FECHA_COMPRA "
End If
WS_SIGNO = 1
  
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = ""
PS_REP02(1) = ""
PS_REP02(2) = ""
PS_REP02(3) = ""
PS_REP02(4) = ""
PS_REP02(5) = 0
PS_REP02(6) = 0
PS_REP02(7) = 0
PS_REP02(8) = 0
PS_REP02(9) = 0
PS_REP02(10) = 0

Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

If checia.Visible And checia.Value = 1 Then
 If Trim(par_llave!par_art_cias) <> "" Then
      nn = 1
    For m_ind = 1 To 15
        ws_codcia = Mid(par_llave!par_art_cias, nn, 2)
        If Trim(ws_codcia) = "" Then Exit For
        PS_REP02(m_ind - 1) = ws_codcia
        nn = nn + 2
    Next m_ind
 End If
Else
 PS_REP02(0) = LK_CODCIA
End If

' el PS_REP1(0) ESTA MAS ABAJO
PS_REP02(5) = wsFECHA1
PS_REP02(6) = wsFECHA2
PS_REP02(7) = TD_F
PS_REP02(8) = TD_B
PS_REP02(9) = TD_N
PS_REP02(10) = TD_D

If fbtxt.Text <> "" And txtSerie.Text = "" Then
PS_REP02(7) = TD_F
PS_REP02(8) = TD_B
PS_REP02(9) = TD_N
PS_REP02(10) = TD_D
ElseIf fbtxt.Text <> "" And txtSerie.Text <> "" Then
PS_REP02(7) = TD_F
PS_REP02(8) = TD_B
PS_REP02(9) = TD_N
PS_REP02(10) = TD_D
PS_REP02(11) = txtSerie.Text

ElseIf fbtxt.Text = "" And txtSerie.Text <> "" Then
PS_REP02(7) = txtSerie.Text
PS_REP02(8) = TD_F
PS_REP02(9) = TD_B
PS_REP02(10) = TD_N
PS_REP02(11) = TD_D

End If



DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep02.Requery
If llave_rep02.EOF Then
  GoTo OTRO_DOCUMENTO
End If
FrmImp2.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.max = llave_rep02.RowCount

WFLAG = ""
SQ_OPER = 1
pu_cp = "C"
pu_codcia = llave_rep02!FAR_CODCIA
wnumfac = llave_rep02!far_numfac
wfecha = llave_rep02!FAR_fecha_compra 'llave_rep02!far_fecha
wserie = llave_rep02!far_numser
wq_fecha = "01/01/1900"
xcuenta = 0
WFLAG = "A"
wflag_numfac = "A"
AWQ_DESCTOS = 0
AWQ_GASTOS = 0
AWQ_FLETES = 0
w_exo = 0
If llave_rep02.EOF Then GoTo CANCELA

Do Until llave_rep02.EOF

  If llave_rep02!FAR_TIPMOV = 10 Then
     If llave_rep02!far_fbg = "F" Or llave_rep02!far_fbg = "B" Then
     Else
        GoTo SALTARIN
     End If
  End If

  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  
  If wfecha <> llave_rep02!FAR_fecha_compra Then '
       GoSub IMPRI_FAC
  End If
  
  wfecha = llave_rep02!FAR_fecha_compra
  ws_tc = 1
  If llave_rep02!FAR_MONEDA = "D" Then
    ws_tc = JALAR(llave_rep02!FAR_fecha_compra)
    If ws_tc <= 0 Then
        MsgBox "Falta Ingresar el Tipo de Cambio del día : " & Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy"), 48, Pub_Titulo
        GoTo CANCELA
    End If
  End If
  If llave_rep02!FAR_TIPMOV = 97 Then
     WS_SIGNO = -1
  Else
     WS_SIGNO = 1
  End If
  
  wq_bruto = Format((val(llave_rep02!FAR_BRUTO) - val(llave_rep02!FAR_TOT_DESCTO)) * ws_tc * WS_SIGNO, "0.000")
  WQ_IMPTO = Format(val(llave_rep02!far_IMPTO) * WS_SIGNO * ws_tc, "0.000")
  WQ_TOTAL = Format((val(llave_rep02!FAR_BRUTO) + val(llave_rep02!far_IMPTO)) * WS_SIGNO * ws_tc, "0.000")
  ts_suma = val(wq_bruto) + val(WQ_IMPTO)
  If ts_suma <> val(WQ_TOTAL) Then
      wq_bruto = val(wq_bruto) + (val(WQ_TOTAL) - ts_suma) ' (Val(wq_bruto) + Val(WQ_IMPTO)) - Val(WQ_TOTAL)
  End If
  'WQ_TOTAL = Val(wq_bruto) + Val(WQ_IMPTO)  ' (Val(llave_rep02!far_bruto) + Val(llave_rep02!far_impto) - Val(llave_rep02!FAR_TOT_DESCTO) + Val(llave_rep02!FAR_GASTOS)) * WS_SIGNO * WS_TC
  PS_REP01(0) = llave_rep02!FAR_CODCIA
  PS_REP01(1) = llave_rep02!far_fbg
  PS_REP01(2) = llave_rep02!far_numser
  PS_REP01(3) = llave_rep02!far_numfac
  PS_REP01(4) = llave_rep02!FAR_TIPMOV
  llave_rep01.Requery
  If llave_rep01.EOF Then
  MsgBox "Verificar el Documento.", 48, Pub_Titulo
  Else
    PS_REP03(0) = llave_rep02!FAR_CODCIA
    PS_REP03(1) = llave_rep01!far_codclie
    llave_rep03.Requery
    If llave_rep03.EOF Then
      ts_codcta = ""
    Else
      ts_codcta = llave_rep03!CLI_CUENTA_CONTAB
    End If
    Fres = Fres + 1
    xl.Sheets(2).Activate
    xl.Cells(Fres, 1) = wq_bruto
    xl.Cells(Fres, 2) = WQ_IMPTO
    xl.Cells(Fres, 3) = WQ_TOTAL
    xl.Cells(Fres, 4) = ts_codcta
    xl.Cells(Fres, 5) = JALA_CTA(ts_codcta)
    xl.Sheets(1).Activate
  End If
  

  
  S_VALOR_PRECIO = S_VALOR_PRECIO + wq_bruto
  s_valor_igv = s_valor_igv + WQ_IMPTO
  S_VALOR_VENTA = S_VALOR_VENTA + WQ_TOTAL
  
  SS_VALOR_PRECIO = SS_VALOR_PRECIO + wq_bruto
  SS_VALOR_IGV = SS_VALOR_IGV + WQ_IMPTO
  SS_VALOR_VENTA = SS_VALOR_VENTA + WQ_TOTAL
SALTARIN:
 llave_rep02.MoveNext
Loop

  GoSub IMPRI_FAC
  
    f1 = f1 + 1
    FILAS = FILAS + 1
    GoSub TOTAL_DIA
    
  xcuenta = c1 + 1
  wranF = "A6:" & "D6"
  xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3
  If WCONTROL = 1 Then
   If cheasiento.Value = 1 Then
   End If
  End If
OTRO_DOCUMENTO:
If WCONTROL >= 1 Then
Else
  GoTo NCREDITO
End If

MOSTRAR:

  f1 = f1 + 2
  xl.Cells(f1, 1) = "Total General = "
  xl.Worksheets(1).Rows(f1).RowHeight = 20
  xl.Cells(f1, 2) = t_valor_precio
  xl.Cells(f1, 3) = t_valor_igv
  xl.Cells(f1, 4) = t_valor_venta

' Ordenando la información para asintos.
'---------------------------------------

If Fres = 0 Then GoTo SALTA

xl.Sheets(2).Activate
wranF = "A" & 1 & ":S" & Fres
'xl.Application.Visible = True
xl.APPLICATION.Worksheets(2).Range(wranF).Sort Key1:=xl.APPLICATION.Worksheets(2).Range("A1")
xl.APPLICATION.Worksheets(2).Range(wranF).Sort Key1:=xl.APPLICATION.Worksheets(2).Range("D1")
fila = 1
f1 = 4
ts_codcta = Trim(Format(xl.Cells(fila, 4), "##########"))
ts_suma = 0
ts_suma_bruto = 0
ts_suma_igv = 0
For fila = 1 To Fres
  If Trim(xl.Cells(fila, 4)) = "" Then GoTo cont_p
  If Trim(ts_codcta) <> Trim(Format(xl.Cells(fila, 4), "##########")) Then
    xl.Sheets(3).Activate
    f1 = f1 + 1
    xl.Cells(f1, 1) = ts_codcta
    xl.Cells(f1, 2) = JALA_CTA(ts_codcta)
    xl.Cells(f1, 3) = ts_suma  ' debe
    xl.Cells(f1, 4) = 0 ' haber
    xl.Cells(f1, 5) = "D"
    xl.Sheets(2).Activate
    ts_codcta = Trim(Format(xl.Cells(fila, 4), "##########"))
    ts_suma = 0
  End If
  ts_suma = ts_suma + val(Format(xl.Cells(fila, 3), "0.000"))
  ts_suma_bruto = ts_suma_bruto + val(Format(xl.Cells(fila, 1), "0.000"))
  ts_suma_igv = ts_suma_igv + val(Format(xl.Cells(fila, 2), "0.000"))
cont_p:
Next fila
xl.Sheets(3).Activate
xl.Cells(1, 1) = "" ' WEMPRESA '
xl.Cells(2, 1) = "PERIODO: '" & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
f1 = f1 + 1
xl.Cells(f1, 1) = ts_codcta
xl.Cells(f1, 2) = JALA_CTA(ts_codcta)
xl.Cells(f1, 3) = ts_suma
xl.Cells(f1, 4) = 0
xl.Cells(f1, 5) = "D"
ts_suma = 0


f1 = f1 + 1
ts_codcta = "401001"
xl.Cells(f1, 1) = ts_codcta
xl.Cells(f1, 2) = JALA_CTA(ts_codcta)
xl.Cells(f1, 3) = 0 ' debe
xl.Cells(f1, 4) = ts_suma_igv ' haber
xl.Cells(f1, 5) = "H"
f1 = f1 + 1
ts_codcta = "701001"
xl.Cells(f1, 1) = ts_codcta
xl.Cells(f1, 2) = JALA_CTA(ts_codcta)
xl.Cells(f1, 3) = 0 ' debe
xl.Cells(f1, 4) = ts_suma_bruto ' haber
xl.Cells(f1, 5) = "H"

' TOTLES Y ORDEN DE ASIENTO
ts_suma = 0
wran1 = "C" & 5
wran2 = "C" & f1
wranF = "C" & f1 + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
ts_suma = ts_suma + val(xl.Range(wranF))
wran1 = "D" & 5
wran2 = "D" & f1
wranF = "D" & f1 + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
ts_suma = ts_suma - val(xl.Range(wranF))

wranF = "C" & f1 + 3
xl.Range(wranF) = ts_suma
wranF = "B" & f1 + 3
xl.Range(wranF) = "Diferencia:"
If chepasa.Value = 1 Then
 xl.Sheets(3).Activate
 ASIENTO_MOVICONT xl, 2
End If

xl.Sheets(1).Activate

SALTA:

  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'" & wsTexto & " -  DEL " & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub

IMPRI_FAC:
       f1 = f1 + 1
       xl.Cells(f1, 1) = "'" & wfecha
       xl.Cells(f1, 2) = Format(val(S_VALOR_PRECIO) - w_exo, "0.000")
       xl.Cells(f1, 3) = val(s_valor_igv)
       xl.Cells(f1, 4) = val(S_VALOR_VENTA)
     
     t_valor_venta = t_valor_venta + S_VALOR_VENTA
     t_descto = t_descto + s_descto
     t_valor_igv = t_valor_igv + s_valor_igv
     t_valor_precio = t_valor_precio + S_VALOR_PRECIO
     
     S_VALOR_VENTA = 0
     s_descto = 0
     s_valor_igv = 0
     S_VALOR_PRECIO = 0
     w_exo = 0

Return

CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\RESU_VENTA_DIA.xls", 0, True, 4, WPAS, WPAS
Return

TOTAL_DIA:
   'xl.Application.Visible = True
  If WCONTROL = 1 Then
    xl.Cells(f1, 1) = "Total Ventas      = "
  ElseIf WCONTROL = 2 Then
    xl.Cells(f1, 1) = "Total N.Creditos  = "
  ElseIf WCONTROL = 3 Then
    xl.Cells(f1, 1) = "Total N.Debito    = "
  End If
  xl.Worksheets(1).Rows(f1).RowHeight = 20
  xl.Cells(f1, 4) = SS_VALOR_VENTA
  xl.Cells(f1, 3) = SS_VALOR_IGV
  xl.Cells(f1, 2) = SS_VALOR_PRECIO
Return




LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"

Return

FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
 GoTo CANCELA
 Resume Next

End Sub
Public Sub PARTE_DIARIO()
On Error GoTo FINTODO
Dim WPCODCIA As String * 2
Dim SUMA_EFEC As Currency
Dim SUMA_EFEC_D As Currency
Dim SUMA_EFEC_DS As Currency
Dim wusu As String * 1
Dim AWQ_NETO_ACT_FIJO As Currency
Dim AWQ_CTA_ACT_FIJO As String
Dim AWQ_BRUTO_ACT_FIJO As Currency
Dim WCONTROL As Integer
Dim WQ
Dim WQ_EFEC As Currency
Dim wfecha
Dim ws_clave
Dim LETRAS(24) As String * 1
Dim ms_numoper As Integer
Dim ms_fecha As Date
Dim wRuta As String
Dim wmonto As Currency
Dim wcodclie As Currency
Dim valor_venta As Currency
Dim descto As Currency
Dim valor_igv As Currency
Dim valor_precio As Currency
Dim S_VALOR_VENTA As Currency
Dim s_descto As Currency
Dim s_valor_igv As Currency
Dim S_VALOR_PRECIO As Currency

Dim t_valor_venta As Currency
Dim t_descto As Currency
Dim t_valor_igv As Currency
Dim t_valor_precio As Currency

Dim wnumfac As Currency
Dim WFLAG As String * 1
Dim wsFECHA1, wsFECHA2
Dim xcuenta As Integer
Dim wq_fecha, wq_codclie, wq_codven, wq_docu, wq_nombre, wq_bruto, wq_gastos, wq_desto, wq_flete, wq_fbg, wq_serie
Dim wq_tot_descto, WQ_IMPTO, WQ_TOTAL, wq_estado, wq_condi
Dim wq_ruc
Dim wflag_numfac
Dim wserie As String * 3
Dim AWQ_BRUTO As Currency
Dim AWQ_DESCTOS As Currency
Dim AWQ_GASTOS As Currency
Dim AWQ_FLETES As Currency
Dim AWQ_IMPTO As Currency
Dim AWQ_NETO As Currency
Dim AWQ_NETO_CRED  As Currency
Dim AWQ_NETO_CONT   As Currency
Dim AWQ_COSTO_VENTA As Currency
Dim WS_SIGNO As Integer
Dim AWQ_TIPO_CAMBIO As Currency
Dim s_suma_dolar As Currency
Dim a_dolar_soles As Currency
Dim s_suma_soles As Currency
Dim d_moneda As String * 1
Dim WWCODUSU As String
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
pub_mensaje = "Imprimir según su Usuario...?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbYes Then
 wusu = "A"
Else
 wusu = " "
End If

GoSub WEXCEL
pub_cadena = ""
xcuenta = 0
WPCODCIA = ""

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
WCONTROL = 0
f1 = 5  'Fila Inicial
t_valor_venta = 0
t_descto = 0
t_valor_igv = 0
t_valor_precio = 0

'NCREDITO: ' empieza

'WCONTROL = WCONTROL + 1
wmonto = 0
wcodclie = 0
valor_venta = 0
descto = 0
valor_igv = 0
valor_precio = 0
wnumfac = 0
WFLAG = ""
xcuenta = 0
wflag_numfac = ""
wserie = ""
AWQ_BRUTO = 0
AWQ_DESCTOS = 0
AWQ_GASTOS = 0
AWQ_FLETES = 0
AWQ_IMPTO = 0
AWQ_NETO = 0
AWQ_NETO_CRED = 0
AWQ_NETO_CONT = 0
AWQ_COSTO_VENTA = 0
AWQ_NETO_ACT_FIJO = 0
AWQ_BRUTO_ACT_FIJO = 0
s_suma_dolar = 0
a_dolar_soles = 0
s_suma_soles = 0
SUMA_EFEC = 0

NCREDITO:
'SUMA_EFEC = 0
S_VALOR_VENTA = 0
s_descto = 0
s_valor_igv = 0
S_VALOR_PRECIO = 0
WQ_EFEC = 0
WCONTROL = WCONTROL + 1

If WCONTROL = 1 Then
  If wusu = "A" Then
    pub_cadena = "SELECT * FROM FACART WHERE (FAR_CODCIA = ? or FAR_CODCIA = ?) " & _
    "AND FAR_TIPMOV = 10 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND " & _
    "FAR_CODUSU  = ? and FAR_FBG <> ' ' AND FAR_FBG <> 'G' and FAR_FBG <> 'X'  AND  FAR_CP = 'C' AND FAR_ESTADO <> 'M'  ORDER BY FAR_FECHA_COMPRA, FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC"
  Else
    pub_cadena = "SELECT * FROM FACART WHERE (FAR_CODCIA = ? or FAR_CODCIA = ?)  AND FAR_TIPMOV = 10 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND  FAR_FBG <> ' ' AND FAR_FBG <> 'G' and FAR_FBG <> 'X'  AND  FAR_CP = 'C' AND FAR_ESTADO <> 'M'  ORDER BY FAR_FECHA_COMPRA, FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC"
  End If
  WS_SIGNO = 1
ElseIf WCONTROL = 2 Then
  If wusu = "A" Then
    pub_cadena = "SELECT * FROM FACART WHERE (FAR_CODCIA = ? OR FAR_CODCIA = ? ) AND FAR_TIPMOV = 97 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND  FAR_CODUSU  = ? AND FAR_FBG <> ' ' AND FAR_FBG <> 'G' and FAR_FBG <> 'X' AND FAR_CP = 'C'  ORDER BY FAR_CODCIA, FAR_TIPMOV, FAR_FBG DESC , FAR_FECHA_COMPRA,  FAR_NUMSER,FAR_NUMFAC"
  Else
    pub_cadena = "SELECT * FROM FACART WHERE (FAR_CODCIA = ? OR FAR_CODCIA = ?) AND FAR_TIPMOV = 97 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND FAR_FBG <> ' ' AND FAR_FBG <> 'G' and FAR_FBG <> 'X' AND FAR_CP = 'C'  ORDER BY FAR_CODCIA, FAR_TIPMOV, FAR_FBG DESC , FAR_FECHA_COMPRA,  FAR_NUMSER,FAR_NUMFAC"
  End If
  WS_SIGNO = -1
ElseIf WCONTROL = 3 Then
  If wusu = "A" Then
    pub_cadena = "SELECT * FROM FACART WHERE (FAR_CODCIA = ? or FAR_CODCIA = ? ) AND FAR_TIPMOV = 98 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND FAR_CODUSU  = ? AND FAR_FBG <> ' ' AND FAR_FBG <> 'G' and FAR_FBG <> 'X' AND FAR_CP = 'C' ORDER BY FAR_CODCIA, FAR_TIPMOV, FAR_FBG DESC , FAR_FECHA_COMPRA,  FAR_NUMSER,FAR_NUMFAC"
  Else
    pub_cadena = "SELECT * FROM FACART WHERE (FAR_CODCIA = ? or FAR_CODCIA = ?) AND FAR_TIPMOV = 98 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND FAR_FBG <> ' ' AND FAR_FBG <> 'G' and FAR_FBG <> 'X' AND FAR_CP = 'C' ORDER BY FAR_CODCIA , FAR_TIPMOV, FAR_FBG DESC , FAR_FECHA_COMPRA,  FAR_NUMSER,FAR_NUMFAC"
  End If
  WS_SIGNO = 1
End If
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
'If WCONTROL = 1 Then
 PS_REP02(0) = 0
 PS_REP02(1) = 0
 PS_REP02(2) = LK_FECHA_DIA
 PS_REP02(3) = LK_FECHA_DIA
 If wusu = "A" Then PS_REP02(4) = 0
'End If
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT ALL_NUMOPER FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ? AND ALL_NUMOPER = ? ORDER BY ALL_CODCIA , ALL_NUMOPER"
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
PS_REP03(0) = 0
PS_REP03(1) = LK_FECHA_DIA
PS_REP03(2) = 0
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


' el PS_REP1(0) ESTA MAS ABAJO
If checia.Value = 0 Then
 PS_REP02(0) = LK_CODCIA
 PS_REP02(1) = " "
Else
 PS_REP02(0) = "01"
 PS_REP02(1) = "02"
End If
PS_REP02(2) = wsFECHA1
PS_REP02(3) = wsFECHA2
If wusu = "A" Then
  PS_REP02(4) = LK_CODUSU
End If

DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep02.Requery
If llave_rep02.EOF Then
  GoTo OTRO_DOCUMENTO
End If
FrmImp2.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.max = llave_rep02.RowCount

WFLAG = ""
SQ_OPER = 1
pu_cp = "C"
pu_codcia = LK_CODCIA
wnumfac = llave_rep02!far_numfac
wfecha = "dd" 'llave_rep02!far_fbg 'llave_rep02!far_fecha
wserie = llave_rep02!far_numser
wq_fecha = "01/01/1900"
xcuenta = 0
WFLAG = "A"
wflag_numfac = "A"
AWQ_DESCTOS = 0
AWQ_GASTOS = 0
AWQ_FLETES = 0

Do Until llave_rep02.EOF
'  If llave_rep02!FAR_numfac = 51955 Then Stop
  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  If wfecha = llave_rep02!far_fbg Then '
     If wflag_numfac = "A" Then
       GoSub IMPRI_FAC
     End If
     wflag_numfac = ""
     f1 = f1 + 1
     GoSub TOTAL_DIA
     t_valor_venta = t_valor_venta + S_VALOR_VENTA
     t_descto = t_descto + s_descto
     t_valor_igv = t_valor_igv + s_valor_igv
     t_valor_precio = t_valor_precio + S_VALOR_PRECIO
     wnumfac = llave_rep02!far_numfac
     wfecha = llave_rep02!far_fbg
     S_VALOR_VENTA = 0
     s_descto = 0
     s_valor_igv = 0
     S_VALOR_PRECIO = 0
  End If
  
  If wnumfac <> llave_rep02!far_numfac Then
     GoSub IMPRI_FAC
     wflag_numfac = ""
     wnumfac = llave_rep02!far_numfac
  ElseIf val(wserie) <> val(llave_rep02!far_numser) And CDate(wq_fecha) = CDate(llave_rep02!FAR_fecha_compra) Then
    If wflag_numfac <> "" Then
      GoSub IMPRI_FAC
      wflag_numfac = ""
      wnumfac = llave_rep02!far_numfac
    End If
  ElseIf val(wserie) = val(llave_rep02!far_numser) And CDate(wq_fecha) = CDate(llave_rep02!FAR_fecha_compra) Then
    If Trim(wq_fbg) <> Trim(llave_rep02!far_fbg) And wflag_numfac <> "" Then
      GoSub IMPRI_FAC
      wflag_numfac = ""
      wnumfac = llave_rep02!far_numfac
    End If
  End If
  wq_fecha = Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy")
  wq_codclie = llave_rep02!far_codclie
  wq_codven = llave_rep02!FAR_CODVEN
  wq_fbg = Trim(llave_rep02!far_fbg)
  wq_serie = "'" & llave_rep02!far_numser
  wserie = llave_rep02!far_numser
  wq_docu = "'" & llave_rep02!far_numfac
  d_moneda = llave_rep02!FAR_MONEDA
  wq_nombre = ""
  wq_bruto = ((val(llave_rep02!FAR_BRUTO) - val(llave_rep02!FAR_TOT_DESCTO)) * WS_SIGNO)
  wq_desto = (val(llave_rep02!FAR_TOT_DESCTO) * WS_SIGNO)
  wq_gastos = (val(llave_rep02!FAR_GASTOS) * WS_SIGNO)
  wq_flete = (val(Nulo_Valor0(llave_rep02!FAR_TOT_FLETE)) * WS_SIGNO)
  WQ_IMPTO = Format((val(llave_rep02!far_IMPTO) * WS_SIGNO), "0.000")
  WQ_TOTAL = redondea((val(llave_rep02!FAR_BRUTO) + val(llave_rep02!far_IMPTO) - val(llave_rep02!FAR_TOT_DESCTO) + val(llave_rep02!FAR_GASTOS)) * WS_SIGNO)
  ms_numoper = llave_rep02!FAR_NUMOPER
  ms_fecha = llave_rep02!FAR_fecha
  WWCODUSU = llave_rep02!far_codusu
  
  If wq_bruto = 0 Then WQ_TOTAL = 0
  wq_estado = llave_rep02!far_estado
  If wq_estado <> "E" Then
    If Left(UCase(llave_rep02!far_subtra), 1) <> "A" Then
     AWQ_COSTO_VENTA = AWQ_COSTO_VENTA + ((llave_rep02!FAR_COSPRO * llave_rep02!far_cantidad)) * WS_SIGNO
    End If
  End If
  WQ_EFEC = 0
  WPCODCIA = llave_rep02!FAR_CODCIA
  If llave_rep02!far_signo_car <> 0 Then
     wq_condi = "CRED"
  Else
     wq_condi = "CONT"
     WQ_EFEC = WQ_TOTAL
  End If
  wflag_numfac = "A"
  WFLAG = "A"
 llave_rep02.MoveNext
Loop
If wflag_numfac = "A" Then
    GoSub IMPRI_FAC
    wflag_numfac = ""
End If
If WFLAG = "A" Then
    f1 = f1 + 1
    GoSub TOTAL_DIA
    t_valor_venta = t_valor_venta + S_VALOR_VENTA
    t_descto = t_descto + s_descto
    t_valor_igv = t_valor_igv + s_valor_igv
    t_valor_precio = t_valor_precio + S_VALOR_PRECIO
End If
  xcuenta = c1 + 1
  wranF = "A6:" & "O6"
  xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3
  If WCONTROL = 1 Then
   If cheasiento.Value = 1 Then
   End If
  End If
OTRO_DOCUMENTO:
If WCONTROL >= 3 Then
Else
  GoTo NCREDITO
End If
  wranF = "A6:" & "O6"
  xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3
  f1 = f1 + 1
  xl.Cells(f1, 7) = "TOTAL S/. = "
  xl.Cells(f1, 12) = s_suma_soles
  xl.Cells(f1, 15) = SUMA_EFEC
  
  f1 = f1 + 1
  xl.Cells(f1, 7) = "TOTAL US$. = "
  xl.Cells(f1, 12) = s_suma_dolar
  xl.Cells(f1, 15) = SUMA_EFEC_D
  f1 = f1 + 1
  xl.Cells(f1, 7) = "AL CAMBIO EN SOLES ="
  xl.Cells(f1, 12) = a_dolar_soles
  xl.Cells(f1, 15) = SUMA_EFEC_DS
    
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'" & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub

IMPRI_FAC:
    PS_REP03(0) = WPCODCIA  'LK_CODCIA
    PS_REP03(1) = ms_fecha
    PS_REP03(2) = ms_numoper
    llave_rep03.Requery
    If llave_rep03.EOF Then
      MsgBox "VER : Relacion ... " & ms_fecha & " " & ms_numoper & "  = " & WQ_TOTAL
    End If
     f1 = f1 + 1
     pu_codclie = wq_codclie
     pu_codcia = WPCODCIA
     LEER_CLI_LLAVE
     If Not cli_llave.EOF Then
         wq_codclie = cli_llave!cli_codclie 'xl.Cells(F1, 2)
         wq_nombre = Trim(cli_llave!CLI_NOMBRE)
         wq_ruc = Trim(cli_llave!cli_ruc_esposo)
     End If
  AWQ_TIPO_CAMBIO = 1
  If d_moneda = "D" Then
     PUB_CAL_INI = wq_fecha ' llave_rep02!FAR_FECHA_COMPRA
     PUB_CAL_FIN = wq_fecha ' llave_rep02!FAR_FECHA_COMPRA
     PUB_CODCIA = LK_CODCIA
     LEER_CAL_LLAVE
     If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 Then
       AWQ_TIPO_CAMBIO = 0
     Else
       AWQ_TIPO_CAMBIO = cal_llave!cal_tipo_cambio
     End If
     If AWQ_TIPO_CAMBIO <= 0 Then
         MsgBox "Definir Tipo de Cambios para el Periodo Actual. Dia : " & llave_rep02!FAR_fecha & " (en el Calendario del Sistema)", 48, Pub_Titulo
         xl.DisplayAlerts = False
         xl.Cells(f1 + 1, 1) = "Falta Tipo de Cambio.... "
         GoTo CANCELA
         Exit Sub
     End If
  End If
  
     xl.Cells(f1, 1) = "'" & wq_fecha
     xl.Cells(f1, 3) = wq_ruc
     xl.Cells(f1, 2) = wq_nombre
     xl.Cells(f1, 4) = WPCODCIA & "-" & wq_condi
     xl.Cells(f1, 5) = wq_fbg
     xl.Cells(f1, 6) = wq_serie
     xl.Cells(f1, 7) = wq_docu
     If wq_estado = "E" Then
         xl.Cells(f1, 2) = "[ANULADO] " & wq_nombre
     Else
         xl.Cells(f1, 2) = wq_nombre
     End If
     If wq_estado <> "E" Then
      If True Then 'Left(cli_llave!CLI_CUENTA_CONTAB, 2) <> "12" Then
        If d_moneda = "S" Then
          xl.Cells(f1, 8) = "S/."
        Else
          xl.Cells(f1, 8) = "US$."
        End If
       If d_moneda = "D" Then
            s_suma_dolar = s_suma_dolar + val(WQ_TOTAL)
            a_dolar_soles = a_dolar_soles + redondea(val(WQ_TOTAL) * AWQ_TIPO_CAMBIO)
            SUMA_EFEC_D = SUMA_EFEC_D + WQ_EFEC
            SUMA_EFEC_DS = SUMA_EFEC_DS + redondea(val(WQ_EFEC) * AWQ_TIPO_CAMBIO)
       Else
            s_suma_soles = s_suma_soles + val(WQ_TOTAL)
            a_dolar_soles = a_dolar_soles + val(WQ_TOTAL)
            SUMA_EFEC = SUMA_EFEC + WQ_EFEC
            SUMA_EFEC_DS = SUMA_EFEC_DS + val(WQ_EFEC)
       End If
       xl.Cells(f1, 9) = wq_bruto
       xl.Cells(f1, 10) = val(wq_desto)
       xl.Cells(f1, 11) = val(WQ_IMPTO)
       xl.Cells(f1, 12) = val(WQ_TOTAL)
       xl.Cells(f1, 15) = val(WQ_EFEC)
       xl.Cells(f1, 16) = WWCODUSU
       
       AWQ_BRUTO_ACT_FIJO = AWQ_BRUTO_ACT_FIJO + wq_bruto
       AWQ_NETO_ACT_FIJO = AWQ_NETO_ACT_FIJO + val(WQ_TOTAL)
       AWQ_CTA_ACT_FIJO = Trim(cli_llave!CLI_CUENTA_CONTAB)
       s_valor_igv = s_valor_igv + val(WQ_IMPTO)
       AWQ_IMPTO = AWQ_IMPTO + val(WQ_IMPTO)
       ' ACUMULA OTROS ********
       S_VALOR_VENTA = S_VALOR_VENTA + wq_bruto
       s_descto = s_descto + val(wq_desto)
       S_VALOR_PRECIO = S_VALOR_PRECIO + val(WQ_TOTAL)
      Else
       S_VALOR_VENTA = S_VALOR_VENTA + wq_bruto
       s_descto = s_descto + val(wq_desto)
       s_valor_igv = s_valor_igv + val(WQ_IMPTO)
       S_VALOR_PRECIO = S_VALOR_PRECIO + val(WQ_TOTAL)
       If d_moneda = "S" Then
          xl.Cells(f1, 8) = "S/."
       Else
          xl.Cells(f1, 8) = "US$."
       End If
          xl.APPLICATION.Visible = True
       xl.Cells(f1, 9) = wq_bruto
       xl.Cells(f1, 10) = val(wq_desto)
       xl.Cells(f1, 11) = val(WQ_IMPTO)
       xl.Cells(f1, 12) = val(WQ_TOTAL)
       xl.Cells(f1, 15) = val(WQ_EFEC)
       'SUMA_EFEC = SUMA_EFEC + WQ_EFEC

       AWQ_IMPTO = AWQ_IMPTO + val(WQ_IMPTO)
       AWQ_NETO = AWQ_NETO + val(WQ_TOTAL)
       AWQ_BRUTO = AWQ_BRUTO + wq_bruto
       AWQ_DESCTOS = wq_desto
       AWQ_GASTOS = wq_gastos
       AWQ_FLETES = wq_flete
       If wq_condi = "CRED" Then
         xl.Cells(f1, 15) = -1
         AWQ_NETO_CRED = AWQ_NETO_CRED + val(WQ_TOTAL)
       Else
         xl.Cells(f1, 15) = 0
         AWQ_NETO_CONT = AWQ_NETO_CONT + val(WQ_TOTAL)
       End If
     End If
     End If
Return

CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo PARTE.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\PARTE.xls", 0, True, 4, WPAS, WPAS
Return

TOTAL_DIA:
Return
  If wfecha = "F" Then
    xl.Cells(f1, 1) = "Total Facturas    = "
  ElseIf wfecha = "B" Then
    xl.Cells(f1, 1) = "Total Boletas     = "
  ElseIf wfecha = "N" Then
    xl.Cells(f1, 1) = "Total N.Creditos  = "
  ElseIf wfecha = "D" Then
    xl.Cells(f1, 1) = "Total N.Debito    = "
  End If
  xl.Cells(f1, 2) = ""
  xl.Cells(f1, 3) = ""
  xl.Cells(f1, 7) = ""
  xl.Cells(f1, 9) = S_VALOR_VENTA
  xl.Cells(f1, 10) = s_descto
  xl.Cells(f1, 11) = s_valor_igv
  xl.Cells(f1, 12) = S_VALOR_PRECIO
Return


LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"

Return

FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
 GoTo CANCELA
 Resume Next
End Sub

Public Sub FECHA_STOCK()
'On Error GoTo FINTODO
Dim wSTOCK As Currency
Dim WSCODART As Currency
Dim ws_clave As String
Dim LETRAS(24) As String * 1
Dim WSFECHA As Date
Dim acu_val_ingresos As Currency
Dim acu_val_salidas As Currency
Dim acu_cant_dia As Currency
Dim acu_saldo As Currency
Dim acu_stock As Currency
Dim wsfile As String
Dim walterno As String
Dim wdnombre As String
Dim WD_COSPRO As Currency
Dim CADENITA  As String
Dim wwfami As Integer
'If Val(Right(listVen.Text, 6)) = 0 Then
'  GoTo CANCELA
'End If
wsfile = ""
Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
CADENITA = ""
For fila = 0 To vendmulti.ListCount - 1
 vendmulti.ListIndex = fila
 If vendmulti.Selected(fila) Then
   CADENITA = CADENITA + "ART_FAMILIA = " & Trim(Right(vendmulti.Text, 8)) & " OR "
 End If
Next fila
If CADENITA <> "" Then
  CADENITA = "(" & Mid(CADENITA, 1, Len(CADENITA) - 4) & ")"
End If

If CADENITA <> "" Then
  pub_cadena = "SELECT ART_FAMILIA, ART_KEY, ART_ALTERNO, ART_NOMBRE, ARM_STOCK, ARM_COSPRO, ARM_SALDO_S FROM ARTI, ARTICULO  WHERE (ARM_CODART = ART_KEY) AND (ARM_CODCIA = ART_CODCIA) AND ART_CODCIA = ? AND ART_KEY <> 0 AND " & CADENITA & " ORDER BY ART_FAMILIA, ART_ALTERNO "
Else
  pub_cadena = "SELECT ART_FAMILIA, ART_KEY, ART_ALTERNO, ART_NOMBRE, ARM_STOCK, ARM_COSPRO, ARM_SALDO_S FROM ARTI, ARTICULO  WHERE (ARM_CODART = ART_KEY) AND (ARM_CODCIA = ART_CODCIA) AND ART_CODCIA = ? AND ART_KEY <> 0  ORDER BY ART_FAMILIA, ART_ALTERNO "
End If
'PRUEBA X ARTI
'pub_cadena = "SELECT ART_FAMILIA, ART_KEY, ART_ALTERNO, ART_NOMBRE, ARM_STOCK, ARM_COSPRO, ARM_SALDO_S FROM ARTI, ARTICULO  WHERE (ARM_CODART = ART_KEY) AND (ARM_CODCIA = ART_CODCIA) AND ART_CODCIA = ? AND ART_KEY= 16798 ORDER BY ART_FAMILIA, ART_ALTERNO "

Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
'Debug.Print pub_cadena
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

'pub_cadena = "SELECT FAR_FECHA_COMPRA, FAR_CANTIDAD, FAR_SIGNO_ARM, FAR_COSPRO, FAR_CODART FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA_COMPRA >= ?  AND FAR_CODART = ? AND FAR_ESTADO <> 'E' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC , FAR_NUMOPER2"

pub_cadena = "SELECT FAR_STOCK, FAR_COSPRO FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA_COMPRA <= ? AND FAR_CODART = ? and far_estado<>'E' ORDER BY FAR_FECHA_COMPRA DESC, FAR_SIGNO_ARM, FAR_NUMOPER DESC"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = LK_FECHA_DIA
PS_REP01(2) = 0
PS_REP01.MaxRows = 1
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
DoEvents
Dim wsFECHA1, wsFECHA2
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If

ws_clave = PUB_CLAVE
GoSub WEXCEL
FrmImp2.ProgBar.Visible = True
DoEvents
'xl.Worksheets(1).Activate
'GoSub LETRAS
xcuenta = 0
xl.Cells(1, 1) = "DISTRIBUIDORA ROMERO S.R.L."
xl.Cells(2, 1) = "'L I S T A D O   D E   S T O C K   A L " & Format(wsFECHA1, "dd/mm/yyyy")
'xl.Cells(3, 1) = "LINEA : " & Left(listVen.Text, 25)
'xl.Cells(4, 1) = "Vendedor : " & Trim(listVen.Text)


f1 = 6 'Fila Inicial
PS_REP02(0) = LK_CODCIA
llave_rep02.Requery
If llave_rep02.RowCount <> 0 Then
 FrmImp2.ProgBar.Min = 0
 FrmImp2.ProgBar.Value = 0
 FrmImp2.ProgBar.max = llave_rep02.RowCount
End If

FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Procesando . . . "
DoEvents

acu_val_ingresos = 0
acu_val_salidas = 0
acu_cant_dia = 0
wSTOCK = 0
WD_COSPRO = 0
acu_saldo = 0
wwfami = -1
Do Until llave_rep02.EOF
        FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
        If wwfami <> llave_rep02!art_familia Then
          f1 = f1 + 1
          xl.Cells(f1, 1) = ""
          SQ_OPER = 1
          PUB_CODCIA = LK_CODCIA
          PUB_TIPREG = 122
          PUB_NUMTAB = llave_rep02!art_familia
          LEER_TAB_LLAVE
          If tab_llave.EOF Then
             MsgBox "El Producto : " & llave_rep02!ART_ALTERNO & " " & llave_rep02!art_nombre & "  Definir Linea nuevamente", 48, Pub_Titulo
             xl.Cells(f1, 2) = "Linea: "
          Else
            xl.Cells(f1, 2) = "Linea: " & Trim(tab_llave!tab_NOMLARGO)
          End If
          wwfami = llave_rep02!art_familia
        End If
        PS_REP01(0) = LK_CODCIA
        PS_REP01(1) = wsFECHA1
        PS_REP01(2) = llave_rep02!ART_KEY ' Val(Right(listven.Text, 6))
        'If llave_rep02!ART_KEY = 77854 Then Stop
        walterno = llave_rep02!ART_ALTERNO
        wdnombre = llave_rep02!art_nombre
        llave_rep01.Requery
        acu_cant_dia = 0
        acu_val_ingresos = 0
        acu_val_salidas = 0
        acu_cant_dia = 0
        WD_COSPRO = 0
        wSTOCK = 0
        If llave_rep01.EOF = True Then
             GoTo SIGUER
             If LK_FLAG_SOS = "A" Then
                If val(llave_rep02!ARM_saldo_s) = 0 Then GoTo SIGUER
             Else
                If val(llave_rep02!arm_stock) = 0 Then GoTo SIGUER
             End If
             PUB_CODART = llave_rep02!ART_KEY
             pu_codcia = LK_CODCIA
             SQ_OPER = 1
             If WD_COSPRO = 0 Then WD_COSPRO = val(llave_rep02!arm_cospro)
             SQ_OPER = 1
             PUB_SECUEN = 0
             LEER_PRE_LLAVE
             f1 = f1 + 1
             xl.Cells(f1, 1) = "'" & Trim(walterno)
             xl.Cells(f1, 2) = Trim(wdnombre)
             xl.Cells(f1, 3) = pre_llave!pre_unidad
             If LK_FLAG_SOS = "A" Then
                xl.Cells(f1, 4) = val(llave_rep02!ARM_saldo_s)
             Else
                xl.Cells(f1, 4) = val(llave_rep02!arm_stock) '- Val(acu_val_ingresos) + Val(acu_val_salidas) + Val(acu_cant_dia)
             End If
             xl.Cells(f1, 5) = val(WD_COSPRO)
             xl.Cells(f1, 6) = Format(val(WD_COSPRO) * val(xl.Cells(f1, 4)), "0.00")
             acu_saldo = acu_saldo + val(xl.Cells(f1, 5))
             GoTo SIGUER
'             MsgBox "!!! NO EXISTEN Datos ...", 48, Pub_Titulo
'             GoTo CANCELA
        End If '
        'fila = 0
    WSCODART = llave_rep02!ART_KEY
    WD_COSPRO = val(llave_rep01!FAR_COSPRO)
    wSTOCK = val(llave_rep01!FAR_STOCK)
    PUB_CODART = WSCODART
    pu_codcia = LK_CODCIA
    SQ_OPER = 1
    SQ_OPER = 1
    PUB_SECUEN = 0
    LEER_PRE_LLAVE
    f1 = f1 + 1
    xl.Cells(f1, 1) = "'" & Trim(walterno)
    xl.Cells(f1, 2) = Trim(wdnombre)
    xl.Cells(f1, 3) = pre_llave!pre_unidad
    xl.Cells(f1, 4) = val(wSTOCK)
    xl.Cells(f1, 5) = val(WD_COSPRO)
    xl.Cells(f1, 6) = Format(val(WD_COSPRO) * val(xl.Cells(f1, 4)), "0.00")
    acu_saldo = acu_saldo + val(xl.Cells(f1, 6))
SIGUER:
llave_rep02.MoveNext
Loop
f1 = f1 + 1
    xl.Cells(f1, 2) = "TOTAL GENERAL = "
    xl.Cells(f1, 6) = Format(acu_saldo, "0.00")
    If f1 <> 6 Then
        wran1 = "D" & 6
        wran2 = "D" & f1 - 1
        wranF = "D" & f1
        xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
     End If
  FrmImp2.lblproceso.Caption = "Procesando . . .  un Momento ."
  'xl.Application.Worksheets("HOJA DE COMISIONES x VENDEDOR").Range(wran1).Sort Key1:=xl.Application.Worksheets("HOJA DE COMISIONES x VENDEDOR").Range("O7")
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
 ' xl.Cells(2, 2) = "INFORME DE COMISIONES x VENDEDOR"
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub



LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"
Return

WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblproceso.Caption = "Abriendo , Archivo Comisiones.xls . . . "
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\STOCKF.xls", 0, True, 4, WPAS, WPAS

Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
  
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.APPLICATION.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
Exit Sub
End Sub







Public Sub IMP_COMI_NETAS()
'On Error GoTo FINTODO
Dim ws_clave As String
Dim LETRAS(24) As String * 1
Dim WSFECHA As Date
Dim wcodven As Integer
Dim wvalor
Dim ws_ingresos As Currency
Dim ws_salidas As Currency
Dim val_ingresos As Currency
Dim val_salidas As Currency
Dim acu_val_ingresos As Currency
Dim acu_val_salidas As Currency
Dim Wche As Integer
Dim wkSELECT As String
Dim wsfile As String
Dim WS_CALCULADA  As Currency
Dim WS_DIAS As Integer
Dim WS_COMI As Currency
wsfile = ""
Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
wcodven = val(Left(listven.Text, 4))
pub_cadena = "SELECT * FROM CARACU WHERE CAA_CODCIA = ? AND CAA_CODVEN = ? AND CAA_FECHA >= ? AND CAA_FECHA <= ? AND CAA_ESTADO <> 'E'  AND CAA_CONCEPTO <> 'Extorno -' AND CAA_SIGNO_CAR = -1 AND CAA_NOTA <> 'N'AND CAA_SIGNO_CAJA <> 55 AND (CAA_TIPDOC ='FA' OR CAA_TIPDOC ='CC')  AND (CAA_FBG ='F' OR CAA_FBG ='B' OR CAA_FBG ='G')  ORDER BY CAA_CODCLIE, CAA_FECHA,CAA_NUM_OPER, CAA_SALDO_CAR"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = LK_FECHA_DIA
PS_REP01(3) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
DoEvents
Dim wsFECHA1, wsFECHA2
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If

PS_REP01(0) = LK_CODCIA
PS_REP01(1) = wcodven
PS_REP01(2) = wsFECHA1
PS_REP01(3) = wsFECHA2
llave_rep01.Requery
If llave_rep01.EOF = True Then
   MsgBox "!!! NO EXISTEN Datos ...", 48, Pub_Titulo
   GoTo CANCELA
End If
ws_clave = PUB_CLAVE
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.max = llave_rep01.RowCount
FrmImp2.ProgBar.Value = 0
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
FrmImp2.ProgBar.Visible = True
DoEvents
xcuenta = 0
xl.Cells(1, 1) = Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia")))
xl.Cells(3, 1) = "'Comisiones del " & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
xl.Cells(4, 1) = "Vendedor : " & Trim(listven.Text)

f1 = 6  'Fila Inicial

FrmImp2.lblproceso.Caption = "Procesando . . .  un Momento ."
DoEvents
fila = 0
acu_val_ingresos = 0
acu_val_salidas = 0
Do Until llave_rep01.EOF
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   f1 = f1 + 1
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   pu_codclie = llave_rep01!CAA_CODCLIE
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
      MsgBox "Error en Codigo de cliente, NO EXISTE ...", 48, Pub_Titulo
      GoTo CANCELA
   End If
   xl.Cells(f1, 1) = wcodven
   If llave_rep01!CAA_FBG = "F" Then
     xl.Cells(f1, 2) = "FAC."
   ElseIf llave_rep01!CAA_FBG = "B" Then
     xl.Cells(f1, 2) = "BOL."
   ElseIf llave_rep01!CAA_FBG = "G" Then
     xl.Cells(f1, 2) = "GUIA"
   End If
   xl.Cells(f1, 3) = "'" & llave_rep01!CAa_numser & " - " & llave_rep01!CAa_numfac
   xl.Cells(f1, 4) = Trim(cli_llave!CLI_NOMBRE)
   SQ_OPER = 1
   pu_cp = "C"
   pu_codclie = cli_llave!cli_codclie
   pu_codcia = LK_CODCIA
   PUB_SERDOC = llave_rep01!caa_serdoc
   PUB_NUMDOC = llave_rep01!CAA_NUMDOC
   PUB_TIPDOC = llave_rep01!CAA_TIPDOC
   LEER_CAR_LLAVE
   If car_llave.EOF Then
    MsgBox "Documento Extornado.... ", 48, Pub_Titulo
    GoTo OTRITO
   End If
   xl.Cells(f1, 5) = "'" & car_llave!CAR_FECHA_INGR
   xl.Cells(f1, 6) = "'" & car_llave!car_fecha_vcto_orig
   xl.Cells(f1, 7) = "'" & llave_rep01!CAA_FECHA
   xl.Cells(f1, 8) = val(llave_rep01!CAA_IMPORTE) * -1 ' Monto Pagado
   xl.Cells(f1, 9) = val(car_llave!CAR_IMP_INI) 'deuda original
   xl.Cells(f1, 10) = val(car_llave!car_importe) 'saldo actual
   xl.Cells(f1, 11) = val(car_llave!CAR_COMISION) 'comision kardex
   'WS_CALCULADA = (llave_rep01!CAA_IMPORTE * -1 * car_llave!CAR_COMISION) / (car_llave!CAR_IMP_INI)
   WS_DIAS = DateDiff("d", car_llave!car_fecha_vcto_orig, llave_rep01!CAA_FECHA)
   If val(car_llave!CAR_NUM_REN) = 24 Or val(car_llave!CAR_NUM_REN) = 27 Then
      WS_COMI = POR_COMI(WS_DIAS, 444)
   Else
      WS_COMI = POR_COMI(WS_DIAS, 445)
   End If
   WS_CALCULADA = (llave_rep01!CAA_IMPORTE * -1) * (WS_COMI / 100)
   If WS_COMI = -99 Then
      GoTo CANCELA
   End If
   xl.Cells(f1, 13) = WS_COMI '% comision
   xl.Cells(f1, 12) = WS_CALCULADA ' * WS_COMI 'comision pagar
   xl.Cells(f1, 14) = (WS_CALCULADA / ((LK_IGV / 100) + 1)) 'comision calculada
   xl.Cells(f1, 15) = WS_DIAS
OTRITO:
   llave_rep01.MoveNext
Loop
  wran1 = "H" & 7
  wran2 = "H" & f1
  wranF = "H" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "I" & 7
  wran2 = "I" & f1
  wranF = "I" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "J" & 7
  wran2 = "J" & f1
  wranF = "J" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "N" & 7
  wran2 = "N" & f1
  wranF = "N" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 
  wran1 = "A" & 7 & ":O" & f1
  xl.APPLICATION.Worksheets("HOJA DE COMISIONES x VENDEDOR").Range(wran1).Sort Key1:=xl.APPLICATION.Worksheets("HOJA DE COMISIONES x VENDEDOR").Range("O7")
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(2, 2) = "INFORME DE COMISIONES x VENDEDOR"
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub



WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblproceso.Caption = "Abriendo , Archivo Comisiones.xls . . . "
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\OFFICE\Comisiones.xls", 0, True, 4, WPAS, WPAS

Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
  
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.APPLICATION.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
 
End Sub

Public Sub REPO_CAJA_DET2()
Dim SOLES_TOTAL_CAJA_I As Currency
Dim DOLAR_TOTAL_CAJA_I As Currency
Dim TOTAL_CAJA_SD_I As Currency

Dim SALDO_ANTERIOR As Currency
FrmImp2.Pantalla.Enabled = False
Dim wusu As String * 1
Dim ww_fbg
Dim ww_numser
Dim ww_numfac
Dim ww_numini
Dim ww_serini
Dim ww_numfin
Dim WMONTO_SOLES As Currency
Dim WMONTO_DOLAR As Currency
Dim wsFECHA1
Dim WDOLAR_CRED As Currency
Dim WSOLES_CRED As Currency
Dim TT_TOTAL_SOLES As Currency
Dim TT_TOTAL_DOLAR  As Currency
Dim WSOLES_DOLAR As Currency
Dim TT_DOLAR_SOLES As Currency
Dim WSOLES_DOLAR_CRED As Currency
Dim WSOLES_DOLAR_TOT As Currency
Dim SOLES_TOTAL_COBRA As Currency
Dim DOLAR_TOTAL_COBRA As Currency
Dim TOTAL_COBRA_SD As Currency
Dim WS_TIPO_C As Currency

Dim SOLES_TOTAL_ADEL    As Currency
Dim TOTAL_ADEL_SD     As Currency
Dim DOLAR_TOTAL_ADEL    As Currency

Dim SOLES_TOTAL_CAJA   As Currency
Dim TOTAL_CAJA_SD     As Currency
Dim DOLAR_TOTAL_CAJA    As Currency

Dim SOLES_TOTAL_BANCO  As Currency
Dim DOLAR_TOTAL_BANCO  As Currency
Dim TOTAL_BANCO_SD  As Currency
Dim SALDO_ANTERIOR_D As Currency

'Dim TOTAL_ADEL_SD  As Currency


If Right(txtFecha.Text, 2) = "__" Then
     wsFECHA1 = Left(txtFecha.Text, 8)
Else
     wsFECHA1 = Trim(txtFecha.Text)
End If
If Not IsDate(wsFECHA1) Then
 FrmImp2.Pantalla.Enabled = True
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
WS_TIPO_C = 0
WS_TIPO_C = JALAR(CDate(wsFECHA1))
If WS_TIPO_C = 0 Then
  MsgBox "Ingresar Tipo de Cambio del Dia ", 48, Pub_Titulo
  Exit Sub
End If


FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.max = 5
FrmImp2.ProgBar.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
f1 = 5
'pub_mensaje = "Imprimir según su Usuario...?"
'Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
'If Pub_Respuesta = vbYes Then
'  wusu = "A"
'Else
  wusu = " " '
pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10  AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 2107 AND ALL_CODTRA <> 1111 AND  ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ? AND ALL_CODTRA = 9999  "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = LK_FECHA_DIA
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = wsFECHA1
llave_rep02.Requery
SALDO_ANTERIOR_D = 0
SALDO_ANTERIOR = 0
If llave_rep02.EOF Then
 MsgBox "NO INICIALIZO EL DIA ", 48, Pub_Titulo
Else
 SALDO_ANTERIOR = val(llave_rep02!ALL_IMPORTE)
 SALDO_ANTERIOR_D = val(llave_rep02!ALL_IMPORTE_DOLL)
End If
'pub_cadena = "SELECT PAR_NOMBRE FROM PARGEN WHERE PAR_CODCIA = ? "
'Set PS_REP02 = CN.CreateQuery("", pub_cadena)
'Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = wsFECHA1
If wusu = "A" Then
  PS_REP01(2) = LK_CODUSU
End If

llave_rep01.Requery
f1 = 6
f1 = f1 + 1
xl.Cells(f1, 1) = "SALDO INICIAL DE CAJA :"
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = SALDO_ANTERIOR

xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = SALDO_ANTERIOR_D

xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = SALDO_ANTERIOR + redondea(val(SALDO_ANTERIOR_D) * val(WS_TIPO_C))

If llave_rep01.EOF Then
'  MsgBox "No Existe Movimientos", 48, Pub_Titulo
  GoTo PASA01 ' GoTo CANCELA
  Exit Sub
End If

f1 = f1 + 1

ww_fbg = llave_rep01!ALL_FBG
ww_numser = llave_rep01!ALL_NUMSER
ww_serini = llave_rep01!ALL_NUMSER
ww_numfac = llave_rep01!all_numfac
ww_numini = llave_rep01!all_numfac
ww_numfin = llave_rep01!all_numfac
Do Until llave_rep01.EOF
    If ww_fbg <> llave_rep01!ALL_FBG Then
        GoSub IMP_LINEA
        ww_fbg = llave_rep01!ALL_FBG
        ww_numser = llave_rep01!ALL_NUMSER
        ww_serini = llave_rep01!ALL_NUMSER
        ww_numfac = llave_rep01!all_numfac
        ww_numini = llave_rep01!all_numfac
    End If
    If ww_numser <> llave_rep01!ALL_NUMSER Then
        GoSub IMP_LINEA
        ww_numser = llave_rep01!ALL_NUMSER
        ww_serini = llave_rep01!ALL_NUMSER
        ww_numfac = llave_rep01!all_numfac
        ww_numini = llave_rep01!all_numfac
    End If
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
        WMONTO_SOLES = WMONTO_SOLES + val(llave_rep01!ALL_IMPORTE_AMORT)
        WSOLES_DOLAR = WSOLES_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
        If llave_rep01!ALL_SIGNO_CAJA = 0 Then
          WSOLES_CRED = WSOLES_CRED + val(llave_rep01!ALL_IMPORTE_AMORT)
          WSOLES_DOLAR_CRED = WSOLES_DOLAR_CRED + val(llave_rep01!ALL_IMPORTE_AMORT)
        End If
    Else
        WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
        WMONTO_DOLAR = WMONTO_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
        WSOLES_DOLAR = WSOLES_DOLAR + (val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
        If llave_rep01!ALL_SIGNO_CAJA = 0 Then
           WDOLAR_CRED = WDOLAR_CRED + val(llave_rep01!ALL_IMPORTE_AMORT)
           WSOLES_DOLAR_CRED = WSOLES_DOLAR_CRED + (val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
        End If
    End If
    ww_numfin = llave_rep01!all_numfac
    llave_rep01.MoveNext
Loop
GoSub IMP_LINEA
 TT_TOTAL_SOLES = 0
 TT_TOTAL_DOLAR = 0
 TT_DOLAR_SOLES = 0
 f1 = f1 + 1
 xl.Cells(f1 + 1, 4) = "TOTAL VENTA = "
 xl.Cells(f1 + 1, 7) = "S/."
 wran1 = "H" & 8
 wran2 = "H" & f1
 wranF = "H" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_SOLES = TT_TOTAL_SOLES + val(xl.Range(wranF))
 xl.Cells(f1 + 1, 9) = "US$."
 wran1 = "J" & 8
 wran2 = "J" & f1
 wranF = "J" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_DOLAR = TT_TOTAL_DOLAR + val(xl.Range(wranF))
 xl.Cells(f1 + 1, 9) = "S/."
 wran1 = "L" & 8
 wran2 = "L" & f1
 wranF = "L" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_DOLAR_SOLES = TT_DOLAR_SOLES + val(xl.Range(wranF))
 
 wranF = "A" & f1 + 1 & ":L" & f1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 wranF = "A" & f1 + 2 & ":L" & f1 + 2
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
PASA01:
f1 = f1 + 4
xl.Cells(f1, 1) = "(-) CREDITOS "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = WSOLES_CRED
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = WDOLAR_CRED
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = WSOLES_DOLAR_CRED

wranF = "A" & f1 + 1 & ":L" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
f1 = f1 + 3
xl.Cells(f1, 1) = "TOTAL EFECTIVO X VENTA = "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = TT_TOTAL_SOLES - WSOLES_CRED
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = TT_TOTAL_DOLAR - WDOLAR_CRED
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = TT_DOLAR_SOLES - WSOLES_DOLAR_CRED

wranF = "A" & f1 + 1 & ":L" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1

' SELECT PARA EGRESOS DE CAJA
SOLES_TOTAL_CAJA = 0
DOLAR_TOTAL_CAJA = 0
TOTAL_CAJA_SD = 0
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
pub_cadena = "SELECT ALL_FECHA_SUNAT , ALL_CODTRA, ALL_IMPORTE_AMORT,  ALL_FECHA_DIA, ALL_MONEDA_CAJA, ALL_IMPORTE, ALL_TIPO_CAMBIO  FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_SUNAT = ? AND ALL_SIGNO_CAJA = -1 AND ( ALL_CODTRA = 5355  OR ALL_CODTRA = 2748 )  AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 AND ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_DIA"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = wsFECHA1
llave_rep01.Requery
Do Until llave_rep01.EOF
 If llave_rep01!ALL_MONEDA_CAJA = "S" Then
    If llave_rep01!ALL_IMPORTE <> 0 Then
      SOLES_TOTAL_CAJA = SOLES_TOTAL_CAJA + llave_rep01!ALL_IMPORTE
      TOTAL_CAJA_SD = TOTAL_CAJA_SD + llave_rep01!ALL_IMPORTE
    Else
      SOLES_TOTAL_CAJA = SOLES_TOTAL_CAJA + llave_rep01!ALL_IMPORTE_AMORT
      TOTAL_CAJA_SD = TOTAL_CAJA_SD + llave_rep01!ALL_IMPORTE_AMORT
    End If
 Else
    WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
    If llave_rep01!ALL_IMPORTE <> 0 Then
      DOLAR_TOTAL_CAJA = DOLAR_TOTAL_CAJA + llave_rep01!ALL_IMPORTE
      TOTAL_CAJA_SD = TOTAL_CAJA_SD + val(Format((val(llave_rep01!ALL_IMPORTE) * WS_TIPO_C), "0.000")) 'Val(llave_rep01!ALL_TIPO_CAMBIO))
    Else
      DOLAR_TOTAL_CAJA = DOLAR_TOTAL_CAJA + llave_rep01!ALL_IMPORTE_AMORT
      TOTAL_CAJA_SD = TOTAL_CAJA_SD + val(Format((val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C), "0.000")) 'Val(llave_rep01!ALL_TIPO_CAMBIO))
    End If
 End If
llave_rep01.MoveNext
Loop

SOLES_TOTAL_CAJA_I = 0
DOLAR_TOTAL_CAJA_I = 0
TOTAL_CAJA_SD_I = 0
pub_cadena = "SELECT ALL_FECHA_SUNAT , ALL_CODTRA, ALL_IMPORTE_AMORT,  ALL_FECHA_DIA, ALL_MONEDA_CAJA, ALL_IMPORTE, ALL_TIPO_CAMBIO  FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_SUNAT = ? AND ALL_SIGNO_CAJA = 1 AND ALL_CODTRA = 5350 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 AND  ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_DIA"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = wsFECHA1
llave_rep01.Requery
Do Until llave_rep01.EOF
 If llave_rep01!ALL_MONEDA_CAJA = "S" Then
    SOLES_TOTAL_CAJA_I = SOLES_TOTAL_CAJA_I + llave_rep01!ALL_IMPORTE
    TOTAL_CAJA_SD_I = TOTAL_CAJA_SD_I + llave_rep01!ALL_IMPORTE
 Else
    WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
    DOLAR_TOTAL_CAJA_I = DOLAR_TOTAL_CAJA_I + llave_rep01!ALL_IMPORTE
    TOTAL_CAJA_SD_I = TOTAL_CAJA_SD_I + val(Format((val(llave_rep01!ALL_IMPORTE) * WS_TIPO_C), "0.000")) 'Val(llave_rep01!ALL_TIPO_CAMBIO))
 End If
llave_rep01.MoveNext
Loop


FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
' SELECT PARA ADELANTOS
pub_cadena = "SELECT ALL_FECHA_SUNAT, ALL_MONEDA_CLI, ALL_IMPORTE_AMORT, ALL_TIPO_CAMBIO  FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV <> 10 AND ALL_TIPDOC = 'PA' AND ALL_SIGNO_CAJA = 1 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 AND ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = wsFECHA1
llave_rep01.Requery
Do Until llave_rep01.EOF
 If llave_rep01!ALL_MONEDA_CLI = "S" Then
    SOLES_TOTAL_ADEL = SOLES_TOTAL_ADEL + llave_rep01!ALL_IMPORTE_AMORT
    TOTAL_ADEL_SD = TOTAL_ADEL_SD + llave_rep01!ALL_IMPORTE_AMORT
 Else
    WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
    DOLAR_TOTAL_ADEL = DOLAR_TOTAL_ADEL + llave_rep01!ALL_IMPORTE_AMORT
    TOTAL_ADEL_SD = TOTAL_ADEL_SD + val(Format((val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C), "0.000")) 'Val(llave_rep01!ALL_TIPO_CAMBIO))
 End If
  
llave_rep01.MoveNext
Loop

' SELECT PARA COBRANZAS
pub_cadena = "SELECT ALL_FECHA_SUNAT, ALL_MONEDA_CLI, ALL_IMPORTE_AMORT, ALL_TIPO_CAMBIO  FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV <> 10 AND ALL_CODCLIE <> 0 AND ALL_SIGNO_CAJA = 1 AND ALL_TIPDOC <> 'PA' AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 AND ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = wsFECHA1
llave_rep01.Requery
Do Until llave_rep01.EOF
 If llave_rep01!ALL_MONEDA_CLI = "S" Then
    SOLES_TOTAL_COBRA = SOLES_TOTAL_COBRA + llave_rep01!ALL_IMPORTE_AMORT
    TOTAL_COBRA_SD = TOTAL_COBRA_SD + llave_rep01!ALL_IMPORTE_AMORT
 Else
    WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
    DOLAR_TOTAL_COBRA = DOLAR_TOTAL_COBRA + llave_rep01!ALL_IMPORTE_AMORT
    TOTAL_COBRA_SD = TOTAL_COBRA_SD + val(Format((val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C), "0.000")) 'Val(llave_rep01!ALL_TIPO_CAMBIO))
 End If
  
llave_rep01.MoveNext
Loop

' SELECT PARA DEPOSITOS A BANCOS
SOLES_TOTAL_BANCO = 0
DOLAR_TOTAL_BANCO = 0
TOTAL_BANCO_SD = 0
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
pub_cadena = "SELECT ALL_MONEDA_CCM, ALL_FECHA_DIA,  ALL_IMPORTE, ALL_TIPO_CAMBIO  FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_SUNAT = ? AND ALL_SIGNO_CAJA = -1 AND ALL_CODTRA = 5310 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 AND  ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_DIA"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = wsFECHA1
llave_rep01.Requery
Do Until llave_rep01.EOF
 If llave_rep01!ALL_moneda_ccm = "S" Then
    SOLES_TOTAL_BANCO = SOLES_TOTAL_BANCO + llave_rep01!ALL_IMPORTE
    TOTAL_BANCO_SD = TOTAL_BANCO_SD + llave_rep01!ALL_IMPORTE
 Else
    'WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
    DOLAR_TOTAL_BANCO = DOLAR_TOTAL_BANCO + llave_rep01!ALL_IMPORTE
    TOTAL_BANCO_SD = TOTAL_BANCO_SD + val(Format((val(llave_rep01!ALL_IMPORTE) * WS_TIPO_C), "0.000")) 'Val(llave_rep01!ALL_TIPO_CAMBIO))
 End If
llave_rep01.MoveNext
Loop



f1 = f1 + 3
xl.Cells(f1, 1) = "'R E S U M E N:"
f1 = f1 + 2
xl.Cells(f1, 1) = "'(+)COBRANZAS DE CLIENTES "
f1 = f1 + 1
xl.Cells(f1, 1) = "'Ventas al Contado"
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = TT_TOTAL_SOLES - WSOLES_CRED
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = TT_TOTAL_DOLAR - WDOLAR_CRED
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = TT_DOLAR_SOLES - WSOLES_DOLAR_CRED


f1 = f1 + 1
xl.Cells(f1, 1) = "'Cobranza del Dia"
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = SOLES_TOTAL_COBRA
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = DOLAR_TOTAL_COBRA
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = TOTAL_COBRA_SD
f1 = f1 + 1
xl.Cells(f1, 1) = "'Adelantos de Clientes "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = SOLES_TOTAL_ADEL
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = DOLAR_TOTAL_ADEL
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = TOTAL_ADEL_SD



f1 = f1 + 2
xl.Cells(f1, 1) = "'(-)EGRESOS DEL DIA"
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = SOLES_TOTAL_CAJA
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = DOLAR_TOTAL_CAJA
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = TOTAL_CAJA_SD

f1 = f1 + 2
xl.Cells(f1, 1) = "'(+)INGRESOS DEL DIA"
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = SOLES_TOTAL_CAJA_I
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = DOLAR_TOTAL_CAJA_I
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = TOTAL_CAJA_SD_I


f1 = f1 + 2
xl.Cells(f1, 1) = "'(-)DEPOSITO A BANCOS"
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = SOLES_TOTAL_BANCO
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = DOLAR_TOTAL_BANCO
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = TOTAL_BANCO_SD


f1 = f1 + 2
xl.Cells(f1, 1) = "'TOTAL EFECTIVO DEL DIA"
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = (TT_TOTAL_SOLES - WSOLES_CRED) + SOLES_TOTAL_COBRA + SOLES_TOTAL_ADEL - SOLES_TOTAL_CAJA - SOLES_TOTAL_BANCO + SALDO_ANTERIOR + SOLES_TOTAL_CAJA_I
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = (TT_TOTAL_DOLAR - WDOLAR_CRED) + DOLAR_TOTAL_COBRA + DOLAR_TOTAL_ADEL - DOLAR_TOTAL_CAJA - DOLAR_TOTAL_BANCO + SALDO_ANTERIOR_D + DOLAR_TOTAL_CAJA_I
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = (TT_DOLAR_SOLES - WSOLES_DOLAR_CRED) + TOTAL_COBRA_SD + TOTAL_ADEL_SD - TOTAL_CAJA_SD - TOTAL_BANCO_SD + SALDO_ANTERIOR + val(redondea(SALDO_ANTERIOR_D * WS_TIPO_C)) + TOTAL_CAJA_SD_I
wranF = "A" & f1 & ":L" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "A" & f1 + 1 & ":L" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
If LK_FECHA_DIA <> wsFECHA1 Then
  GoTo SIGUE001:
End If

SQ_OPER = 2
PUB_FECHA = LK_FECHA_DIA
pu_codcia = LK_CODCIA
LEER_ALL_LLAVE
If all_menor.EOF = False Then
   PUB_NUM_OPER_XXX = all_menor!ALL_NUMOPER
Else
   PUB_NUM_OPER_XXX = 0
End If


SQ_OPER = 1
PUB_NUMTAB = 0
PUB_CODCIA = LK_CODCIA
PUB_TIPREG = 1000
LEER_TAB_LLAVE
If tab_llave.EOF Then
   tab_llave.AddNew
   tab_llave!TAB_NUMTAB = 0
   tab_llave!TAB_CODCIA = LK_CODCIA
   tab_llave!TAB_TIPREG = 1000

Else
   tab_llave.Edit
End If
tab_llave!tab_NOMLARGO = PUB_NUM_OPER_XXX
tab_llave!tab_nomcorto = Format(LK_FECHA_DIA, "dd/mm/yyyy")
tab_llave!TAB_contable2 = Format(val(xl.Cells(f1, 8)), "#########.000")
tab_llave.Update

SQ_OPER = 1
PUB_NUMTAB = 0
PUB_CODCIA = LK_CODCIA
PUB_TIPREG = 1001
LEER_TAB_LLAVE
If tab_llave.EOF Then
   tab_llave.AddNew
   tab_llave!TAB_NUMTAB = 0
   tab_llave!TAB_CODCIA = LK_CODCIA
   tab_llave!TAB_TIPREG = 1001

Else
   tab_llave.Edit
End If
tab_llave!tab_NOMLARGO = PUB_NUM_OPER_XXX
tab_llave!tab_nomcorto = Format(LK_FECHA_DIA, "dd/mm/yyyy")
tab_llave!TAB_contable2 = Format(val(xl.Cells(f1, 10)), "#########.000") 'Val(xl.Cells(f1, 10))
tab_llave.Update

SIGUE001:

 

    
DoEvents
FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
DoEvents
xcuenta = 1
xl.APPLICATION.Visible = True
xl.Cells(2, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
xl.Cells(4, 10) = "FECHA:"
xl.Cells(4, 11) = "'" & Format(wsFECHA1, "dd mmm yyyy")
xl.DisplayAlerts = False
xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
xl.Worksheets("Hoja1").Protect PUB_CLAVE
xl.APPLICATION.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = False
FrmImp2.ProgBar.Visible = False
Set xl = Nothing
Screen.MousePointer = 0
FrmImp2.Pantalla.Enabled = True
FrmImp2.Pantalla.Caption = "Por &Pantalla"
FrmImp2.lblproceso.Visible = False

Exit Sub

IMP_LINEA:
f1 = f1 + 1
If ww_fbg = "F" Then
 xl.Cells(f1, 1) = "FACTURAS"
ElseIf ww_fbg = "B" Then
 xl.Cells(f1, 1) = "BOLETAS"
ElseIf ww_fbg = "G" Then
 xl.Cells(f1, 1) = "GUIAS"
Else
 xl.Cells(f1, 1) = "PENDIENTES"
End If
xl.Cells(f1, 2) = "Nº "
xl.Cells(f1, 3) = "'" & Format(ww_serini, "000")
xl.Cells(f1, 4) = ww_numini
xl.Cells(f1, 5) = "'al Nº"
xl.Cells(f1, 6) = ww_numfin
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = Format(WMONTO_SOLES, "0.000")
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = Format(WMONTO_DOLAR, "0.000")
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = Format(WSOLES_DOLAR, "0.000")

WMONTO_SOLES = 0
WMONTO_DOLAR = 0
WSOLES_DOLAR = 0
Return



WEXCEL:
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FrmImp2.lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
  DoEvents
  WPAS = "131296"
  'WPAS = PUB_RUTA_OTRO + "CAJA_DET.xls"
  'DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open PUB_RUTA_OTRO & "\CAJA_DET.xls", 0, True, 4, WPAS, WPAS

  'xl.Workbooks.Open , "C:\ADMIN\HERTISA\CAJA_DET.XLS", 0, True, 4, PUB_CLAVE, PUB_CLAVE
Return

Exit Sub

'LLENA_VALOR:
'For I = QJ To WDIF
'  If I >= 3 Then
'    xl.Cells(F1 + 7, I) = Format(LOC_VALOR, "0.000")
'  End If
'Next I
'Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
Exit Sub

End Sub

Public Sub POR_ANULADOS()
Dim AWQ_GASTOS As Currency
Dim AWQ_FLETES As Currency
Dim wsFECHA1
Dim wsFECHA2

Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) <> cop_llave!cop_fecha_proceso Then cheasiento.Value = 0
If CDate(wsFECHA2) <> cop_llave!cop_fecha_proceso2 Then cheasiento.Value = 0
GoSub WEXCEL
pub_cadena = ""
xcuenta = 0

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents

f1 = 5  'Fila Inicial


pub_cadena = "SELECT FAR_ESTADO FROM FACART WHERE FAR_CODCIA = ? AND FAR_FBG= ? AND FAR_NUMSER = ? AND FAR_NUMFAC= ? AND FAR_TIPMOV = 10 AND FAR_ESTADO <> 'E' "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = 0
PS_REP01(3) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT DISTINCT FAR_CODCIA, FAR_MONEDA,FAR_FBG,FAR_NUMSER, FAR_NUMFAC FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?  AND FAR_TIPMOV = 10 AND FAR_MONEDA = '" & Trim(Left(cmdmoneda.Text, 1)) & "' ORDER BY FAR_CODCIA, FAR_FBG,FAR_NUMSER, FAR_NUMFAC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = ""
PS_REP02(1) = LK_FECHA_DIA
PS_REP02(2) = LK_FECHA_DIA
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = wsFECHA1
PS_REP02(2) = wsFECHA2
DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep02.Requery
If llave_rep02.EOF Then
'  GoTo OTRO_DOCUMENTO
End If
FrmImp2.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
If llave_rep02.RowCount <> 0 Then FrmImp2.ProgBar.max = llave_rep02.RowCount

SQ_OPER = 1
pu_cp = "C"


AWQ_GASTOS = 0
AWQ_FLETES = 0
If llave_rep02.EOF Then GoTo CANCELA
Do Until llave_rep02.EOF
'   If Trim(Left(cmdmoneda.Text, 1)) <> Trim(llave_rep02!FAR_MONEDA) Then GoTo SALTARIN
  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = llave_rep02!far_fbg
  PS_REP01(2) = llave_rep02!far_numser
  PS_REP01(3) = llave_rep02!far_numfac
  llave_rep01.Requery
  If llave_rep01.EOF Then
      AWQ_GASTOS = AWQ_GASTOS + 1
  Else
      AWQ_FLETES = AWQ_FLETES + 1
   End If
SALTARIN:
 llave_rep02.MoveNext
Loop

  f1 = f1 + 2
  xl.Cells(f1, 1) = AWQ_GASTOS
  xl.Cells(f1, 2) = AWQ_FLETES
  xl.Cells(f1, 3) = AWQ_FLETES + AWQ_GASTOS
  If (AWQ_FLETES + AWQ_GASTOS) <> 0 Then
    xl.Cells(f1, 4) = (AWQ_GASTOS * 100) / (AWQ_FLETES + AWQ_GASTOS)
  End If

  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'Procesados del " & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
  If Left(cmdmoneda.Text, 1) = "S" Then
   xl.Cells(4, 1) = "'Documentos en Soles"
  Else
   xl.Cells(4, 1) = "'Documentos en Dolares"
  End If
  
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect PUB_CLAVE
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub

IMPRI_FAC:

Return

CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\POR_ANULADOS.xls", 0, True, 4, PUB_CLAVE, PUB_CLAVE
Return

TOTAL_DIA:
Return




LETRAS:

Return

FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
 GoTo CANCELA
 Resume Next

End Sub

Public Sub RESU_CAJA()
FrmImp2.Pantalla.Enabled = False
FrmImp2.cerrar.Enabled = False
Dim WIMPORTE As Currency
Dim COBRA_EFEC_SOLES As Currency
Dim COBRA_EFEC_DOLAR As Currency
Dim COBRA_EFEC_SOLES_DOLAR As Currency


Dim VENTA_TOTAL_SOLES As Currency
Dim VENTA_TOTAL_DOLAR As Currency
Dim VENTA_TOTAL_SOLES_DOLAR As Currency
Dim K_DET_SOLES As Currency
Dim K_DET_DOLAR As Currency
Dim K_DET_SOLES_DOLAR As Currency

Dim wnumfac As Currency
Dim wusu As String * 1
Dim ww_fbg
Dim ww_numser
Dim ww_numfac
Dim ww_numini
Dim ww_serini
Dim ww_numfin
Dim WMONTO_SOLES As Currency
Dim WMONTO_DOLAR As Currency
Dim wsFECHA1
Dim WDOLAR_CRED As Currency
Dim WSOLES_CRED As Currency
Dim TT_TOTAL_SOLES As Currency
Dim TT_TOTAL_DOLAR  As Currency
Dim WSOLES_DOLAR As Currency
Dim TT_DOLAR_SOLES As Currency
Dim WSOLES_DOLAR_CRED As Currency
Dim WSOLES_DOLAR_TOT As Currency
Dim SOLES_TOTAL_COBRA As Currency
Dim DOLAR_TOTAL_COBRA As Currency
Dim TOTAL_COBRA_SD As Currency
Dim WS_TIPO_C As Currency

If Right(txtFecha.Text, 2) = "__" Then
     wsFECHA1 = Left(txtFecha.Text, 8)
Else
     wsFECHA1 = Trim(txtFecha.Text)
End If
If Not IsDate(wsFECHA1) Then
 FrmImp2.Pantalla.Enabled = True
 FrmImp2.cerrar.Enabled = True
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If


FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
f1 = 5
pub_mensaje = "Imprimir según su Usuario...?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbYes Then
  wusu = "A"
Else
  wusu = " "
End If

If wusu = "A" Then
   pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10 AND ALL_CODUSU = ? AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
Else
   pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10   AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
  PS_REP01(0) = 0
  PS_REP01(1) = 0
Else
  PS_REP01(0) = 0
  PS_REP01(1) = 0
End If
PS_REP01(2) = LK_FECHA_DIA
If wusu = "A" Then
  PS_REP01(3) = 0
End If
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
  PS_REP01(0) = "01"
  PS_REP01(1) = "02"
Else
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If


llave_rep01.Requery

If llave_rep01.EOF Then
  MsgBox "No Existe Movimientos", 48, Pub_Titulo
  GoTo CANCELA
  Exit Sub
End If
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.max = llave_rep01.RowCount
f1 = 6
ww_fbg = llave_rep01!ALL_FBG
ww_numser = llave_rep01!ALL_NUMSER
ww_serini = llave_rep01!ALL_NUMSER
ww_numfac = llave_rep01!all_numfac
ww_numini = llave_rep01!all_numfac
ww_numfin = llave_rep01!all_numfac
wnumfac = llave_rep01!all_numfac
Do Until llave_rep01.EOF
    FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
    If ww_fbg <> llave_rep01!ALL_FBG Then
        GoSub IMP_LINEA
        ww_fbg = llave_rep01!ALL_FBG
        ww_numser = llave_rep01!ALL_NUMSER
        ww_serini = llave_rep01!ALL_NUMSER
        ww_numfac = llave_rep01!all_numfac
        ww_numini = llave_rep01!all_numfac
        wnumfac = llave_rep01!all_numfac
    End If
    If ww_numser <> llave_rep01!ALL_NUMSER Then
        GoSub IMP_LINEA
        ww_numser = llave_rep01!ALL_NUMSER
        ww_serini = llave_rep01!ALL_NUMSER
        ww_numfac = llave_rep01!all_numfac
        ww_numini = llave_rep01!all_numfac
        wnumfac = llave_rep01!all_numfac
    End If
    If wnumfac <> llave_rep01!all_numfac Then
      
   '   MsgBox "Falta el Nro. :" & wnumfac
   '   wnumfac = llave_rep01!all_numfac
    End If
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
        WMONTO_SOLES = WMONTO_SOLES + val(llave_rep01!ALL_IMPORTE_AMORT)
        WSOLES_DOLAR = WSOLES_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
        If llave_rep01!ALL_SIGNO_CAJA = 0 Then
          WSOLES_CRED = WSOLES_CRED + val(llave_rep01!ALL_IMPORTE_AMORT)
          WSOLES_DOLAR_CRED = WSOLES_DOLAR_CRED + val(llave_rep01!ALL_IMPORTE_AMORT)
        End If
    Else
        WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
        WMONTO_DOLAR = redondea(WMONTO_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT))
        WSOLES_DOLAR = WSOLES_DOLAR + redondea(val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
        If llave_rep01!ALL_SIGNO_CAJA = 0 Then
          WDOLAR_CRED = WDOLAR_CRED + val(llave_rep01!ALL_IMPORTE_AMORT)
          WSOLES_DOLAR_CRED = WSOLES_DOLAR_CRED + redondea(val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
        End If
        
    End If
    ww_numfin = llave_rep01!all_numfac
    wnumfac = wnumfac + 1
    llave_rep01.MoveNext
Loop
 GoSub IMP_LINEA
 TT_TOTAL_SOLES = 0
 TT_TOTAL_DOLAR = 0
 TT_DOLAR_SOLES = 0
 f1 = f1 + 1
 xl.Cells(f1 + 1, 4) = "TOTAL VENTA = "
 xl.Cells(f1 + 1, 7) = "S/."
 wran1 = "H" & 6
 wran2 = "H" & f1
 wranF = "H" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_SOLES = TT_TOTAL_SOLES + val(xl.Range(wranF))
 xl.Cells(f1 + 1, 9) = "US$."
 wran1 = "J" & 6
 wran2 = "J" & f1
 wranF = "J" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_DOLAR = TT_TOTAL_DOLAR + val(xl.Range(wranF))
 xl.Cells(f1 + 1, 9) = "S/."
 wran1 = "L" & 6
 wran2 = "L" & f1
 wranF = "L" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_DOLAR_SOLES = TT_DOLAR_SOLES + val(xl.Range(wranF))
 
 wranF = "A" & f1 + 1 & ":L" & f1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 wranF = "A" & f1 + 2 & ":L" & f1 + 2
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 

f1 = f1 + 3
xl.Cells(f1 - 1, 1) = "(-) CREDITOS "
' A SOLO FIRMA AL CREDITO
xl.Cells(f1, 1) = "A SOLO FIRMA:"
llave_rep01.MoveFirst
Do Until llave_rep01.EOF
 If llave_rep01!ALL_SIGNO_CAJA <> 0 Then GoTo PASA_CRED
 If llave_rep01!all_SECUENCIA <> 24 Then GoTo PASA_CRED
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    pu_cp = llave_rep01!ALL_CP
    pu_codclie = val(llave_rep01!ALL_CODCLIE)
    LEER_CLI_LLAVE
    xl.Cells(f1, 1) = Trim(cli_llave!CLI_NOMBRE)
    xl.Cells(f1, 2) = llave_rep01!ALL_FBG
    xl.Cells(f1, 3) = llave_rep01!ALL_NUMSER
    xl.Cells(f1, 4) = llave_rep01!all_numfac
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
        xl.Cells(f1, 7) = "'S/."
        xl.Cells(f1, 8) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000")
    Else
        xl.Cells(f1, 9) = "'US$."
        xl.Cells(f1, 10) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000")
    End If
PASA_CRED:
llave_rep01.MoveNext
Loop
' CON CHEQUE AL BANCO
f1 = f1 + 2
xl.Cells(f1, 1) = "CON CHEQUE:"
llave_rep01.MoveFirst
Do Until llave_rep01.EOF
 If llave_rep01!ALL_SIGNO_CAJA <> 0 Then GoTo PASA_CHE
 If llave_rep01!all_SECUENCIA <> 50 Then GoTo PASA_CHE
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    pu_cp = llave_rep01!ALL_CP
    pu_codclie = val(llave_rep01!ALL_CODCLIE)
    LEER_CLI_LLAVE
    xl.Cells(f1, 1) = Trim(cli_llave!CLI_NOMBRE)
    xl.Cells(f1, 2) = llave_rep01!ALL_FBG
    xl.Cells(f1, 3) = llave_rep01!ALL_NUMSER
    xl.Cells(f1, 4) = llave_rep01!all_numfac
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
        xl.Cells(f1, 7) = "'S/."
        xl.Cells(f1, 8) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000")
    Else
        xl.Cells(f1, 9) = "'US$."
        xl.Cells(f1, 10) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000")
    End If
PASA_CHE:
llave_rep01.MoveNext
Loop
 
f1 = f1 + 2
xl.Cells(f1, 1) = "TOTAL CREDITOS "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = WSOLES_CRED
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = WDOLAR_CRED
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = WSOLES_DOLAR_CRED

wranF = "A" & f1 & ":L" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "A" & f1 + 1 & ":L" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1

' DETERMINACION DE EFECTIVO NETO
f1 = f1 + 2
xl.Cells(f1, 1) = "'VENTAS CONTADO"
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = TT_TOTAL_SOLES - WSOLES_CRED
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = TT_TOTAL_DOLAR - WDOLAR_CRED
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = TT_DOLAR_SOLES - WSOLES_DOLAR_CRED
wranF = "A" & f1 & ":L" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "A" & f1 + 1 & ":L" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1

K_DET_SOLES = 0
K_DET_DOLAR = 0
K_DET_SOLES_DOLAR = 0

f1 = f1 + 1
xl.Cells(f1, 1) = "DETERMINACION DE EFECTIVO NETO:"
f1 = f1 + 1
xl.Cells(f1, 1) = "CHEQUES RECIBIDOS:"
xl.Cells(f1, 5) = "NRO."
xl.Cells(f1, 6) = "BANCO"
llave_rep01.MoveFirst
Do Until llave_rep01.EOF
 If llave_rep01!ALL_SIGNO_CAJA <> 1 Then GoTo PASA_CHE2
 If llave_rep01!all_SECUENCIA <> 30 Then GoTo PASA_CHE2
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    pu_cp = llave_rep01!ALL_CP
    pu_codclie = val(llave_rep01!ALL_CODCLIE)
    LEER_CLI_LLAVE
    xl.Cells(f1, 1) = Trim(cli_llave!CLI_NOMBRE)
    xl.Cells(f1, 2) = llave_rep01!ALL_FBG
    xl.Cells(f1, 3) = llave_rep01!ALL_NUMSER
    xl.Cells(f1, 4) = llave_rep01!all_numfac
    xl.Cells(f1, 5) = llave_rep01!all_chenum
    PUB_CODBAN = llave_rep01!all_codban
    pu_codcia = LK_CODCIA
    LEER_CCM_LLAVE
    xl.Cells(f1, 6) = Left(ccm_llave!CCM_NOMBRE, 10)
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
        xl.Cells(f1, 7) = "'S/."
        xl.Cells(f1, 8) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000")
        K_DET_SOLES = K_DET_SOLES + val(llave_rep01!ALL_IMPORTE_AMORT)
        K_DET_SOLES_DOLAR = K_DET_SOLES_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
    Else
        WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
        xl.Cells(f1, 9) = "'US$."
        xl.Cells(f1, 10) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000")
        K_DET_DOLAR = K_DET_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
        K_DET_SOLES_DOLAR = K_DET_SOLES_DOLAR + redondea(val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
    End If
PASA_CHE2:
llave_rep01.MoveNext
Loop
If K_DET_SOLES_DOLAR <> 0 Then
 f1 = f1 + 1
 xl.Cells(f1, 6) = "TOTAL = "
 xl.Cells(f1, 7) = "'S/."
 xl.Cells(f1, 8) = K_DET_SOLES
 xl.Cells(f1, 9) = "'US$."
 xl.Cells(f1, 10) = K_DET_DOLAR
' xl.Cells(F1, 11) = "'S/."
' xl.Cells(F1, 12) = K_DET_SOLES_DOLAR
 wranF = "F" & f1 & ":J" & f1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 wranF = "F" & f1 + 1 & ":J" & f1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
End If

 
f1 = f1 + 2
xl.Cells(f1, 1) = "DEPOSITO:"
xl.Cells(f1, 5) = "OP."
xl.Cells(f1, 6) = "BANCO"
llave_rep01.MoveFirst
Do Until llave_rep01.EOF
 If llave_rep01!ALL_SIGNO_CAJA <> 1 Then GoTo PASA_CHE3
 If llave_rep01!all_SECUENCIA <> 40 Then GoTo PASA_CHE3
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    pu_cp = llave_rep01!ALL_CP
    pu_codclie = val(llave_rep01!ALL_CODCLIE)
    LEER_CLI_LLAVE
    xl.Cells(f1, 1) = Trim(cli_llave!CLI_NOMBRE)
    xl.Cells(f1, 2) = llave_rep01!ALL_FBG
    xl.Cells(f1, 3) = llave_rep01!ALL_NUMSER
    xl.Cells(f1, 4) = llave_rep01!all_numfac
    xl.Cells(f1, 5) = llave_rep01!all_chenum
    PUB_CODBAN = llave_rep01!all_codban
    pu_codcia = llave_rep01!all_CODCIA
    LEER_CCM_LLAVE
    xl.Cells(f1, 6) = Left(ccm_llave!CCM_NOMBRE, 10)
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
        xl.Cells(f1, 7) = "'S/."
        xl.Cells(f1, 8) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000")
        K_DET_SOLES = K_DET_SOLES + val(llave_rep01!ALL_IMPORTE_AMORT)
        K_DET_SOLES_DOLAR = K_DET_SOLES_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
    Else
        WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
        xl.Cells(f1, 9) = "'US$."
        xl.Cells(f1, 10) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000")
        K_DET_DOLAR = K_DET_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
        K_DET_SOLES_DOLAR = K_DET_SOLES_DOLAR + redondea(val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
    End If
PASA_CHE3:
llave_rep01.MoveNext
Loop
If K_DET_SOLES_DOLAR <> 0 Then
 f1 = f1 + 1
 xl.Cells(f1, 6) = "TOTAL = "
 xl.Cells(f1, 7) = "'S/."
 xl.Cells(f1, 8) = K_DET_SOLES
 xl.Cells(f1, 9) = "'US$."
 xl.Cells(f1, 10) = K_DET_DOLAR
' xl.Cells(F1, 11) = "'S/."
' xl.Cells(F1, 12) = K_DET_SOLES_DOLAR
 wranF = "F" & f1 & ":J" & f1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 wranF = "F" & f1 + 1 & ":J" & f1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
End If

f1 = f1 + 2
xl.Cells(f1, 1) = "TOTAL VENTAS EN EFECTIVO= "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = TT_TOTAL_SOLES - WSOLES_CRED - K_DET_SOLES
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = TT_TOTAL_DOLAR - WDOLAR_CRED - K_DET_DOLAR
'xl.Cells(F1, 11) = "'S/."
'xl.Cells(F1, 12) = TT_DOLAR_SOLES - WSOLES_DOLAR_CRED '+ K_DET_SOLES_DOLAR

VENTA_TOTAL_SOLES = TT_TOTAL_SOLES - WSOLES_CRED - K_DET_SOLES
VENTA_TOTAL_DOLAR = TT_TOTAL_DOLAR - WDOLAR_CRED - K_DET_DOLAR
VENTA_TOTAL_SOLES_DOLAR = TT_DOLAR_SOLES - WSOLES_DOLAR_CRED ' K_DET_SOLES_DOLAR

wranF = "A" & f1 & ":J" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "A" & f1 + 1 & ":J" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1


f1 = f1 + 2
xl.Cells(f1, 1) = "'(+)COBRANZAS "
f1 = f1 + 1
xl.Cells(f1, 1) = "'EFECTIVO: "
xl.Cells(f1, 6) = "RECIBO"

' SELECT PARA COBRANZAS DE EFECTIVO
If wusu = "A" Then
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_CODUSU = ? AND ALL_TIPMOV <> 10 AND ALL_SIGNO_CAJA = 1 AND all_flag_ext <> 'E' AND ALL_SIGNO_CAR <> 0 AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
Else
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ?  AND ALL_TIPMOV <> 10 AND ALL_SIGNO_CAJA = 1 AND all_flag_ext <> 'E' AND ALL_SIGNO_CAR <> 0  AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
 PS_REP01(0) = 0
 PS_REP01(1) = 0
Else
 PS_REP01(0) = 0
 PS_REP01(1) = 0
End If
PS_REP01(2) = LK_FECHA_DIA
If wusu = "A" Then
  PS_REP01(3) = 0
End If
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
 PS_REP01(0) = "01"
 PS_REP01(1) = "02"
Else
 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If

llave_rep01.Requery
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Min = 0
If Not llave_rep01.EOF Then FrmImp2.ProgBar.max = llave_rep01.RowCount
Do Until llave_rep01.EOF
    FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
    If llave_rep01!all_SECUENCIA = 30 Or llave_rep01!all_SECUENCIA = 40 Then
      GoTo PASA_COBRA1
    End If
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = llave_rep01!all_CODCIA
    pu_cp = llave_rep01!ALL_CP
    pu_codclie = val(llave_rep01!ALL_CODCLIE)
    LEER_CLI_LLAVE
    xl.Cells(f1, 1) = Trim(cli_llave!CLI_NOMBRE)
    xl.Cells(f1, 2) = llave_rep01!ALL_FBG
    xl.Cells(f1, 3) = llave_rep01!all_numser_c
    xl.Cells(f1, 4) = llave_rep01!all_numfac_c
    xl.Cells(f1, 6) = llave_rep01!ALL_NUM_RECIBO
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
       WIMPORTE = llave_rep01!ALL_IMPORTE_AMORT
       If llave_rep01!ALL_MONEDA_CLI = "D" Then
         WIMPORTE = redondea(WIMPORTE * llave_rep01!ALL_TIPO_CAMBIO)
       End If
       xl.Cells(f1, 7) = "'S/."
       xl.Cells(f1, 8) = Format(WIMPORTE, "0.000")
       SOLES_TOTAL_COBRA = SOLES_TOTAL_COBRA + WIMPORTE
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + WIMPORTE
    Else
       WIMPORTE = llave_rep01!ALL_IMPORTE_AMORT
       If llave_rep01!ALL_MONEDA_CLI = "S" Then
         WIMPORTE = redondea(llave_rep01!ALL_IMPORTE_AMORT / llave_rep01!ALL_TIPO_CAMBIO)
       End If
       xl.Cells(f1, 9) = "'US$."
       xl.Cells(f1, 10) = Format(WIMPORTE, "0.000")
       WS_TIPO_C = llave_rep01!ALL_TIPO_CAMBIO
       DOLAR_TOTAL_COBRA = DOLAR_TOTAL_COBRA + WIMPORTE
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + val(Format((WIMPORTE * WS_TIPO_C), "0.000")) 'Val(llave_rep01!ALL_TIPO_CAMBIO))
    End If
PASA_COBRA1:
llave_rep01.MoveNext
Loop
If TOTAL_COBRA_SD <> 0 Then
  f1 = f1 + 1
  xl.Cells(f1, 6) = "'TOTAL = "
  xl.Cells(f1, 7) = "'S/."
  xl.Cells(f1, 8) = SOLES_TOTAL_COBRA
  xl.Cells(f1, 9) = "'US$."
  xl.Cells(f1, 10) = DOLAR_TOTAL_COBRA
'  xl.Cells(F1, 11) = "'S/."
 ' xl.Cells(F1, 12) = TOTAL_COBRA_SD
  wranF = "F" & f1 & ":J" & f1
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  wranF = "F" & f1 + 1 & ":J" & f1 + 1
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
End If
COBRA_EFEC_SOLES = SOLES_TOTAL_COBRA
COBRA_EFEC_DOLAR = DOLAR_TOTAL_COBRA
COBRA_EFEC_SOLES_DOLAR = TOTAL_COBRA_SD


Dim K_COBRA_SOLES As Currency
Dim K_COBRA_DOLAR As Currency
Dim K_COBRA_SOLES_DOLAR As Currency
Dim K2_COBRA_SOLES As Currency
Dim K2_COBRA_DOLAR As Currency
Dim K2_COBRA_SOLES_DOLAR As Currency

K_COBRA_SOLES_DOLAR = 0
K_COBRA_SOLES = 0
K_COBRA_DOLAR = 0
K2_COBRA_SOLES_DOLAR = 0
K2_COBRA_SOLES = 0
K2_COBRA_DOLAR = 0

f1 = f1 + 2
xl.Cells(f1, 1) = "'CHEQUE: "
xl.Cells(f1, 5) = "NRO."
xl.Cells(f1, 6) = "BANCO"
' COBRANZA CON CHEQUE
llave_rep01.MoveFirst
Do Until llave_rep01.EOF
'    FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
    If llave_rep01!all_SECUENCIA <> 30 Then GoTo PASA_COBRA2
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = llave_rep01!all_CODCIA
    pu_cp = llave_rep01!ALL_CP
    pu_codclie = val(llave_rep01!ALL_CODCLIE)
    LEER_CLI_LLAVE
    xl.Cells(f1, 1) = Trim(cli_llave!CLI_NOMBRE)
    xl.Cells(f1, 2) = llave_rep01!ALL_FBG
    xl.Cells(f1, 3) = llave_rep01!all_numser_c
    xl.Cells(f1, 4) = llave_rep01!all_numfac_c
    PUB_CODBAN = llave_rep01!all_codban
    pu_codcia = llave_rep01!all_CODCIA
    LEER_CCM_LLAVE
    xl.Cells(f1, 5) = llave_rep01!all_chenum
    xl.Cells(f1, 6) = Left(ccm_llave!CCM_NOMBRE, 10)
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
       xl.Cells(f1, 7) = "'S/."
       WIMPORTE = llave_rep01!ALL_IMPORTE_AMORT
       If llave_rep01!ALL_MONEDA_CLI = "D" Then
         WIMPORTE = redondea(WIMPORTE * llave_rep01!ALL_TIPO_CAMBIO)
       End If
       xl.Cells(f1, 8) = Format(WIMPORTE, "0.000")
       SOLES_TOTAL_COBRA = SOLES_TOTAL_COBRA + WIMPORTE
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + WIMPORTE
       K_COBRA_SOLES = K_COBRA_SOLES + val(WIMPORTE)
       K_COBRA_SOLES_DOLAR = K_COBRA_SOLES_DOLAR + val(WIMPORTE)
    Else
       WIMPORTE = llave_rep01!ALL_IMPORTE_AMORT
       If llave_rep01!ALL_MONEDA_CLI = "S" Then
         WIMPORTE = redondea(llave_rep01!ALL_IMPORTE_AMORT / llave_rep01!ALL_TIPO_CAMBIO)
       End If
       xl.Cells(f1, 9) = "'US$."
       xl.Cells(f1, 10) = Format(WIMPORTE, "0.000")
       WS_TIPO_C = llave_rep01!ALL_TIPO_CAMBIO
       DOLAR_TOTAL_COBRA = DOLAR_TOTAL_COBRA + WIMPORTE
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + val(Format((WIMPORTE * WS_TIPO_C), "0.000")) 'Val(llave_rep01!ALL_TIPO_CAMBIO))
       K_COBRA_DOLAR = K_COBRA_DOLAR + WIMPORTE
       K_COBRA_SOLES_DOLAR = K_COBRA_SOLES_DOLAR + val(Format((WIMPORTE * WS_TIPO_C), "0.000"))
 End If
PASA_COBRA2:
llave_rep01.MoveNext
Loop

If K_COBRA_SOLES_DOLAR <> 0 Then
  f1 = f1 + 1
  xl.Cells(f1, 6) = "'TOTAL = "
  xl.Cells(f1, 7) = "'S/."
  xl.Cells(f1, 8) = K_COBRA_SOLES
  xl.Cells(f1, 9) = "'US$."
  xl.Cells(f1, 10) = K_COBRA_DOLAR
 ' xl.Cells(F1, 11) = "'S/."
 ' xl.Cells(F1, 12) = K_COBRA_SOLES_DOLAR
  wranF = "F" & f1 & ":J" & f1
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  wranF = "F" & f1 + 1 & ":J" & f1 + 1
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
End If

f1 = f1 + 2
xl.Cells(f1, 1) = "'DEPOSITO: "
xl.Cells(f1, 5) = "OP."
xl.Cells(f1, 6) = "BANCO"
' COBRANZA CON CHEQUE
llave_rep01.MoveFirst
Do Until llave_rep01.EOF
'    FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
    If llave_rep01!all_SECUENCIA <> 40 Then GoTo PASA_COBRA3
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = llave_rep01!all_CODCIA
    pu_cp = llave_rep01!ALL_CP
    pu_codclie = val(llave_rep01!ALL_CODCLIE)
    LEER_CLI_LLAVE
    xl.Cells(f1, 1) = Trim(cli_llave!CLI_NOMBRE)
    xl.Cells(f1, 2) = llave_rep01!ALL_FBG
    xl.Cells(f1, 3) = llave_rep01!all_numser_c
    xl.Cells(f1, 4) = llave_rep01!all_numfac_c
    PUB_CODBAN = llave_rep01!all_codban
    pu_codcia = llave_rep01!all_CODCIA
    LEER_CCM_LLAVE
    xl.Cells(f1, 5) = llave_rep01!all_chenum
    xl.Cells(f1, 6) = Left(ccm_llave!CCM_NOMBRE, 10)
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
       WIMPORTE = llave_rep01!ALL_IMPORTE_AMORT
       If llave_rep01!ALL_MONEDA_CLI = "D" Then
         WIMPORTE = redondea(WIMPORTE * llave_rep01!ALL_TIPO_CAMBIO)
       End If
       xl.Cells(f1, 7) = "'S/."
       xl.Cells(f1, 8) = Format(WIMPORTE, "0.000")
       SOLES_TOTAL_COBRA = SOLES_TOTAL_COBRA + WIMPORTE
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + WIMPORTE
       K2_COBRA_SOLES = K2_COBRA_SOLES + WIMPORTE
       K2_COBRA_SOLES_DOLAR = K2_COBRA_SOLES_DOLAR + WIMPORTE
    Else
        WIMPORTE = llave_rep01!ALL_IMPORTE_AMORT
        If llave_rep01!ALL_MONEDA_CLI = "S" Then
         WIMPORTE = redondea(llave_rep01!ALL_IMPORTE_AMORT / llave_rep01!ALL_TIPO_CAMBIO)
        End If
        xl.Cells(f1, 9) = "'US$."
        xl.Cells(f1, 10) = Format(WIMPORTE, "0.000")
        WS_TIPO_C = llave_rep01!ALL_TIPO_CAMBIO
        DOLAR_TOTAL_COBRA = DOLAR_TOTAL_COBRA + WIMPORTE
        TOTAL_COBRA_SD = TOTAL_COBRA_SD + val(Format((WIMPORTE * WS_TIPO_C), "0.000"))
        K2_COBRA_DOLAR = K2_COBRA_DOLAR + val(WIMPORTE)
        K2_COBRA_SOLES_DOLAR = K2_COBRA_SOLES_DOLAR + val(Format((WIMPORTE * WS_TIPO_C), "0.000"))
    End If
 
PASA_COBRA3:
llave_rep01.MoveNext
Loop

If K2_COBRA_SOLES_DOLAR <> 0 Then
  f1 = f1 + 1
  xl.Cells(f1, 6) = "'TOTAL = "
  xl.Cells(f1, 7) = "'S/."
  xl.Cells(f1, 8) = K2_COBRA_SOLES
  xl.Cells(f1, 9) = "'US$."
  xl.Cells(f1, 10) = K2_COBRA_DOLAR
  'xl.Cells(F1, 11) = "'S/."
  'xl.Cells(F1, 12) = K2_COBRA_SOLES_DOLAR
  wranF = "F" & f1 & ":J" & f1
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  wranF = "F" & f1 + 1 & ":J" & f1 + 1
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
End If


f1 = f1 + 2
xl.Cells(f1, 1) = "'TOTAL COBRANZAS"
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = SOLES_TOTAL_COBRA
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = DOLAR_TOTAL_COBRA
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = TOTAL_COBRA_SD
wranF = "A" & f1 & ":L" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "A" & f1 + 1 & ":L" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1

f1 = f1 + 2
xl.Cells(f1, 1) = "'TOTAL EFECTIVO RECIBIDO"
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = VENTA_TOTAL_SOLES + COBRA_EFEC_SOLES
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = VENTA_TOTAL_DOLAR + COBRA_EFEC_DOLAR
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = VENTA_TOTAL_SOLES_DOLAR + COBRA_EFEC_SOLES_DOLAR + K_COBRA_SOLES_DOLAR + K2_COBRA_SOLES_DOLAR
wranF = "A" & f1 & ":L" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "A" & f1 + 1 & ":L" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1

' COMPRA DE MONEDAS
If wusu = "A" Then
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_CODUSU = ? AND ALL_CODTRA = 5345 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
Else
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ?  AND ALL_CODTRA = 5345 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
 PS_REP01(0) = 0
 PS_REP01(1) = 0
Else
 PS_REP01(0) = 0
 PS_REP01(1) = 0
End If
PS_REP01(2) = LK_FECHA_DIA
If wusu = "A" Then
  PS_REP01(3) = 0
End If
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
 PS_REP01(0) = "01"
 PS_REP01(1) = "02"
Else
 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If

llave_rep01.Requery
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Min = 0
If Not llave_rep01.EOF Then FrmImp2.ProgBar.max = llave_rep01.RowCount
f1 = f1 + 2
xl.Cells(f1, 1) = "'COMPRAS DE MONEDAS: "
xl.Cells(f1, 5) = ""
xl.Cells(f1, 6) = "T.C."

DOLAR_TOTAL_COBRA = 0
SOLES_TOTAL_COBRA = 0
TOTAL_COBRA_SD = 0
Do Until llave_rep01.EOF
    f1 = f1 + 1
    xl.Cells(f1, 2) = llave_rep01!ALL_FBG
    xl.Cells(f1, 3) = llave_rep01!all_numser_c
    xl.Cells(f1, 4) = llave_rep01!all_numfac_c
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
       xl.Cells(f1, 1) = "COMPRA S/."
       xl.Cells(f1, 6) = Format(llave_rep01!ALL_TIPO_CAMBIO, "0.0000")
       xl.Cells(f1, 7) = "'S/."
       xl.Cells(f1, 8) = Format(llave_rep01!ALL_IMPORTE, "0.000")
       xl.Cells(f1, 9) = "'US$."
       xl.Cells(f1, 10) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000") * -1
       SOLES_TOTAL_COBRA = SOLES_TOTAL_COBRA + val(llave_rep01!ALL_IMPORTE)
       DOLAR_TOTAL_COBRA = DOLAR_TOTAL_COBRA - val(llave_rep01!ALL_IMPORTE_AMORT)
       WS_TIPO_C = llave_rep01!ALL_TIPO_CAMBIO
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + val(llave_rep01!ALL_IMPORTE)
       TOTAL_COBRA_SD = TOTAL_COBRA_SD - redondea(val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
    Else
       xl.Cells(f1, 1) = "COMPRA US$."
       xl.Cells(f1, 6) = Format(llave_rep01!ALL_TIPO_CAMBIO, "0.0000")
       xl.Cells(f1, 7) = "'S/."
       xl.Cells(f1, 8) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000") * -1
       xl.Cells(f1, 9) = "'US$."
       xl.Cells(f1, 10) = Format(llave_rep01!ALL_IMPORTE, "0.000")
       SOLES_TOTAL_COBRA = SOLES_TOTAL_COBRA - val(llave_rep01!ALL_IMPORTE_AMORT)
       DOLAR_TOTAL_COBRA = DOLAR_TOTAL_COBRA + val(llave_rep01!ALL_IMPORTE)
       WS_TIPO_C = llave_rep01!ALL_TIPO_CAMBIO
       TOTAL_COBRA_SD = TOTAL_COBRA_SD - val(llave_rep01!ALL_IMPORTE_AMORT)
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + redondea(val(llave_rep01!ALL_IMPORTE) * WS_TIPO_C)
 End If

llave_rep01.MoveNext
Loop
If SOLES_TOTAL_COBRA = 0 And DOLAR_TOTAL_COBRA = 0 Then

Else
 f1 = f1 + 1
 xl.Cells(f1, 6) = "'TOTAL = "
 xl.Cells(f1, 7) = "'S/."
 xl.Cells(f1, 8) = SOLES_TOTAL_COBRA
 xl.Cells(f1, 9) = "'US$."
 xl.Cells(f1, 10) = DOLAR_TOTAL_COBRA
' xl.Cells(F1, 11) = "'S/."
' xl.Cells(F1, 12) = TOTAL_COBRA_SD
 wranF = "A" & f1 & ":J" & f1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 wranF = "A" & f1 + 1 & ":J" & f1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
End If

   
' INGRESO VARIOS DE CAJA
If wusu = "A" Then
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_CODUSU = ? AND ALL_CODTRA = 5350 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
Else
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ?  AND ALL_CODTRA = 5350 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
 PS_REP01(0) = 0
 PS_REP01(1) = 0
Else
 PS_REP01(0) = 0
 PS_REP01(1) = 0
End If
PS_REP01(2) = LK_FECHA_DIA
If wusu = "A" Then
  PS_REP01(3) = 0
End If
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
 PS_REP01(0) = "01"
 PS_REP01(1) = "02"
Else
 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If

llave_rep01.Requery
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Min = 0
If Not llave_rep01.EOF Then FrmImp2.ProgBar.max = llave_rep01.RowCount
f1 = f1 + 2
xl.Cells(f1, 1) = "'INGRESOS VARIOS: "
xl.Cells(f1, 5) = ""
xl.Cells(f1, 6) = ""
K_COBRA_SOLES = 0
K_COBRA_DOLAR = 0
K_COBRA_SOLES_DOLAR = 0
Do Until llave_rep01.EOF
   f1 = f1 + 1
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
       xl.Cells(f1, 1) = Left(llave_rep01!all_concepto, 20)
       xl.Cells(f1, 7) = "'S/."
       xl.Cells(f1, 8) = Format(llave_rep01!ALL_IMPORTE, "0.000")
       K_COBRA_SOLES = K_COBRA_SOLES + val(llave_rep01!ALL_IMPORTE)
       K_COBRA_SOLES_DOLAR = K_COBRA_SOLES_DOLAR + val(llave_rep01!ALL_IMPORTE)
    Else
       WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
       xl.Cells(f1, 1) = "COMPRA US$."
       xl.Cells(f1, 9) = "'US$."
       xl.Cells(f1, 10) = Format(llave_rep01!ALL_IMPORTE, "0.000")
       K_COBRA_DOLAR = K_COBRA_DOLAR + val(llave_rep01!ALL_IMPORTE)
       K_COBRA_SOLES_DOLAR = K_COBRA_SOLES_DOLAR + redondea(val(llave_rep01!ALL_IMPORTE) * WS_TIPO_C)
 End If

llave_rep01.MoveNext
Loop

If K_COBRA_SOLES_DOLAR <> 0 Then
 f1 = f1 + 1
 xl.Cells(f1, 6) = "'TOTAL = "
 xl.Cells(f1, 7) = "'S/."
 xl.Cells(f1, 8) = K_COBRA_SOLES
 xl.Cells(f1, 9) = "'US$."
 xl.Cells(f1, 10) = K_COBRA_DOLAR
 'xl.Cells(F1, 11) = "'S/."
 'xl.Cells(F1, 12) = K_COBRA_SOLES_DOLAR
 wranF = "A" & f1 & ":J" & f1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 wranF = "A" & f1 + 1 & ":J" & f1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
End If
    

' EGRESOS  VARIOS DE CAJA
If wusu = "A" Then
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_CODUSU = ? AND ALL_CODTRA = 5355 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
Else
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ?  AND ALL_CODTRA = 5355 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
 PS_REP01(0) = 0
 PS_REP01(1) = 0
Else
 PS_REP01(0) = 0
 PS_REP01(1) = 0
End If
PS_REP01(2) = LK_FECHA_DIA
If wusu = "A" Then
  PS_REP01(3) = 0
End If

Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
 PS_REP01(0) = "01"
 PS_REP01(1) = "02"
Else
 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If

llave_rep01.Requery

f1 = f1 + 2
xl.Cells(f1, 1) = "'EGRESOS VARIOS: "
xl.Cells(f1, 5) = ""
xl.Cells(f1, 6) = ""
K2_COBRA_SOLES = 0
K2_COBRA_DOLAR = 0
K2_COBRA_SOLES_DOLAR = 0
Do Until llave_rep01.EOF
   f1 = f1 + 1
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
       xl.Cells(f1, 1) = Left(llave_rep01!all_concepto, 20)
       xl.Cells(f1, 7) = "'S/."
       xl.Cells(f1, 8) = Format(llave_rep01!ALL_IMPORTE, "0.000")
       K2_COBRA_SOLES = K2_COBRA_SOLES + val(llave_rep01!ALL_IMPORTE)
       K2_COBRA_SOLES_DOLAR = K2_COBRA_SOLES_DOLAR + val(llave_rep01!ALL_IMPORTE)
    Else
       WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
       xl.Cells(f1, 1) = "COMPRA US$."
       xl.Cells(f1, 9) = "'US$."
       xl.Cells(f1, 10) = Format(llave_rep01!ALL_IMPORTE, "0.000")
       K2_COBRA_DOLAR = K2_COBRA_DOLAR + val(llave_rep01!ALL_IMPORTE)
       K2_COBRA_SOLES_DOLAR = K2_COBRA_SOLES_DOLAR + redondea(val(llave_rep01!ALL_IMPORTE) * WS_TIPO_C)
 End If

llave_rep01.MoveNext
Loop

If K2_COBRA_SOLES_DOLAR <> 0 Then
  f1 = f1 + 1
  xl.Cells(f1, 6) = "'TOTAL = "
  xl.Cells(f1, 7) = "'S/."
  xl.Cells(f1, 8) = K2_COBRA_SOLES
  xl.Cells(f1, 9) = "'US$."
  xl.Cells(f1, 10) = K2_COBRA_DOLAR
  'xl.Cells(F1, 11) = "'S/."
  'xl.Cells(F1, 12) = K2_COBRA_SOLES_DOLAR
  wranF = "A" & f1 & ":J" & f1
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  wranF = "A" & f1 + 1 & ":J" & f1 + 1
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
End If

f1 = f1 + 2
xl.Cells(f1, 1) = "'ENTREGA AL BANCO= "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = (VENTA_TOTAL_SOLES + COBRA_EFEC_SOLES) + SOLES_TOTAL_COBRA + K_COBRA_SOLES - K2_COBRA_SOLES
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = (VENTA_TOTAL_DOLAR + COBRA_EFEC_DOLAR) + DOLAR_TOTAL_COBRA + K_COBRA_DOLAR - K2_COBRA_DOLAR
'xl.Cells(F1, 11) = "'S/."
'xl.Cells(F1, 12) = (VENTA_TOTAL_SOLES_DOLAR + COBRA_EFEC_SOLES_DOLAR) + TOTAL_COBRA_SD + K_COBRA_SOLES_DOLAR - K2_COBRA_SOLES_DOLAR
wranF = "A" & f1 & ":J" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "A" & f1 + 1 & ":J" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
    
    
DoEvents
FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
DoEvents
xcuenta = 1
xl.APPLICATION.Visible = True
xl.Cells(2, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
If checia.Value = 1 Then
  xl.Cells(2, 1) = "Almacenes 01 - 02 "
End If
xl.Cells(1, 10) = "IMP.: " & Now
xl.Cells(4, 10) = "FECHA:"
xl.Cells(4, 11) = "'" & Format(wsFECHA1, "dd mmm yyyy")
xl.DisplayAlerts = False
xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
xl.Worksheets("Hoja1").Protect PUB_CLAVE
xl.APPLICATION.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = False
FrmImp2.ProgBar.Visible = False
Set xl = Nothing
Screen.MousePointer = 0
FrmImp2.Pantalla.Enabled = True
FrmImp2.cerrar.Enabled = True
FrmImp2.Pantalla.Caption = "Por &Pantalla"
FrmImp2.lblproceso.Visible = False

Exit Sub

IMP_LINEA:
f1 = f1 + 1
If ww_fbg = "F" Then
 xl.Cells(f1, 1) = "FACTURAS"
ElseIf ww_fbg = "B" Then
 xl.Cells(f1, 1) = "BOLETAS"
ElseIf ww_fbg = "G" Then
 xl.Cells(f1, 1) = "GUIAS"
Else
 xl.Cells(f1, 1) = "PENDIENTES"
End If
xl.Cells(f1, 2) = "Nº "
xl.Cells(f1, 3) = ww_serini
xl.Cells(f1, 4) = ww_numini
xl.Cells(f1, 5) = "'al Nº"
xl.Cells(f1, 6) = ww_numfin
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = Format(WMONTO_SOLES, "0.000")
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = Format(WMONTO_DOLAR, "0.000")
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = Format(WSOLES_DOLAR, "0.000")

WMONTO_SOLES = 0
WMONTO_DOLAR = 0
WSOLES_DOLAR = 0
Return



WEXCEL:
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FrmImp2.lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
  DoEvents
  WPAS = "131296"
  'WPAS = PUB_RUTA_OTRO + "CAJA_DET.xls"
  'DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open PUB_RUTA_OTRO & "\RESU_CAJA.xls", 0, True, 4, WPAS, WPAS

  'xl.Workbooks.Open , "C:\ADMIN\HERTISA\CAJA_DET.XLS", 0, True, 4, PUB_CLAVE, PUB_CLAVE
Return

Exit Sub

'LLENA_VALOR:
'For I = QJ To WDIF
'  If I >= 3 Then
'    xl.Cells(F1 + 7, I) = Format(LOC_VALOR, "0.000")
'  End If
'Next I
'Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
Exit Sub

End Sub

Public Sub RESU_ENTREGA()
Dim FINI As Integer
FrmImp2.Pantalla.Enabled = False
Dim WSUMA_EFECTIVO_S As Currency
Dim WSUMA_EFECTIVO_D As Currency
Dim WSUMA_DEPOSITO_S As Currency
Dim WSUMA_DEPOSITO_D As Currency
Dim WSUMA_CHEQUE_S As Currency
Dim WSUMA_CHEQUE_D As Currency

Dim WIMPORTE   As Currency
Dim COBRA_EFEC_SOLES As Currency
Dim COBRA_EFEC_DOLAR As Currency
Dim COBRA_EFEC_SOLES_DOLAR As Currency


Dim VENTA_TOTAL_SOLES As Currency
Dim VENTA_TOTAL_DOLAR As Currency
Dim VENTA_TOTAL_SOLES_DOLAR As Currency
Dim K_DET_SOLES As Currency
Dim K_DET_DOLAR As Currency
Dim K_DET_SOLES_DOLAR As Currency

Dim wnumfac As Currency
Dim wusu As String * 1
Dim ww_fbg
Dim ww_numser
Dim ww_numfac
Dim ww_numini
Dim ww_serini
Dim ww_numfin
Dim WMONTO_SOLES As Currency
Dim WMONTO_DOLAR As Currency
Dim wsFECHA1
Dim WDOLAR_CRED As Currency
Dim WSOLES_CRED As Currency
Dim TT_TOTAL_SOLES As Currency
Dim TT_TOTAL_DOLAR  As Currency
Dim WSOLES_DOLAR As Currency
Dim TT_DOLAR_SOLES As Currency
Dim WSOLES_DOLAR_CRED As Currency
Dim WSOLES_DOLAR_TOT As Currency
Dim SOLES_TOTAL_COBRA As Currency
Dim DOLAR_TOTAL_COBRA As Currency
Dim TOTAL_COBRA_SD As Currency
Dim WS_TIPO_C As Currency

WSUMA_EFECTIVO_S = 0
WSUMA_EFECTIVO_D = 0
WSUMA_DEPOSITO_S = 0
WSUMA_DEPOSITO_D = 0
WSUMA_CHEQUE_S = 0
WSUMA_CHEQUE_D = 0

If Right(txtFecha.Text, 2) = "__" Then
     wsFECHA1 = Left(txtFecha.Text, 8)
Else
     wsFECHA1 = Trim(txtFecha.Text)
End If
If Not IsDate(wsFECHA1) Then
 FrmImp2.Pantalla.Enabled = True
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If


FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
f1 = 5
FINI = 5
pub_mensaje = "Imprimir según su Usuario...?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbYes Then
  wusu = "A"
Else
  wusu = " "
End If

If wusu = "A" Then
   pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10 AND ALL_CODUSU = ? AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
Else
   pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10   AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
  PS_REP01(0) = 0
  PS_REP01(1) = 0
Else
  PS_REP01(0) = 0
  PS_REP01(1) = 0
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = 0
End If

Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
  PS_REP01(0) = "01"
  PS_REP01(1) = "02"
Else
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If
llave_rep01.Requery
If llave_rep01.EOF Then
  MsgBox "No Existe Movimientos", 48, Pub_Titulo
  GoTo CANCELA
  Exit Sub
End If

f1 = f1 + 1
xl.Cells(f1, 1) = "I.- EFECTIVO:"
xl.Worksheets(1).Rows(f1).RowHeight = 16
wranF = "A" & f1
xl.Range(wranF).Font.Bold = True
xl.Range(wranF).Font.Size = 14
  
  
f1 = f1 + 1
xl.Cells(f1, 2) = "VENTAS"
llave_rep01.MoveFirst
K_DET_SOLES = 0
K_DET_DOLAR = 0
K_DET_SOLES_DOLAR = 0
Do Until llave_rep01.EOF
    If llave_rep01!ALL_SIGNO_CAJA <> 1 Then GoTo PASA_CHE2
    If llave_rep01!all_SECUENCIA = 30 Or llave_rep01!all_SECUENCIA = 40 Then
     GoTo PASA_CHE2
    End If
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
        K_DET_SOLES = K_DET_SOLES + val(llave_rep01!ALL_IMPORTE_AMORT)
        K_DET_SOLES_DOLAR = K_DET_SOLES_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
    Else
        WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
        K_DET_DOLAR = K_DET_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
        K_DET_SOLES_DOLAR = K_DET_SOLES_DOLAR + redondea(val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
    End If
PASA_CHE2:
llave_rep01.MoveNext
Loop
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = K_DET_SOLES
xl.Cells(f1, 9) = "'US$"
xl.Cells(f1, 10) = K_DET_DOLAR

' COMPRA DE MONEDAS
If wusu = "A" Then
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_CODUSU = ? AND ALL_CODTRA = 5345 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
Else
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ?  AND ALL_CODTRA = 5345 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
 PS_REP01(0) = 0
 PS_REP01(1) = 0
Else
 PS_REP01(0) = 0
 PS_REP01(1) = 0
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = 0
End If
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
 PS_REP01(0) = "01"
 PS_REP01(1) = "02"
Else
 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If

llave_rep01.Requery
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Min = 0
If Not llave_rep01.EOF Then FrmImp2.ProgBar.max = llave_rep01.RowCount
'F1 = F1 + 1
DOLAR_TOTAL_COBRA = 0
SOLES_TOTAL_COBRA = 0
TOTAL_COBRA_SD = 0
Do Until llave_rep01.EOF
    f1 = f1 + 1
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
       xl.Cells(f1, 2) = "COMPRA S/."
       xl.Cells(f1, 4) = Format(llave_rep01!ALL_TIPO_CAMBIO, "0.0000")
       xl.Cells(f1, 7) = "'S/."
       xl.Cells(f1, 8) = Format(llave_rep01!ALL_IMPORTE, "0.000")
       xl.Cells(f1, 9) = "'US$"
       xl.Cells(f1, 10) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000") * -1
       SOLES_TOTAL_COBRA = SOLES_TOTAL_COBRA + val(llave_rep01!ALL_IMPORTE)
       DOLAR_TOTAL_COBRA = DOLAR_TOTAL_COBRA - val(llave_rep01!ALL_IMPORTE_AMORT)
       WS_TIPO_C = llave_rep01!ALL_TIPO_CAMBIO
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + val(llave_rep01!ALL_IMPORTE)
       TOTAL_COBRA_SD = TOTAL_COBRA_SD - redondea(val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
    Else
       xl.Cells(f1, 2) = "COMPRA US$"
       xl.Cells(f1, 4) = Format(llave_rep01!ALL_TIPO_CAMBIO, "0.0000")
       xl.Cells(f1, 7) = "'S/."
       xl.Cells(f1, 8) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000") * -1
       xl.Cells(f1, 9) = "'US$"
       xl.Cells(f1, 10) = Format(llave_rep01!ALL_IMPORTE, "0.000")
       SOLES_TOTAL_COBRA = SOLES_TOTAL_COBRA - val(llave_rep01!ALL_IMPORTE_AMORT)
       DOLAR_TOTAL_COBRA = DOLAR_TOTAL_COBRA + val(llave_rep01!ALL_IMPORTE)
       WS_TIPO_C = llave_rep01!ALL_TIPO_CAMBIO
       TOTAL_COBRA_SD = TOTAL_COBRA_SD - val(llave_rep01!ALL_IMPORTE_AMORT)
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + redondea(val(llave_rep01!ALL_IMPORTE) * WS_TIPO_C)
 End If

llave_rep01.MoveNext
Loop
K_DET_SOLES = K_DET_SOLES + SOLES_TOTAL_COBRA
K_DET_DOLAR = K_DET_DOLAR + DOLAR_TOTAL_COBRA
WSUMA_EFECTIVO_S = WSUMA_EFECTIVO_S + K_DET_SOLES
WSUMA_EFECTIVO_D = WSUMA_EFECTIVO_D + K_DET_DOLAR

f1 = f1 + 1
'xl.Cells(F1, 6) = "'TOTAL = "
'xl.Cells(F1, 8) = SOLES_TOTAL_COBRA
'xl.Cells(F1, 10) = DOLAR_TOTAL_COBRA
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = K_DET_SOLES
xl.Cells(f1, 9) = "'US$"
xl.Cells(f1, 10) = K_DET_DOLAR
wranF = "G" & f1 & ":J" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "G" & f1 + 1 & ":J" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1


f1 = f1 + 1
f1 = f1 + 1
xl.Cells(f1, 2) = "COBRANZAS "
' SELECT PARA COBRANZAS DE EFECTIVO
If wusu = "A" Then
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_CODUSU = ? AND ALL_TIPMOV <> 10 AND ALL_SIGNO_CAJA = 1 AND all_flag_ext <> 'E' AND ALL_SIGNO_CAR <> 0 AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
 ' pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10   AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
Else
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ?  AND ALL_TIPMOV <> 10 AND ALL_SIGNO_CAJA = 1 AND all_flag_ext <> 'E' AND ALL_SIGNO_CAR <> 0  AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
 'pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10 AND ALL_CODUSU = ? AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
End If

Set PS_REP02 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
 PS_REP02(0) = 0
 PS_REP02(1) = 0
Else
 PS_REP02(0) = 0
 PS_REP02(1) = 0
End If
PS_REP02(2) = wsFECHA1
If wusu = "A" Then
  PS_REP02(3) = 0
End If
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
 PS_REP02(0) = "01"
 PS_REP02(1) = "02"
Else
 PS_REP02(0) = LK_CODCIA
 PS_REP02(1) = ""
End If
PS_REP02(2) = wsFECHA1
If wusu = "A" Then
  PS_REP02(3) = LK_CODUSU
End If
llave_rep02.Requery
SOLES_TOTAL_COBRA = 0
TOTAL_COBRA_SD = 0
DOLAR_TOTAL_COBRA = 0
xl.Cells(f1, 3) = "NRO."
xl.Cells(f1, 4) = "DOC."
Do Until llave_rep02.EOF
    If llave_rep02!all_SECUENCIA = 30 Or llave_rep02!all_SECUENCIA = 40 Then
      GoTo PASA_COBRA1
    End If
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = llave_rep02!all_CODCIA
    pu_cp = llave_rep02!ALL_CP
    pu_codclie = val(llave_rep02!ALL_CODCLIE)
    LEER_CLI_LLAVE
    xl.Cells(f1, 2) = "'- " & Trim(cli_llave!CLI_NOMBRE)
    xl.Cells(f1, 3) = llave_rep02!all_numser_c
    xl.Cells(f1, 4) = llave_rep02!all_numfac_c
    If llave_rep02!all_codban <> 0 Then
    PUB_CODBAN = llave_rep02!all_codban
    pu_codcia = llave_rep02!all_CODCIA
    LEER_CCM_LLAVE
    xl.Cells(f1, 5) = llave_rep02!all_chenum
    xl.Cells(f1, 6) = Left(ccm_llave!CCM_NOMBRE, 10)
    End If
    
    
    If llave_rep02!ALL_MONEDA_CAJA = "S" Then
       WIMPORTE = llave_rep02!ALL_IMPORTE_AMORT
       If llave_rep02!ALL_MONEDA_CLI = "D" Then
         WIMPORTE = redondea(WIMPORTE * llave_rep02!ALL_TIPO_CAMBIO)
       End If
       xl.Cells(f1, 7) = "'S/."
       xl.Cells(f1, 8) = Format(WIMPORTE, "0.000")
       SOLES_TOTAL_COBRA = SOLES_TOTAL_COBRA + WIMPORTE
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + WIMPORTE
    Else
       WIMPORTE = llave_rep02!ALL_IMPORTE_AMORT
       If llave_rep02!ALL_MONEDA_CLI = "S" Then
         WIMPORTE = redondea(llave_rep02!ALL_IMPORTE_AMORT / llave_rep02!ALL_TIPO_CAMBIO)
       End If
       xl.Cells(f1, 9) = "'US$"
       xl.Cells(f1, 10) = Format(WIMPORTE, "0.000")
       WS_TIPO_C = llave_rep02!ALL_TIPO_CAMBIO
       DOLAR_TOTAL_COBRA = DOLAR_TOTAL_COBRA + WIMPORTE
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + val(Format((WIMPORTE * WS_TIPO_C), "0.000")) 'Val(llave_rep01!ALL_TIPO_CAMBIO))
    End If
PASA_COBRA1:
llave_rep02.MoveNext
Loop
f1 = f1 + 1
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = SOLES_TOTAL_COBRA
xl.Cells(f1, 9) = "'US$"
xl.Cells(f1, 10) = DOLAR_TOTAL_COBRA
wranF = "G" & f1 & ":J" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "G" & f1 + 1 & ":J" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1

WSUMA_EFECTIVO_S = WSUMA_EFECTIVO_S + SOLES_TOTAL_COBRA
WSUMA_EFECTIVO_D = WSUMA_EFECTIVO_D + DOLAR_TOTAL_COBRA

f1 = f1 + 1
xl.Cells(f1, 4) = "'TOTAL EFECTIVO = "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = WSUMA_EFECTIVO_S
xl.Cells(f1, 9) = "'US$"
xl.Cells(f1, 10) = WSUMA_EFECTIVO_D
xl.Worksheets(1).Rows(f1).RowHeight = 16
wranF = "D" & f1 & ":J" & f1
xl.Range(wranF).Font.Bold = True

wranF = "G" & f1 & ":J" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "G" & f1 + 1 & ":J" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1


f1 = f1 + 1
xl.Cells(f1, 1) = "II.- DEPOSITOS:"
xl.Worksheets(1).Rows(f1).RowHeight = 16
wranF = "A" & f1
xl.Range(wranF).Font.Bold = True
xl.Range(wranF).Font.Size = 14

f1 = f1 + 1
xl.Cells(f1, 2) = "VENTAS CONTADOS:"

If wusu = "A" Then
   pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10 AND ALL_CODUSU = ? AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
Else
   pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10   AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
  PS_REP01(0) = 0
  PS_REP01(1) = 0
Else
  PS_REP01(0) = 0
  PS_REP01(1) = 0
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = 0
End If
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
  PS_REP01(0) = "01"
  PS_REP01(1) = "02"
Else
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If
llave_rep01.Requery
K_DET_SOLES = 0
K_DET_DOLAR = 0
K_DET_SOLES_DOLAR = 0
Do Until llave_rep01.EOF
    'If llave_rep01!ALL_SIGNO_CAJA <> 0 Then GoTo PASA_CHE3
    If llave_rep01!all_SECUENCIA <> 40 Then GoTo PASA_CHE3
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
        K_DET_SOLES = K_DET_SOLES + val(llave_rep01!ALL_IMPORTE_AMORT)
        K_DET_SOLES_DOLAR = K_DET_SOLES_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
    Else
        WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
        K_DET_DOLAR = K_DET_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
        K_DET_SOLES_DOLAR = K_DET_SOLES_DOLAR + redondea(val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
    End If
PASA_CHE3:
llave_rep01.MoveNext
Loop
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = K_DET_SOLES
xl.Cells(f1, 9) = "'US$"
xl.Cells(f1, 10) = K_DET_DOLAR
wranF = "G" & f1 & ":J" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "G" & f1 + 1 & ":J" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1

WSUMA_DEPOSITO_S = WSUMA_DEPOSITO_S + K_DET_SOLES
WSUMA_DEPOSITO_D = WSUMA_DEPOSITO_D + K_DET_DOLAR

f1 = f1 + 1
xl.Cells(f1, 2) = "COBRANZAS:"
SOLES_TOTAL_COBRA = 0
TOTAL_COBRA_SD = 0
DOLAR_TOTAL_COBRA = 0
llave_rep02.MoveFirst
Do Until llave_rep02.EOF
    If llave_rep02!all_SECUENCIA <> 40 Then GoTo PASA_COBRA2
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = llave_rep02!all_CODCIA
    pu_cp = llave_rep02!ALL_CP
    pu_codclie = val(llave_rep02!ALL_CODCLIE)
    LEER_CLI_LLAVE
    xl.Cells(f1, 2) = "'- " & Trim(cli_llave!CLI_NOMBRE)
    xl.Cells(f1, 3) = llave_rep02!all_numser_c
    xl.Cells(f1, 4) = llave_rep02!all_numfac_c
    If llave_rep02!all_codban <> 0 Then
    PUB_CODBAN = llave_rep02!all_codban
    pu_codcia = llave_rep02!all_CODCIA
    LEER_CCM_LLAVE
    xl.Cells(f1, 5) = llave_rep02!all_chenum
    xl.Cells(f1, 6) = Left(ccm_llave!CCM_NOMBRE, 10)
    End If
    If llave_rep02!ALL_MONEDA_CAJA = "S" Then
       WIMPORTE = llave_rep02!ALL_IMPORTE_AMORT
       If llave_rep02!ALL_MONEDA_CLI = "D" Then
         WIMPORTE = redondea(WIMPORTE * llave_rep02!ALL_TIPO_CAMBIO)
       End If
       xl.Cells(f1, 7) = "'S/."
       xl.Cells(f1, 8) = Format(WIMPORTE, "0.000")
       SOLES_TOTAL_COBRA = SOLES_TOTAL_COBRA + WIMPORTE
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + WIMPORTE
    Else
       WIMPORTE = llave_rep02!ALL_IMPORTE_AMORT
       If llave_rep02!ALL_MONEDA_CLI = "S" Then
         WIMPORTE = redondea(llave_rep02!ALL_IMPORTE_AMORT / llave_rep02!ALL_TIPO_CAMBIO)
       End If
       xl.Cells(f1, 9) = "'US$"
       xl.Cells(f1, 10) = Format(WIMPORTE, "0.000")
       WS_TIPO_C = llave_rep02!ALL_TIPO_CAMBIO
       DOLAR_TOTAL_COBRA = DOLAR_TOTAL_COBRA + WIMPORTE
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + val(Format((WIMPORTE * WS_TIPO_C), "0.000")) 'Val(llave_rep01!ALL_TIPO_CAMBIO))
    End If
    
    'If llave_rep02!ALL_MONEDA_CAJA = "S" Then
    '    xl.Cells(F1, 8) = Format(llave_rep02!ALL_IMPORTE_AMORT, "0.000")
    '   SOLES_TOTAL_COBRA = SOLES_TOTAL_COBRA + llave_rep02!ALL_IMPORTE_AMORT
    '   TOTAL_COBRA_SD = TOTAL_COBRA_SD + llave_rep02!ALL_IMPORTE_AMORT
    'Else
    '   WIMPORTE = llave_rep02!ALL_IMPORTE_AMORT
    '   If llave_rep02!ALL_MONEDA_CLI = "S" Then
    '     WIMPORTE = redondea(llave_rep02!ALL_IMPORTE_AMORT / llave_rep02!ALL_TIPO_CAMBIO)
    '   End If
    '   xl.Cells(F1, 10) = Format(WIMPORTE, "0.000")
    '   WS_TIPO_C = llave_rep02!ALL_TIPO_CAMBIO
    '   DOLAR_TOTAL_COBRA = DOLAR_TOTAL_COBRA + WIMPORTE
    '   TOTAL_COBRA_SD = TOTAL_COBRA_SD + Val(Format((WIMPORTE * WS_TIPO_C), "0.000")) 'Val(llave_rep01!ALL_TIPO_CAMBIO))
    '  End If
PASA_COBRA2:
llave_rep02.MoveNext
Loop
f1 = f1 + 1
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = SOLES_TOTAL_COBRA
xl.Cells(f1, 9) = "'US$"
xl.Cells(f1, 10) = DOLAR_TOTAL_COBRA
wranF = "G" & f1 & ":J" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "G" & f1 + 1 & ":J" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1

WSUMA_DEPOSITO_S = WSUMA_DEPOSITO_S + SOLES_TOTAL_COBRA
WSUMA_DEPOSITO_D = WSUMA_DEPOSITO_D + DOLAR_TOTAL_COBRA
f1 = f1 + 1
xl.Cells(f1, 4) = "'TOTAL DEPOSITOS = "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = WSUMA_DEPOSITO_S
xl.Cells(f1, 9) = "'US$"
xl.Cells(f1, 10) = WSUMA_DEPOSITO_D
xl.Worksheets(1).Rows(f1).RowHeight = 16
wranF = "D" & f1 & ":J" & f1
xl.Range(wranF).Font.Bold = True



wranF = "G" & f1 & ":J" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "G" & f1 + 1 & ":J" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  
  
  
' CHEQUES

f1 = f1 + 1
xl.Cells(f1, 1) = "III.- CHEQUES:"
xl.Worksheets(1).Rows(f1).RowHeight = 16
wranF = "A" & f1
xl.Range(wranF).Font.Bold = True
xl.Range(wranF).Font.Size = 14

f1 = f1 + 1
xl.Cells(f1, 2) = "VENTAS :"

llave_rep01.MoveFirst
WSUMA_CHEQUE_D = 0
WSUMA_CHEQUE_S = 0
K_DET_SOLES = 0
K_DET_DOLAR = 0
K_DET_SOLES_DOLAR = 0
Do Until llave_rep01.EOF
    If llave_rep01!ALL_SIGNO_CAJA <> 1 Then GoTo PASA_CHE4
    If llave_rep01!all_SECUENCIA = 30 Or llave_rep01!all_SECUENCIA = 50 Then
    Else
     GoTo PASA_CHE4
    End If
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
        K_DET_SOLES = K_DET_SOLES + val(llave_rep01!ALL_IMPORTE_AMORT)
        K_DET_SOLES_DOLAR = K_DET_SOLES_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
    Else
        WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
        K_DET_DOLAR = K_DET_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
        K_DET_SOLES_DOLAR = K_DET_SOLES_DOLAR + redondea(val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
    End If
PASA_CHE4:
llave_rep01.MoveNext
Loop
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = K_DET_SOLES
xl.Cells(f1, 9) = "'US$"
xl.Cells(f1, 10) = K_DET_DOLAR
wranF = "G" & f1 & ":J" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "G" & f1 + 1 & ":J" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1

WSUMA_CHEQUE_S = WSUMA_CHEQUE_S + K_DET_SOLES
WSUMA_CHEQUE_D = WSUMA_CHEQUE_D + K_DET_DOLAR
 
f1 = f1 + 1
xl.Cells(f1, 2) = "COBRANZAS:"
SOLES_TOTAL_COBRA = 0
TOTAL_COBRA_SD = 0
DOLAR_TOTAL_COBRA = 0
llave_rep02.MoveFirst
Do Until llave_rep02.EOF
    If llave_rep02!all_SECUENCIA <> 30 Then GoTo PASA_COBRA4
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = llave_rep02!all_CODCIA
    pu_cp = llave_rep02!ALL_CP
    pu_codclie = val(llave_rep02!ALL_CODCLIE)
    LEER_CLI_LLAVE
    xl.Cells(f1, 2) = "'- " & Trim(cli_llave!CLI_NOMBRE)
    xl.Cells(f1, 3) = llave_rep02!all_numser_c
    xl.Cells(f1, 4) = llave_rep02!all_numfac_c
    If llave_rep02!all_codban <> 0 Then
    PUB_CODBAN = llave_rep02!all_codban
    pu_codcia = llave_rep02!all_CODCIA
    LEER_CCM_LLAVE
    xl.Cells(f1, 5) = llave_rep02!all_chenum
    xl.Cells(f1, 6) = Left(ccm_llave!CCM_NOMBRE, 10)
    End If
    If llave_rep02!ALL_MONEDA_CAJA = "S" Then
       WIMPORTE = llave_rep02!ALL_IMPORTE_AMORT
       If llave_rep02!ALL_MONEDA_CLI = "D" Then
         WIMPORTE = redondea(WIMPORTE * llave_rep02!ALL_TIPO_CAMBIO)
       End If
       xl.Cells(f1, 7) = "'S/."
       xl.Cells(f1, 8) = Format(WIMPORTE, "0.000")
       SOLES_TOTAL_COBRA = SOLES_TOTAL_COBRA + WIMPORTE
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + WIMPORTE
    Else
       WIMPORTE = llave_rep02!ALL_IMPORTE_AMORT
       If llave_rep02!ALL_MONEDA_CLI = "S" Then
         WIMPORTE = redondea(llave_rep02!ALL_IMPORTE_AMORT / llave_rep02!ALL_TIPO_CAMBIO)
       End If
       xl.Cells(f1, 9) = "'US$"
       xl.Cells(f1, 10) = Format(WIMPORTE, "0.000")
       WS_TIPO_C = llave_rep02!ALL_TIPO_CAMBIO
       DOLAR_TOTAL_COBRA = DOLAR_TOTAL_COBRA + WIMPORTE
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + val(Format((WIMPORTE * WS_TIPO_C), "0.000")) 'Val(llave_rep01!ALL_TIPO_CAMBIO))
    End If
    
PASA_COBRA4:
llave_rep02.MoveNext
Loop
f1 = f1 + 1
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = SOLES_TOTAL_COBRA
xl.Cells(f1, 9) = "'US$"
xl.Cells(f1, 10) = DOLAR_TOTAL_COBRA
wranF = "G" & f1 & ":J" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "G" & f1 + 1 & ":J" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1

WSUMA_CHEQUE_S = WSUMA_CHEQUE_S + SOLES_TOTAL_COBRA
WSUMA_CHEQUE_D = WSUMA_CHEQUE_D + DOLAR_TOTAL_COBRA

f1 = f1 + 1
xl.Cells(f1, 4) = "'TOTAL CHEQUE = "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = WSUMA_CHEQUE_S
xl.Cells(f1, 9) = "'US$"
xl.Cells(f1, 10) = WSUMA_CHEQUE_D
xl.Worksheets(1).Rows(f1).RowHeight = 16
wranF = "D" & f1 & ":J" & f1
xl.Range(wranF).Font.Bold = True

wranF = "G" & f1 & ":J" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "G" & f1 + 1 & ":J" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1

f1 = f1 + 1
xl.Cells(f1, 4) = "'TOTAL GENERAL = "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = WSUMA_EFECTIVO_S + WSUMA_DEPOSITO_S + WSUMA_CHEQUE_S
xl.Cells(f1, 9) = "'US$"
xl.Cells(f1, 10) = WSUMA_EFECTIVO_D + WSUMA_DEPOSITO_D + WSUMA_CHEQUE_D
wranF = "A" & f1 & ":J" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "A" & f1 + 1 & ":J" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  
'xl.Application.Visible = True
'XlBorderType: xlInsideHorizontal, xlInsideVertical, xlDiagonalDown, xlDiagonalUp, xlEdgeBottom, xlEdgeLeft, xlEdgeRight o xlEdgeTop.
wranF = "K" & 6 & ":N" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlInsideHorizontal).LineStyle = 1
wranF = "J" & 6 & ":O" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlInsideVertical).LineStyle = 1

wranF = "H" & 6 & ":I" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlInsideVertical).LineStyle = 1
wranF = "F" & 6 & ":G" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlInsideVertical).LineStyle = 1

    
DoEvents
FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
DoEvents
xcuenta = 1
xl.APPLICATION.Visible = True
xl.Cells(2, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
If checia.Value = 1 Then
  xl.Cells(2, 1) = "Almacenes 01 - 02 "
End If
xl.Cells(1, 1) = "IMP.: " & Now
xl.Cells(3, 1) = "REPORTE PARA BANCOS DEL : " & Format(wsFECHA1, "dd mmm yyyy")
xl.DisplayAlerts = False
xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
xl.Worksheets("Hoja1").Protect PUB_CLAVE
xl.APPLICATION.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = False
FrmImp2.ProgBar.Visible = False
Set xl = Nothing
Screen.MousePointer = 0
FrmImp2.Pantalla.Enabled = True
FrmImp2.Pantalla.Caption = "Por &Pantalla"
FrmImp2.lblproceso.Visible = False

Exit Sub

IMP_LINEA:
'F1 = F1 + 1
'If ww_fbg = "F" Then
' xl.Cells(F1, 1) = "FACTURAS"
'ElseIf ww_fbg = "B" Then
' xl.Cells(F1, 1) = "BOLETAS"
'ElseIf ww_fbg = "G" Then
' xl.Cells(F1, 1) = "GUIAS"
'Else
' xl.Cells(F1, 1) = "PENDIENTES"
'End If
'xl.Cells(F1, 2) = "Nº "
'xl.Cells(F1, 3) = ww_serini
'xl.Cells(F1, 4) = ww_numini
'xl.Cells(F1, 5) = "'al Nº"
'xl.Cells(F1, 6) = ww_numfin
'xl.Cells(F1, 7) = "'S/."
'xl.Cells(F1, 8) = Format(WMONTO_SOLES, "0.000")
'xl.Cells(F1, 9) = "'US$"
'xl.Cells(F1, 10) = Format(WMONTO_DOLAR, "0.000")
'xl.Cells(F1, 11) = "'S/."
'xl.Cells(F1, 12) = Format(WSOLES_DOLAR, "0.000")

WMONTO_SOLES = 0
WMONTO_DOLAR = 0
WSOLES_DOLAR = 0
Return



WEXCEL:
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FrmImp2.lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
  DoEvents
  WPAS = "131296"
  'WPAS = PUB_RUTA_OTRO + "CAJA_DET.xls"
  'DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open PUB_RUTA_OTRO & "\RESU_BANCO.xls", 0, True, 4, WPAS, WPAS

  'xl.Workbooks.Open , "C:\ADMIN\HERTISA\CAJA_DET.XLS", 0, True, 4, PUB_CLAVE, PUB_CLAVE
Return

Exit Sub

'LLENA_VALOR:
'For I = QJ To WDIF
'  If I >= 3 Then
'    xl.Cells(F1 + 7, I) = Format(LOC_VALOR, "0.000")
'  End If
'Next I
'Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
Exit Sub

End Sub


Private Sub Transferencia()
Dim wd_banco As String * 20
Dim wd_chenum As String * 15
Dim wd_codigo  As String * 15
Dim RR
Dim PSFAR As rdoQuery
Dim far_r  As rdoResultset
Dim CONTADOR As Integer
Dim ArchOrigen, ArchDestino As String
Dim pub_mensaje, estilo, respuesta
Dim PSfar_menorx As rdoQuery
Dim far_menorx As rdoResultset
Dim all_banco As rdoResultset
Dim PSALL_BANCO As rdoQuery

Dim WLUGAR As String
Dim WZONA As String
Dim WSUBZONA As String

'On Error GoTo SALE

pub_cadena = "SELECT ALL_CODBAN,ALL_CHENUM FROM ALLOG WHERE ALL_CODCIA= ? AND ALL_FECHA_DIA = ? AND ALL_FBG = ? AND ALL_NUMSER = ? AND ALL_NUMFAC = ? AND ALL_NUMOPER = ? AND ALL_CODTRA = 2401 AND ALL_FLAG_EXT <> 'E' "
Set PSALL_BANCO = CN.CreateQuery("", pub_cadena)
PSALL_BANCO(0) = 0
PSALL_BANCO(1) = LK_FECHA_DIA
PSALL_BANCO(2) = 0
PSALL_BANCO(3) = 0
PSALL_BANCO(4) = 0
PSALL_BANCO(5) = 0
Set all_banco = PSALL_BANCO.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM facart WHERE FAR_CODCIA= ? AND FAR_FECHA>= ? AND FAR_FECHA<= ? AND FAR_TIPMOV = 10  AND FAR_ESTADO<>'E'  ORDER BY FAR_FECHA, FAR_NUMOPER"
Set PSfar_menorx = CN.CreateQuery("", pub_cadena)
Set far_menorx = PSfar_menorx.OpenResultset(rdOpenKeyset, rdConcurValues)

PSfar_menorx.rdoParameters(0) = LK_CODCIA
PSfar_menorx.rdoParameters(1) = txtCampo1.Text
PSfar_menorx.rdoParameters(2) = txtcampo2.Text
far_menorx.Requery

Screen.MousePointer = 11

DoEvents
transfer.DatabaseName = "C:\Admin\Standar\TRANSFER.mdb"
transfer.RecordSource = "SELECT * FROM FACART"
On Error GoTo NOHAY
transfer.Refresh
On Error GoTo 0
Do Until transfer.Recordset.EOF
   transfer.Recordset.Edit
   transfer.Recordset.Delete
   transfer.Recordset.MoveNext
Loop

SALTA:
FrmImp2.lblproceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblproceso.Caption = "Transferencia de Ventas"
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.max = far_menorx.RowCount
CONTADOR = 0
If Not far_menorx.EOF Then
   PUB_FECHA = far_menorx!FAR_fecha
End If

CN.Execute "Begin Transaction", rdExecDirect
      
      Do Until far_menorx.EOF
         wd_chenum = " "
         wd_codigo = " "
         wd_banco = " "
         PSALL_BANCO(0) = far_menorx!FAR_CODCIA
         PSALL_BANCO(1) = far_menorx!FAR_fecha
         PSALL_BANCO(2) = far_menorx!far_fbg
         PSALL_BANCO(3) = far_menorx!far_numser
         PSALL_BANCO(4) = far_menorx!far_numfac
         PSALL_BANCO(5) = far_menorx!FAR_NUMOPER
         all_banco.Requery
         If Not all_banco.EOF Then
         wd_chenum = all_banco!all_chenum
         PUB_CODBAN = all_banco!all_codban
           If PUB_CODBAN <> 0 Then
              pu_codcia = LK_CODCIA
              SQ_OPER = 1
              LEER_CCM_LLAVE
              If Not ccm_llave.EOF Then
               wd_banco = Trim(ccm_llave!CCM_NOMBRE)
               wd_codigo = Trim(Nulo_Valors(ccm_llave!CCM_ALTERNO))
              End If
           End If
         End If
         transfer.Recordset.AddNew
         transfer.Recordset!FAR_CODCIA = far_menorx!FAR_CODCIA
         transfer.Recordset!far_fbg = far_menorx!far_fbg
         transfer.Recordset!far_numser = far_menorx!far_numser
         transfer.Recordset!far_numfac = far_menorx!far_numfac
         transfer.Recordset!FAR_NUMSEC = far_menorx!FAR_NUMSEC
         transfer.Recordset!FAR_fecha = far_menorx!FAR_fecha
         transfer.Recordset!far_codclie = 0
         transfer.Recordset!far_codart = 0
         transfer.Recordset!FAR_PRECIO = far_menorx!FAR_PRECIO
         transfer.Recordset!FAR_pordescto1 = far_menorx!FAR_pordescto1
         transfer.Recordset!far_IMPTO = far_menorx!far_IMPTO
         transfer.Recordset!FAR_SUBTOTAL = far_menorx!FAR_SUBTOTAL
         transfer.Recordset!FAR_CODVEN = far_menorx!FAR_CODVEN
         transfer.Recordset!far_cantidad = far_menorx!far_cantidad
         transfer.Recordset!far_signo_car = far_menorx!far_signo_car
         transfer.Recordset!FAR_MONEDA = far_menorx!FAR_MONEDA
         transfer.Recordset!far_serguia = far_menorx!far_serguia
         transfer.Recordset!far_numguia = far_menorx!far_numguia
         transfer.Recordset!FAR_DIAS = far_menorx!FAR_DIAS
         transfer.Recordset!far_descri = far_menorx!far_descri
         transfer.Recordset!far_PESO = far_menorx!far_PESO
         transfer.Recordset!far_signo_arm = far_menorx!far_signo_arm
         transfer.Recordset!FAR_GASTOS = far_menorx!FAR_GASTOS
         transfer.Recordset!FAR_NUMSER_C = far_menorx!FAR_NUMSER_C
         transfer.Recordset!FAR_NUMFAC_C = far_menorx!FAR_NUMFAC_C
         transfer.Recordset!FAR_cp = far_menorx!FAR_cp
         transfer.Recordset!far_cod_sunat = far_menorx!far_cod_sunat
         transfer.Recordset!FAR_TIPDOC = far_menorx!FAR_TIPDOC
         transfer.Recordset!FAR_BRUTO = far_menorx!FAR_BRUTO
         
         transfer.Recordset!FAR_CHENUM = wd_chenum
         transfer.Recordset!FAR_CODBAN = wd_codigo
         transfer.Recordset!FAR_BANCO = wd_banco
         transfer.Recordset!FAR_SECUENCIA = far_menorx!FAR_NUM_LOTE
         SQ_OPER = 1
         pu_codclie = far_menorx!far_codclie
         pu_cp = "C"
         pu_codcia = LK_CODCIA
         LEER_CLI_LLAVE
         PUB_KEY = far_menorx!far_codart
         LEER_ART_LLAVE
         transfer.Recordset!far_nomart = Trim(art_LLAVE!art_nombre)
         transfer.Recordset!FAR_NOMCLI = cli_llave!CLI_NOMBRE
         transfer.Recordset!far_ALTERNO = art_LLAVE!ART_ALTERNO
         If Trim(cli_llave!cli_ruc_esposo) <> "" Then
            transfer.Recordset!FAR_RUC = Trim(cli_llave!cli_ruc_esposo)
         Else
            transfer.Recordset!FAR_RUC = " "
         End If
         If Trim(cli_llave!cli_RUC_ESPOSA) <> "" Then
            transfer.Recordset!FAR_DNI = Trim(cli_llave!cli_RUC_ESPOSA)
         Else
            transfer.Recordset!FAR_DNI = " "
         End If
         
         
        SQ_OPER = 1
        PUB_CODCIA = "00"
        PUB_NUMTAB = cli_llave!CLI_LUGAR_TRAB
        PUB_TIPREG = 25
        LEER_TAB_LLAVE
        WLUGAR = ""
        If Not tab_llave.EOF Then
        WLUGAR = Trim(tab_llave!tab_NOMLARGO)
        End If
        
        PUB_NUMTAB = cli_llave!cli_TRAB_ZONA
        PUB_TIPREG = 20
        LEER_TAB_LLAVE
        WZONA = ""
        If Not tab_llave.EOF Then
        WZONA = Trim(tab_llave!tab_NOMLARGO)
        End If
        PUB_NUMTAB = cli_llave!cli_TRAB_SUBZONA
        PUB_TIPREG = 35
        LEER_TAB_LLAVE
        WSUBZONA = ""
        If Not tab_llave.EOF Then
        WSUBZONA = Trim(tab_llave!tab_NOMLARGO)
        End If
        If val(cli_llave!CLI_TRAB_NUM) <> 0 Then
           transfer.Recordset!FAR_DIREC = Trim(WLUGAR) + " " + Trim(cli_llave!CLI_TRAB_DIREC) + " # " + Trim(cli_llave!CLI_TRAB_NUM) & "  " & WZONA & "  " & WSUBZONA
        Else
           transfer.Recordset!FAR_DIREC = Trim(WLUGAR) + " " + Trim(cli_llave!CLI_TRAB_DIREC) & "  " & WZONA & "  " & WSUBZONA
        End If
         transfer.Recordset!FAR_DIREC = Left(Trim(transfer.Recordset!FAR_DIREC), 150)
         If Trim(transfer.Recordset!FAR_DIREC) = "" Then transfer.Recordset!FAR_DIREC = " "
         PUB_FECHA_VCTO = far_menorx!FAR_fecha
         transfer.Recordset.Update
         far_menorx.MoveNext
         CONTADOR = CONTADOR + 1
         FrmImp2.ProgBar.Value = CONTADOR
         DoEvents
  Loop
  
CN.Execute "Commit Transaction", rdExecDirect

transfer.RecordSource = "SELECT * FROM FECHAS"
transfer.Refresh
transfer.Recordset.Edit
transfer.Recordset!fecha_final = PUB_FECHA_VCTO
transfer.Recordset.Update
On Error GoTo EMES
RR = Shell("C:\WSADMIN\WST.BAT", 1)

MsgBox "PROCESO DE VENTAS TRANSFERIDO " & Chr(13) & "DEL :" & PUB_FECHA & "  AL " & PUB_FECHA_VCTO & Chr(13) & "Ubicación de Archivo :  C:\WSADMIN\WSTraf.ZIP"
Screen.MousePointer = 0
Unload FrmImp2

Exit Sub
SALE:
Screen.MousePointer = 0
MsgBox Err.Description, 48, Pub_Titulo
CN.Execute "Rollback Transaction", rdExecDirect
 Screen.MousePointer = 0
 Unload FrmImp2

Exit Sub
EMES:
  MsgBox "Proceso Terminado .. Pero EL proceso de Empaquetar hacerlo Manual", vbCritical, Pub_Titulo
  Screen.MousePointer = 0
  Unload FrmImp2
Exit Sub
NOHAY:
   MsgBox "No se encontro el Transfer.mdb .. ", vbCritical, Pub_Titulo
   Screen.MousePointer = 0
   Unload FrmImp2
End Sub

Public Sub RESU_CAJA_VENUS()
Dim var_soles As Currency
Dim var_dolar As Currency
Dim var_soles_dolar  As Currency
Dim TOTALS1 As Currency
Dim TOTALD1 As Currency
Dim TOTALSD1 As Currency

Dim K_COBRA_SOLES As Currency
Dim K_COBRA_DOLAR As Currency
Dim K_COBRA_SOLES_DOLAR As Currency
Dim K2_COBRA_SOLES As Currency
Dim K2_COBRA_DOLAR As Currency
Dim K2_COBRA_SOLES_DOLAR As Currency

FrmImp2.Pantalla.Enabled = False
FrmImp2.cerrar.Enabled = False
Dim WIMPORTE As Currency
Dim COBRA_EFEC_SOLES As Currency
Dim COBRA_EFEC_DOLAR As Currency
Dim COBRA_EFEC_SOLES_DOLAR As Currency


Dim VENTA_TOTAL_SOLES As Currency
Dim VENTA_TOTAL_DOLAR As Currency
Dim VENTA_TOTAL_SOLES_DOLAR As Currency
Dim K_DET_SOLES As Currency
Dim K_DET_DOLAR As Currency
Dim K_DET_SOLES_DOLAR As Currency

Dim wnumfac As Currency
Dim wusu As String * 1
Dim ww_fbg
Dim ww_numser
Dim ww_numfac
Dim ww_numini
Dim ww_serini
Dim ww_numfin
Dim WMONTO_SOLES As Currency
Dim WMONTO_DOLAR As Currency
Dim wsFECHA1
Dim WDOLAR_CRED As Currency
Dim WSOLES_CRED As Currency
Dim TT_TOTAL_SOLES As Currency
Dim TT_TOTAL_DOLAR  As Currency
Dim WSOLES_DOLAR As Currency
Dim TT_DOLAR_SOLES As Currency
Dim WSOLES_DOLAR_CRED As Currency
Dim WSOLES_DOLAR_TOT As Currency
Dim SOLES_TOTAL_COBRA As Currency
Dim DOLAR_TOTAL_COBRA As Currency
Dim TOTAL_COBRA_SD As Currency
Dim WS_TIPO_C As Currency

If Right(txtFecha.Text, 2) = "__" Then
     wsFECHA1 = Left(txtFecha.Text, 8)
Else
     wsFECHA1 = Trim(txtFecha.Text)
End If
If Not IsDate(wsFECHA1) Then
 FrmImp2.Pantalla.Enabled = True
 FrmImp2.cerrar.Enabled = True
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If


FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
f1 = 5
pub_mensaje = "Imprimir según su Usuario...?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbYes Then
  wusu = "A"
Else
  wusu = " "
End If

If wusu = "A" Then
   pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10 AND ALL_CODUSU = ? AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
Else
   pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10   AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
  PS_REP01(0) = 0
  PS_REP01(1) = 0
Else
  PS_REP01(0) = 0
  PS_REP01(1) = 0
End If
PS_REP01(2) = LK_FECHA_DIA
If wusu = "A" Then
  PS_REP01(3) = 0
End If
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
  PS_REP01(0) = "01"
  PS_REP01(1) = "02"
Else
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If


llave_rep01.Requery

If Not llave_rep01.EOF Then
FrmImp2.ProgBar.Min = 0
If llave_rep01.RowCount <> 0 Then FrmImp2.ProgBar.max = llave_rep01.RowCount
f1 = 6
ww_fbg = llave_rep01!ALL_FBG
ww_numser = llave_rep01!ALL_NUMSER
ww_serini = llave_rep01!ALL_NUMSER
ww_numfac = llave_rep01!all_numfac
ww_numini = llave_rep01!all_numfac
ww_numfin = llave_rep01!all_numfac
wnumfac = llave_rep01!all_numfac
End If
f1 = f1 + 1
xl.Cells(f1, 1) = "'INGRESOS: "
f1 = f1 + 1
Do Until llave_rep01.EOF
    FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
    If ww_fbg <> llave_rep01!ALL_FBG Then
        GoSub IMP_LINEA
        ww_fbg = llave_rep01!ALL_FBG
        ww_numser = llave_rep01!ALL_NUMSER
        ww_serini = llave_rep01!ALL_NUMSER
        ww_numfac = llave_rep01!all_numfac
        ww_numini = llave_rep01!all_numfac
        wnumfac = llave_rep01!all_numfac
    End If
    If ww_numser <> llave_rep01!ALL_NUMSER Then
        GoSub IMP_LINEA
        ww_numser = llave_rep01!ALL_NUMSER
        ww_serini = llave_rep01!ALL_NUMSER
        ww_numfac = llave_rep01!all_numfac
        ww_numini = llave_rep01!all_numfac
        wnumfac = llave_rep01!all_numfac
    End If
    If wnumfac <> llave_rep01!all_numfac Then
      
   '   MsgBox "Falta el Nro. :" & wnumfac
   '   wnumfac = llave_rep01!all_numfac
    End If
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
        WMONTO_SOLES = WMONTO_SOLES + val(llave_rep01!ALL_IMPORTE_AMORT)
        WSOLES_DOLAR = WSOLES_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
        If llave_rep01!ALL_SIGNO_CAJA = 0 Then
          WSOLES_CRED = WSOLES_CRED + val(llave_rep01!ALL_IMPORTE_AMORT)
          WSOLES_DOLAR_CRED = WSOLES_DOLAR_CRED + val(llave_rep01!ALL_IMPORTE_AMORT)
        End If
    Else
        WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
        WMONTO_DOLAR = redondea(WMONTO_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT))
        WSOLES_DOLAR = WSOLES_DOLAR + redondea(val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
        If llave_rep01!ALL_SIGNO_CAJA = 0 Then
          WDOLAR_CRED = WDOLAR_CRED + val(llave_rep01!ALL_IMPORTE_AMORT)
          WSOLES_DOLAR_CRED = WSOLES_DOLAR_CRED + redondea(val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
        End If
        
    End If
    ww_numfin = llave_rep01!all_numfac
    wnumfac = wnumfac + 1
    llave_rep01.MoveNext
Loop
 GoSub IMP_LINEA
 TT_TOTAL_SOLES = 0
 TT_TOTAL_DOLAR = 0
 TT_DOLAR_SOLES = 0
 
 f1 = f1 + 1
 xl.Cells(f1 + 1, 4) = "TOTAL VENTA = "
 xl.Cells(f1 + 1, 7) = "S/."
 wran1 = "H" & 6
 wran2 = "H" & f1
 wranF = "H" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_SOLES = TT_TOTAL_SOLES + val(xl.Range(wranF))
 xl.Cells(f1 + 1, 9) = "US$."
 wran1 = "J" & 6
 wran2 = "J" & f1
 wranF = "J" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_DOLAR = TT_TOTAL_DOLAR + val(xl.Range(wranF))
 xl.Cells(f1 + 1, 9) = "S/."
 wran1 = "L" & 6
 wran2 = "L" & f1
 wranF = "L" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_DOLAR_SOLES = TT_DOLAR_SOLES + val(xl.Range(wranF))
 
 wranF = "D" & f1 + 1 & ":L" & f1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 wranF = "D" & f1 + 2 & ":L" & f1 + 2
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 



' TOTAL COBRANZAS
SOLES_TOTAL_COBRA = 0
DOLAR_TOTAL_COBRA = 0
TOTAL_COBRA_SD = 0
 
 
If wusu = "A" Then
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_CODUSU = ? AND ALL_TIPMOV <> 10 AND ALL_SIGNO_CAJA = 1 AND all_flag_ext <> 'E' AND ALL_SIGNO_CAR <> 0 AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
Else
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ?  AND ALL_TIPMOV <> 10 AND ALL_SIGNO_CAJA = 1 AND all_flag_ext <> 'E' AND ALL_SIGNO_CAR <> 0  AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
 PS_REP01(0) = 0
 PS_REP01(1) = 0
Else
 PS_REP01(0) = 0
 PS_REP01(1) = 0
End If
PS_REP01(2) = LK_FECHA_DIA
If wusu = "A" Then
  PS_REP01(3) = 0
End If
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
 PS_REP01(0) = "01"
 PS_REP01(1) = "02"
Else
 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If

llave_rep01.Requery
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Min = 0
If Not llave_rep01.EOF Then FrmImp2.ProgBar.max = llave_rep01.RowCount
Do Until llave_rep01.EOF
    FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
       WIMPORTE = llave_rep01!ALL_IMPORTE_AMORT
       If llave_rep01!ALL_MONEDA_CLI = "D" Then
         WIMPORTE = redondea(WIMPORTE * llave_rep01!ALL_TIPO_CAMBIO)
       End If
       SOLES_TOTAL_COBRA = SOLES_TOTAL_COBRA + WIMPORTE
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + WIMPORTE
    Else
       WIMPORTE = llave_rep01!ALL_IMPORTE_AMORT
       If llave_rep01!ALL_MONEDA_CLI = "S" Then
         WIMPORTE = redondea(llave_rep01!ALL_IMPORTE_AMORT / llave_rep01!ALL_TIPO_CAMBIO)
       End If
       WS_TIPO_C = llave_rep01!ALL_TIPO_CAMBIO
       DOLAR_TOTAL_COBRA = DOLAR_TOTAL_COBRA + WIMPORTE
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + val(Format((WIMPORTE * WS_TIPO_C), "0.000")) 'Val(llave_rep01!ALL_TIPO_CAMBIO))
    End If
PASA_COBRA1:
llave_rep01.MoveNext
Loop
'If TOTAL_COBRA_SD <> 0 Then
'xl.Application.Visible = True
  f1 = f1 + 3
  xl.Cells(f1, 4) = "'TOTAL DE COBRANZAS= "
  xl.Cells(f1, 7) = "'S/."
  xl.Cells(f1, 8) = SOLES_TOTAL_COBRA
  xl.Cells(f1, 9) = "'US$."
  xl.Cells(f1, 10) = DOLAR_TOTAL_COBRA
  xl.Cells(f1, 11) = "'S/."
  xl.Cells(f1, 12) = TOTAL_COBRA_SD
  wranF = "D" & f1 & ":L" & f1
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  wranF = "D" & f1 + 1 & ":L" & f1 + 1
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
'End If
  
' INGRESO VARIOS DE CAJA
If wusu = "A" Then
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_CODUSU = ? AND ALL_CODTRA = 5350 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
Else
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ?  AND ALL_CODTRA = 5350 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
 PS_REP01(0) = 0
 PS_REP01(1) = 0
Else
 PS_REP01(0) = 0
 PS_REP01(1) = 0
End If
PS_REP01(2) = LK_FECHA_DIA
If wusu = "A" Then
  PS_REP01(3) = 0
End If
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
 PS_REP01(0) = "01"
 PS_REP01(1) = "02"
Else
 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If

llave_rep01.Requery
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Min = 0
If Not llave_rep01.EOF Then FrmImp2.ProgBar.max = llave_rep01.RowCount
K_COBRA_SOLES = 0
K_COBRA_DOLAR = 0
K_COBRA_SOLES_DOLAR = 0
Do Until llave_rep01.EOF
 '  F1 = F1 + 1
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
     '  xl.Cells(F1, 1) = Left(llave_rep01!all_concepto, 20)
     '  xl.Cells(F1, 7) = "'S/."
     '  xl.Cells(F1, 8) = Format(llave_rep01!ALL_IMPORTE, "0.000")
       K_COBRA_SOLES = K_COBRA_SOLES + val(llave_rep01!ALL_IMPORTE)
       K_COBRA_SOLES_DOLAR = K_COBRA_SOLES_DOLAR + val(llave_rep01!ALL_IMPORTE)
    Else
       WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
    '   xl.Cells(F1, 1) = "COMPRA US$."
    '   xl.Cells(F1, 9) = "'US$."
    '   xl.Cells(F1, 10) = Format(llave_rep01!ALL_IMPORTE, "0.000")
       K_COBRA_DOLAR = K_COBRA_DOLAR + val(llave_rep01!ALL_IMPORTE)
       K_COBRA_SOLES_DOLAR = K_COBRA_SOLES_DOLAR + redondea(val(llave_rep01!ALL_IMPORTE) * WS_TIPO_C)
 End If

llave_rep01.MoveNext
Loop

'If K_COBRA_SOLES_DOLAR <> 0 Then
 f1 = f1 + 2
 xl.Cells(f1, 4) = "'INGRESOS VARIOS = "
 xl.Cells(f1, 7) = "'S/."
 xl.Cells(f1, 8) = K_COBRA_SOLES
 xl.Cells(f1, 9) = "'US$."
 xl.Cells(f1, 10) = K_COBRA_DOLAR
 xl.Cells(f1, 11) = "'S/."
 xl.Cells(f1, 12) = K_COBRA_SOLES_DOLAR
 wranF = "D" & f1 & ":L" & f1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 wranF = "D" & f1 + 1 & ":L" & f1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
'End If
    

f1 = f1 + 2
xl.Cells(f1, 1) = "'TOTAL DE INGRESOS = "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = (TT_TOTAL_SOLES + SOLES_TOTAL_COBRA + K_COBRA_SOLES)
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = (TT_TOTAL_DOLAR + DOLAR_TOTAL_COBRA + K_COBRA_DOLAR)
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = (TT_DOLAR_SOLES + TOTAL_COBRA_SD + K_COBRA_SOLES_DOLAR)
wranF = "A" & f1 & ":L" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "A" & f1 + 1 & ":L" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
    
    
f1 = f1 + 2
xl.Cells(f1, 1) = "'EGRESOS: "
f1 = f1 + 1

If wusu = "A" Then
   pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10 AND ALL_CODUSU = ? AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
Else
   pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10   AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
  PS_REP01(0) = 0
  PS_REP01(1) = 0
Else
  PS_REP01(0) = 0
  PS_REP01(1) = 0
End If
PS_REP01(2) = LK_FECHA_DIA
If wusu = "A" Then
  PS_REP01(3) = 0
End If
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
  PS_REP01(0) = "01"
  PS_REP01(1) = "02"
Else
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If


llave_rep01.Requery

f1 = f1 + 2
xl.Cells(f1 - 1, 1) = "(-) CREDITOS "
' A SOLO FIRMA AL CREDITO
xl.Cells(f1, 1) = "A SOLO FIRMA:"
llave_rep01.MoveFirst
Do Until llave_rep01.EOF
 If llave_rep01!ALL_SIGNO_CAJA <> 0 Then GoTo PASA_CRED
 If llave_rep01!all_SECUENCIA <> 24 Then GoTo PASA_CRED
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    pu_cp = llave_rep01!ALL_CP
    pu_codclie = val(llave_rep01!ALL_CODCLIE)
    LEER_CLI_LLAVE
    xl.Cells(f1, 1) = Trim(cli_llave!CLI_NOMBRE)
    xl.Cells(f1, 2) = llave_rep01!ALL_FBG
    xl.Cells(f1, 3) = llave_rep01!ALL_NUMSER
    xl.Cells(f1, 4) = llave_rep01!all_numfac
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
        xl.Cells(f1, 7) = "'S/."
        xl.Cells(f1, 8) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000")
    Else
        xl.Cells(f1, 9) = "'US$."
        xl.Cells(f1, 10) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000")
    End If
PASA_CRED:
llave_rep01.MoveNext
Loop
' CON CHEQUE AL BANCO
f1 = f1 + 2
xl.Cells(f1, 1) = "REG. POR ANTICIPO:"
llave_rep01.MoveFirst
Do Until llave_rep01.EOF
 If llave_rep01!ALL_SIGNO_CAJA <> 0 Then GoTo PASA_CHE
 If llave_rep01!all_SECUENCIA <> 50 Then GoTo PASA_CHE
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    pu_cp = llave_rep01!ALL_CP
    pu_codclie = val(llave_rep01!ALL_CODCLIE)
    LEER_CLI_LLAVE
    xl.Cells(f1, 1) = Trim(cli_llave!CLI_NOMBRE)
    xl.Cells(f1, 2) = llave_rep01!ALL_FBG
    xl.Cells(f1, 3) = llave_rep01!ALL_NUMSER
    xl.Cells(f1, 4) = llave_rep01!all_numfac
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
        xl.Cells(f1, 7) = "'S/."
        xl.Cells(f1, 8) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000")
    Else
        xl.Cells(f1, 9) = "'US$."
        xl.Cells(f1, 10) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000")
    End If
PASA_CHE:
llave_rep01.MoveNext
Loop
 
f1 = f1 + 2
xl.Cells(f1, 4) = "TOTAL CREDITOS "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = WSOLES_CRED
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = WDOLAR_CRED
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = WSOLES_DOLAR_CRED

wranF = "D" & f1 & ":L" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "D" & f1 + 1 & ":L" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1



' EGRESOS  VARIOS DE CAJA
If wusu = "A" Then
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_CODUSU = ? AND ALL_CODTRA = 5355 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
Else
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ?  AND ALL_CODTRA = 5355 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
 PS_REP01(0) = 0
 PS_REP01(1) = 0
Else
 PS_REP01(0) = 0
 PS_REP01(1) = 0
End If
PS_REP01(2) = LK_FECHA_DIA
If wusu = "A" Then
  PS_REP01(3) = 0
End If

Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
 PS_REP01(0) = "01"
 PS_REP01(1) = "02"
Else
 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If

llave_rep01.Requery

f1 = f1 + 2
xl.Cells(f1, 1) = "'EGRESOS VARIOS: "
xl.Cells(f1, 5) = ""
xl.Cells(f1, 6) = ""
K2_COBRA_SOLES = 0
K2_COBRA_DOLAR = 0
K2_COBRA_SOLES_DOLAR = 0
Do Until llave_rep01.EOF
   f1 = f1 + 1
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
       xl.Cells(f1, 1) = Left(llave_rep01!all_concepto, 20)
       xl.Cells(f1, 7) = "'S/."
       xl.Cells(f1, 8) = Format(llave_rep01!ALL_IMPORTE, "0.000")
       K2_COBRA_SOLES = K2_COBRA_SOLES + val(llave_rep01!ALL_IMPORTE)
       K2_COBRA_SOLES_DOLAR = K2_COBRA_SOLES_DOLAR + val(llave_rep01!ALL_IMPORTE)
    Else
       WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
       xl.Cells(f1, 1) = "COMPRA US$."
       xl.Cells(f1, 9) = "'US$."
       xl.Cells(f1, 10) = Format(llave_rep01!ALL_IMPORTE, "0.000")
       K2_COBRA_DOLAR = K2_COBRA_DOLAR + val(llave_rep01!ALL_IMPORTE)
       K2_COBRA_SOLES_DOLAR = K2_COBRA_SOLES_DOLAR + redondea(val(llave_rep01!ALL_IMPORTE) * WS_TIPO_C)
 End If

llave_rep01.MoveNext
Loop

'If K2_COBRA_SOLES_DOLAR <> 0 Then
  f1 = f1 + 1
  xl.Cells(f1, 4) = "'EGRESOS VARIOS = "
  xl.Cells(f1, 7) = "'S/."
  xl.Cells(f1, 8) = K2_COBRA_SOLES
  xl.Cells(f1, 9) = "'US$."
  xl.Cells(f1, 10) = K2_COBRA_DOLAR
  xl.Cells(f1, 11) = "'S/."
  xl.Cells(f1, 12) = K2_COBRA_SOLES_DOLAR
  var_soles = K2_COBRA_SOLES
  var_dolar = K2_COBRA_DOLAR
  var_soles_dolar = K2_COBRA_SOLES_DOLAR
  wranF = "D" & f1 & ":J" & f1
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  wranF = "D" & f1 + 1 & ":J" & f1 + 1
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
'End If

  
f1 = f1 + 2
xl.Cells(f1, 1) = "'TOTAL DE EGRESOS = "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = (K2_COBRA_SOLES + WSOLES_CRED)
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = (K2_COBRA_DOLAR + WDOLAR_CRED)
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = (K2_COBRA_SOLES_DOLAR + WSOLES_DOLAR_CRED)
wranF = "A" & f1 & ":L" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "A" & f1 + 1 & ":L" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1

f1 = f1 + 2
xl.Cells(f1, 1) = "'SALDO POR DEPOSITAR = "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = (TT_TOTAL_SOLES + SOLES_TOTAL_COBRA + K_COBRA_SOLES) - (K2_COBRA_SOLES + WSOLES_CRED)
TOTALS1 = (TT_TOTAL_SOLES + SOLES_TOTAL_COBRA + K_COBRA_SOLES) - (K2_COBRA_SOLES + WSOLES_CRED)
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = (TT_TOTAL_DOLAR + DOLAR_TOTAL_COBRA + K_COBRA_DOLAR) - (K2_COBRA_DOLAR + WDOLAR_CRED)
TOTALD1 = (TT_TOTAL_DOLAR + DOLAR_TOTAL_COBRA + K_COBRA_DOLAR) - (K2_COBRA_DOLAR + WDOLAR_CRED)
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = (TT_DOLAR_SOLES + TOTAL_COBRA_SD + K_COBRA_SOLES_DOLAR) - (K2_COBRA_SOLES_DOLAR + WSOLES_DOLAR_CRED)
TOTALSD1 = (TT_DOLAR_SOLES + TOTAL_COBRA_SD + K_COBRA_SOLES_DOLAR) - (K2_COBRA_SOLES_DOLAR + WSOLES_DOLAR_CRED)
wranF = "A" & f1 & ":L" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "A" & f1 + 1 & ":L" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1




If wusu = "A" Then
   pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10 AND ALL_CODUSU = ? AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
Else
   pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_TIPMOV = 10   AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FBG, ALL_NUMSER,  ALL_NUMFAC"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
  PS_REP01(0) = 0
  PS_REP01(1) = 0
Else
  PS_REP01(0) = 0
  PS_REP01(1) = 0
End If
PS_REP01(2) = LK_FECHA_DIA
If wusu = "A" Then
  PS_REP01(3) = 0
End If
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
  PS_REP01(0) = "01"
  PS_REP01(1) = "02"
Else
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If


llave_rep01.Requery


K_DET_SOLES = 0
K_DET_DOLAR = 0
K_DET_SOLES_DOLAR = 0

f1 = f1 + 1
xl.Cells(f1, 1) = "DETERMINACION DE EFECTIVO NETO:"
f1 = f1 + 1
xl.Cells(f1, 1) = "CHEQUES RECIBIDOS:"
xl.Cells(f1, 5) = "NRO."
xl.Cells(f1, 6) = "BANCO"
llave_rep01.MoveFirst
Do Until llave_rep01.EOF
 If llave_rep01!ALL_SIGNO_CAJA <> 1 Then GoTo PASA_CHE2
 If llave_rep01!all_SECUENCIA <> 30 Then GoTo PASA_CHE2
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    pu_cp = llave_rep01!ALL_CP
    pu_codclie = val(llave_rep01!ALL_CODCLIE)
    LEER_CLI_LLAVE
    xl.Cells(f1, 1) = Trim(cli_llave!CLI_NOMBRE)
    xl.Cells(f1, 2) = llave_rep01!ALL_FBG
    xl.Cells(f1, 3) = llave_rep01!ALL_NUMSER
    xl.Cells(f1, 4) = llave_rep01!all_numfac
    xl.Cells(f1, 5) = llave_rep01!all_chenum
    PUB_CODBAN = llave_rep01!all_codban
    pu_codcia = LK_CODCIA
    LEER_CCM_LLAVE
    xl.Cells(f1, 6) = Left(ccm_llave!CCM_NOMBRE, 10)
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
        xl.Cells(f1, 7) = "'S/."
        xl.Cells(f1, 8) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000")
        K_DET_SOLES = K_DET_SOLES + val(llave_rep01!ALL_IMPORTE_AMORT)
        K_DET_SOLES_DOLAR = K_DET_SOLES_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
    Else
        WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
        xl.Cells(f1, 9) = "'US$."
        xl.Cells(f1, 10) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000")
        K_DET_DOLAR = K_DET_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
        K_DET_SOLES_DOLAR = K_DET_SOLES_DOLAR + redondea(val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
    End If
PASA_CHE2:
llave_rep01.MoveNext
Loop
If K_DET_SOLES_DOLAR <> 0 Then
' F1 = F1 + 1
' xl.Cells(F1, 4) = "TOTAL = "
' xl.Cells(F1, 7) = "'S/."
' xl.Cells(F1, 8) = K_DET_SOLES
' xl.Cells(F1, 9) = "'US$."
' xl.Cells(F1, 10) = K_DET_DOLAR
' xl.Cells(F1, 11) = "'S/."
' xl.Cells(F1, 12) = K_DET_SOLES_DOLAR
' wranF = "D" & F1 & ":L" & F1
' xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
' wranF = "D" & F1 + 1 & ":L" & F1 + 1
' xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
End If

 
f1 = f1 + 2
xl.Cells(f1, 1) = "VENTAS DEPOSITADAS:"
xl.Cells(f1, 5) = "OP."
xl.Cells(f1, 6) = "BANCO"
llave_rep01.MoveFirst
Do Until llave_rep01.EOF
 If llave_rep01!ALL_SIGNO_CAJA <> 1 Then GoTo PASA_CHE3
 If llave_rep01!all_SECUENCIA <> 40 Then GoTo PASA_CHE3
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    pu_cp = llave_rep01!ALL_CP
    pu_codclie = val(llave_rep01!ALL_CODCLIE)
    LEER_CLI_LLAVE
    xl.Cells(f1, 1) = Trim(cli_llave!CLI_NOMBRE)
    xl.Cells(f1, 2) = llave_rep01!ALL_FBG
    xl.Cells(f1, 3) = llave_rep01!ALL_NUMSER
    xl.Cells(f1, 4) = llave_rep01!all_numfac
    xl.Cells(f1, 5) = llave_rep01!all_chenum
    PUB_CODBAN = llave_rep01!all_codban
    pu_codcia = llave_rep01!all_CODCIA
    LEER_CCM_LLAVE
    xl.Cells(f1, 6) = Left(ccm_llave!CCM_NOMBRE, 10)
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
        xl.Cells(f1, 7) = "'S/."
        xl.Cells(f1, 8) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000")
        K_DET_SOLES = K_DET_SOLES + val(llave_rep01!ALL_IMPORTE_AMORT)
        K_DET_SOLES_DOLAR = K_DET_SOLES_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
    Else
        WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
        xl.Cells(f1, 9) = "'US$."
        xl.Cells(f1, 10) = Format(llave_rep01!ALL_IMPORTE_AMORT, "0.000")
        K_DET_DOLAR = K_DET_DOLAR + val(llave_rep01!ALL_IMPORTE_AMORT)
        K_DET_SOLES_DOLAR = K_DET_SOLES_DOLAR + redondea(val(llave_rep01!ALL_IMPORTE_AMORT) * WS_TIPO_C)
    End If
PASA_CHE3:
llave_rep01.MoveNext
Loop


If wusu = "A" Then
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_CODUSU = ? AND ALL_TIPMOV <> 10 AND ALL_SIGNO_CAJA = 1 AND all_flag_ext <> 'E' AND ALL_SIGNO_CAR <> 0 AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
Else
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ?  AND ALL_TIPMOV <> 10 AND ALL_SIGNO_CAJA = 1 AND all_flag_ext <> 'E' AND ALL_SIGNO_CAR <> 0  AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
 PS_REP01(0) = 0
 PS_REP01(1) = 0
Else
 PS_REP01(0) = 0
 PS_REP01(1) = 0
End If
PS_REP01(2) = LK_FECHA_DIA
If wusu = "A" Then
  PS_REP01(3) = 0
End If
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
 PS_REP01(0) = "01"
 PS_REP01(1) = "02"
Else
 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If

llave_rep01.Requery


SOLES_TOTAL_COBRA = 0
DOLAR_TOTAL_COBRA = 0
TOTAL_COBRA_SD = 0
f1 = f1 + 1
xl.Cells(f1, 1) = "'COBRANZAS DEPOSITADAS"

' COBRANZA CON CHEQUE
llave_rep01.MoveFirst
Do Until llave_rep01.EOF
'    FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
    If llave_rep01!all_SECUENCIA <> 30 Then GoTo PASA_COBRA2
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = llave_rep01!all_CODCIA
    pu_cp = llave_rep01!ALL_CP
    pu_codclie = val(llave_rep01!ALL_CODCLIE)
    LEER_CLI_LLAVE
    xl.Cells(f1, 1) = Trim(cli_llave!CLI_NOMBRE)
    xl.Cells(f1, 2) = llave_rep01!ALL_FBG
    xl.Cells(f1, 3) = llave_rep01!all_numser_c
    xl.Cells(f1, 4) = llave_rep01!all_numfac_c
    PUB_CODBAN = llave_rep01!all_codban
    pu_codcia = llave_rep01!all_CODCIA
    LEER_CCM_LLAVE
    xl.Cells(f1, 5) = llave_rep01!all_chenum
    xl.Cells(f1, 6) = Left(ccm_llave!CCM_NOMBRE, 10)
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
       xl.Cells(f1, 7) = "'S/."
       WIMPORTE = llave_rep01!ALL_IMPORTE_AMORT
       If llave_rep01!ALL_MONEDA_CLI = "D" Then
         WIMPORTE = redondea(WIMPORTE * llave_rep01!ALL_TIPO_CAMBIO)
       End If
       xl.Cells(f1, 8) = Format(WIMPORTE, "0.000")
       
       SOLES_TOTAL_COBRA = SOLES_TOTAL_COBRA + WIMPORTE
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + WIMPORTE
       K_COBRA_SOLES = K_COBRA_SOLES + val(WIMPORTE)
       K_COBRA_SOLES_DOLAR = K_COBRA_SOLES_DOLAR + val(WIMPORTE)
    Else
       WIMPORTE = llave_rep01!ALL_IMPORTE_AMORT
       If llave_rep01!ALL_MONEDA_CLI = "S" Then
         WIMPORTE = redondea(llave_rep01!ALL_IMPORTE_AMORT / llave_rep01!ALL_TIPO_CAMBIO)
       End If
       xl.Cells(f1, 9) = "'US$."
       xl.Cells(f1, 10) = Format(WIMPORTE, "0.000")
       WS_TIPO_C = llave_rep01!ALL_TIPO_CAMBIO
       DOLAR_TOTAL_COBRA = DOLAR_TOTAL_COBRA + WIMPORTE
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + val(Format((WIMPORTE * WS_TIPO_C), "0.000")) 'Val(llave_rep01!ALL_TIPO_CAMBIO))
       K_COBRA_DOLAR = K_COBRA_DOLAR + WIMPORTE
       K_COBRA_SOLES_DOLAR = K_COBRA_SOLES_DOLAR + val(Format((WIMPORTE * WS_TIPO_C), "0.000"))
 End If
PASA_COBRA2:
llave_rep01.MoveNext
Loop

' COBRANZA CON CHEQUE
llave_rep01.MoveFirst
Do Until llave_rep01.EOF
'    FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
    If llave_rep01!all_SECUENCIA <> 40 Then GoTo PASA_COBRA3
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = llave_rep01!all_CODCIA
    pu_cp = llave_rep01!ALL_CP
    pu_codclie = val(llave_rep01!ALL_CODCLIE)
    LEER_CLI_LLAVE
    xl.Cells(f1, 1) = Trim(cli_llave!CLI_NOMBRE)
    xl.Cells(f1, 2) = llave_rep01!ALL_FBG
    xl.Cells(f1, 3) = llave_rep01!all_numser_c
    xl.Cells(f1, 4) = llave_rep01!all_numfac_c
    PUB_CODBAN = llave_rep01!all_codban
    pu_codcia = llave_rep01!all_CODCIA
    LEER_CCM_LLAVE
    xl.Cells(f1, 5) = llave_rep01!all_chenum
    xl.Cells(f1, 6) = Left(ccm_llave!CCM_NOMBRE, 10)
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
       WIMPORTE = llave_rep01!ALL_IMPORTE_AMORT
       If llave_rep01!ALL_MONEDA_CLI = "D" Then
         WIMPORTE = redondea(WIMPORTE * llave_rep01!ALL_TIPO_CAMBIO)
       End If
       xl.Cells(f1, 7) = "'S/."
       xl.Cells(f1, 8) = Format(WIMPORTE, "0.000")
       SOLES_TOTAL_COBRA = SOLES_TOTAL_COBRA + WIMPORTE
       TOTAL_COBRA_SD = TOTAL_COBRA_SD + WIMPORTE
       K2_COBRA_SOLES = K2_COBRA_SOLES + WIMPORTE
       K2_COBRA_SOLES_DOLAR = K2_COBRA_SOLES_DOLAR + WIMPORTE
    Else
        WIMPORTE = llave_rep01!ALL_IMPORTE_AMORT
        If llave_rep01!ALL_MONEDA_CLI = "S" Then
         WIMPORTE = redondea(llave_rep01!ALL_IMPORTE_AMORT / llave_rep01!ALL_TIPO_CAMBIO)
        End If
        xl.Cells(f1, 9) = "'US$."
        xl.Cells(f1, 10) = Format(WIMPORTE, "0.000")
        WS_TIPO_C = llave_rep01!ALL_TIPO_CAMBIO
        DOLAR_TOTAL_COBRA = DOLAR_TOTAL_COBRA + WIMPORTE
        TOTAL_COBRA_SD = TOTAL_COBRA_SD + val(Format((WIMPORTE * WS_TIPO_C), "0.000"))
        K2_COBRA_DOLAR = K2_COBRA_DOLAR + val(WIMPORTE)
        K2_COBRA_SOLES_DOLAR = K2_COBRA_SOLES_DOLAR + val(Format((WIMPORTE * WS_TIPO_C), "0.000"))
    End If
 
PASA_COBRA3:
llave_rep01.MoveNext
Loop

If wusu = "A" Then
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_CODUSU = ? AND ALL_CODTRA = 5350 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
Else
pub_cadena = "SELECT  * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ?  AND ALL_CODTRA = 5350 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
If checia.Value = 1 Then
 PS_REP01(0) = 0
 PS_REP01(1) = 0
Else
 PS_REP01(0) = 0
 PS_REP01(1) = 0
End If
PS_REP01(2) = LK_FECHA_DIA
If wusu = "A" Then
  PS_REP01(3) = 0
End If
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If checia.Value = 1 Then
 PS_REP01(0) = "01"
 PS_REP01(1) = "02"
Else
 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = ""
End If
PS_REP01(2) = wsFECHA1
If wusu = "A" Then
  PS_REP01(3) = LK_CODUSU
End If

llave_rep01.Requery

K_COBRA_SOLES = 0
K_COBRA_SOLES_DOLAR = 0
K_COBRA_DOLAR = 0

f1 = f1 + 1
xl.Cells(f1, 1) = "'OTROS DEPOSITOS: "
Do Until llave_rep01.EOF
   f1 = f1 + 1
    If llave_rep01!ALL_MONEDA_CAJA = "S" Then
       xl.Cells(f1, 1) = Trim(llave_rep01!all_concepto)
       xl.Cells(f1, 7) = "'S/."
       xl.Cells(f1, 8) = Format(llave_rep01!ALL_IMPORTE, "0.000")
       K_COBRA_SOLES = K_COBRA_SOLES + val(llave_rep01!ALL_IMPORTE)
       K_COBRA_SOLES_DOLAR = K_COBRA_SOLES_DOLAR + val(llave_rep01!ALL_IMPORTE)
    Else
       WS_TIPO_C = JALAR(llave_rep01!ALL_FECHA_SUNAT)
       xl.Cells(f1, 1) = "COMPRA US$."
       xl.Cells(f1, 9) = "'US$."
       xl.Cells(f1, 10) = Format(llave_rep01!ALL_IMPORTE, "0.000")
       K_COBRA_DOLAR = K_COBRA_DOLAR + val(llave_rep01!ALL_IMPORTE)
       K_COBRA_SOLES_DOLAR = K_COBRA_SOLES_DOLAR + redondea(val(llave_rep01!ALL_IMPORTE) * WS_TIPO_C)
 End If

llave_rep01.MoveNext
Loop

f1 = f1 + 2
xl.Cells(f1, 4) = "TOTAL = "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = (K_DET_SOLES + SOLES_TOTAL_COBRA + K_COBRA_SOLES)
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = (K_DET_DOLAR + DOLAR_TOTAL_COBRA + K_COBRA_DOLAR)
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = (K_DET_SOLES_DOLAR + TOTAL_COBRA_SD + K_COBRA_SOLES_DOLAR)

 '  otros egresos de deposito
f1 = f1 + 2
xl.Cells(f1, 1) = "'OTROS EGRESOS: "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = var_soles
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = var_dolar
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = var_soles_dolar

f1 = f1 + 2
xl.Cells(f1, 1) = "'POR DEPOSITAR = "
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = TOTALS1 - (K_DET_SOLES + SOLES_TOTAL_COBRA + K_COBRA_SOLES) + var_soles
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = TOTALD1 - (K_DET_DOLAR + DOLAR_TOTAL_COBRA + K_COBRA_DOLAR) + var_dolar
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = TOTALSD1 - (K_DET_SOLES_DOLAR + TOTAL_COBRA_SD + K_COBRA_SOLES_DOLAR) + var_soles_dolar
wranF = "A" & f1 & ":L" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "A" & f1 + 1 & ":L" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1





'VENTA_TOTAL_SOLES = TT_TOTAL_SOLES - WSOLES_CRED - K_DET_SOLES
'VENTA_TOTAL_DOLAR = TT_TOTAL_DOLAR - WDOLAR_CRED - K_DET_DOLAR
'VENTA_TOTAL_SOLES_DOLAR = TT_DOLAR_SOLES - WSOLES_DOLAR_CRED ' K_DET_SOLES_DOLAR

wranF = "D" & f1 & ":L" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "D" & f1 + 1 & ":L" & f1 + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1




    
DoEvents
FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
DoEvents
xcuenta = 1
xl.APPLICATION.Visible = True
xl.Cells(2, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
If checia.Value = 1 Then
  xl.Cells(2, 1) = "Almacenes 01 - 02 "
End If
xl.Cells(1, 10) = "IMP.: " & Now
xl.Cells(4, 10) = "FECHA:"
xl.Cells(4, 11) = "'" & Format(wsFECHA1, "dd mmm yyyy")
xl.DisplayAlerts = False
xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
xl.Worksheets("Hoja1").Protect PUB_CLAVE
xl.APPLICATION.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = False
FrmImp2.ProgBar.Visible = False
Set xl = Nothing
Screen.MousePointer = 0
FrmImp2.Pantalla.Enabled = True
FrmImp2.cerrar.Enabled = True
FrmImp2.Pantalla.Caption = "Por &Pantalla"
FrmImp2.lblproceso.Visible = False

Exit Sub

IMP_LINEA:
f1 = f1 + 1
If ww_fbg = "F" Then
 xl.Cells(f1, 1) = "FACTURAS"
ElseIf ww_fbg = "B" Then
 xl.Cells(f1, 1) = "BOLETAS"
ElseIf ww_fbg = "G" Then
 xl.Cells(f1, 1) = "GUIAS"
Else
 xl.Cells(f1, 1) = "PENDIENTES"
End If
xl.Cells(f1, 2) = "Nº "
xl.Cells(f1, 3) = ww_serini
xl.Cells(f1, 4) = ww_numini
xl.Cells(f1, 5) = "'al Nº"
xl.Cells(f1, 6) = ww_numfin
xl.Cells(f1, 7) = "'S/."
xl.Cells(f1, 8) = Format(WMONTO_SOLES, "0.000")
xl.Cells(f1, 9) = "'US$."
xl.Cells(f1, 10) = Format(WMONTO_DOLAR, "0.000")
xl.Cells(f1, 11) = "'S/."
xl.Cells(f1, 12) = Format(WSOLES_DOLAR, "0.000")

WMONTO_SOLES = 0
WMONTO_DOLAR = 0
WSOLES_DOLAR = 0
Return



WEXCEL:
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FrmImp2.lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
  DoEvents
  WPAS = "131296"
  'WPAS = PUB_RUTA_OTRO + "CAJA_DET.xls"
  'DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open PUB_RUTA_OTRO & "\RESU_CAJA.xls", 0, True, 4, WPAS, WPAS

  'xl.Workbooks.Open , "C:\ADMIN\HERTISA\CAJA_DET.XLS", 0, True, 4, PUB_CLAVE, PUB_CLAVE
Return

Exit Sub

'LLENA_VALOR:
'For I = QJ To WDIF
'  If I >= 3 Then
'    xl.Cells(F1 + 7, I) = Format(LOC_VALOR, "0.000")
'  End If
'Next I
'Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
Exit Sub

End Sub

Public Sub RESU_REGISTRO()
Dim qver_onlyCont As Integer
Dim ts_suma  As Currency
Dim ts_suma_bruto  As Currency
Dim ts_suma_igv  As Currency
Dim ts_codcta As String
Dim Fres As Integer
Dim w_exo As Currency
Dim ws_codcia As String * 2
Dim FILAS As Integer
Dim nn, m_ind As Integer
Dim AWQ_NETO_ACT_FIJO As Currency
Dim AWQ_CTA_ACT_FIJO As String
Dim AWQ_BRUTO_ACT_FIJO As Currency
Dim WCONTROL As Integer
FILAS = 5
Dim WQ
Dim wfecha
Dim ws_clave
Dim LETRAS(24) As String * 1
Dim wRuta As String
Dim wmonto As Currency
Dim wcodclie As Currency
Dim valor_venta As Currency
Dim descto As Currency
Dim valor_igv As Currency
Dim valor_precio As Currency
Dim S_VALOR_VENTA As Currency
Dim s_descto As Currency
Dim s_valor_igv As Currency
Dim S_VALOR_PRECIO As Currency

Dim t_valor_venta As Currency
Dim t_descto As Currency
Dim t_valor_igv As Currency
Dim t_valor_precio As Currency

Dim wnumfac As Currency
Dim WFLAG As String * 1
Dim wsFECHA1, wsFECHA2
Dim xcuenta As Integer
Dim wq_fecha, wq_codclie, wq_codven, wq_docu, wq_nombre, wq_bruto, wq_gastos, wq_desto, wq_flete, wq_fbg, wq_serie
Dim wq_tot_descto, WQ_IMPTO, WQ_TOTAL, wq_estado, wq_condi
Dim wq_ruc
Dim wflag_numfac
Dim wserie As String * 3
Dim AWQ_BRUTO As Currency
Dim AWQ_DESCTOS As Currency
Dim AWQ_GASTOS As Currency
Dim AWQ_FLETES As Currency
Dim AWQ_IMPTO As Currency
Dim AWQ_NETO As Currency
Dim AWQ_NETO_CRED  As Currency
Dim AWQ_NETO_CONT   As Currency
Dim AWQ_COSTO_VENTA As Currency
Dim WS_SIGNO As Integer
Dim ws_tc As Currency
Dim wsTexto  As String

Dim ww_fbg
Dim ww_numser
Dim ww_serini
Dim ww_numfac
Dim ww_numini
Dim ww_numfin
'wnumfac



Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If

GoSub WEXCEL
pub_cadena = ""
xcuenta = 0

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
WCONTROL = 0
f1 = 5  'Fila Inicial
t_valor_venta = 0
t_descto = 0
t_valor_igv = 0
t_valor_precio = 0

'NCREDITO: ' empieza

'WCONTROL = WCONTROL + 1
wmonto = 0
wcodclie = 0
valor_venta = 0
descto = 0
valor_igv = 0
valor_precio = 0
wnumfac = 0
WFLAG = ""
xcuenta = 0
wflag_numfac = ""
wserie = ""
Dim SS_VALOR_VENTA As Currency
Dim SS_VALOR_PRECIO As Currency
Dim SS_VALOR_IGV As Currency



Dim TD_F As String * 1
Dim TD_B As String * 1
Dim TD_N As String * 1
Dim TD_D As String * 1


TD_F = " "
TD_B = " "
TD_N = " "
TD_D = " "
wsTexto = "TIPO: "
For fila = 0 To 3
 txttp.ListIndex = fila
 If fila = 0 And txttp.Selected(fila) Then
   TD_F = "F"
   wsTexto = wsTexto + "- FACT."
 End If
 If fila = 1 And txttp.Selected(fila) Then
   TD_B = "B"
   wsTexto = wsTexto + "- BOL."
 End If
 If fila = 2 And txttp.Selected(fila) Then
   TD_N = "N"
   wsTexto = wsTexto + "- NCRED."
 End If
 If fila = 3 And txttp.Selected(fila) Then
   TD_D = "D"
   wsTexto = wsTexto + "- NDEB."
 End If
Next fila

pub_cadena = "SELECT CLI_CUENTA_CONTAB FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_CODCLIE = ? AND CLI_CP = 'C' "
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
PS_REP03(0) = 0
PS_REP03(1) = 0
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


pub_cadena = "SELECT FAR_CODCLIE FROM FACART WHERE  FAR_CODCIA = ? AND FAR_FBG = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ?  AND  FAR_TIPMOV = ? AND FAR_ESTADO <> 'E' "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = 0
PS_REP01(3) = 0
PS_REP01(4) = 0
PS_REP01.MaxRows = 1
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

Fres = 0
SS_VALOR_VENTA = 0
SS_VALOR_IGV = 0
SS_VALOR_PRECIO = 0

NCREDITO:
S_VALOR_VENTA = 0
s_descto = 0
s_valor_igv = 0
S_VALOR_PRECIO = 0


WCONTROL = WCONTROL + 1
pub_cadena = "SELECT DISTINCT FAR_TIPMOV, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_FECHA_COMPRA, FAR_CODCIA, FAR_MONEDA, FAR_BRUTO, FAR_IMPTO, FAR_TOT_DESCTO, FAR_EX_IGV  FROM FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? ) AND (FAR_TIPMOV = 10 OR FAR_TIPMOV = 97 OR FAR_TIPMOV = 98) AND FAR_ESTADO<>'E' AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ?)  ORDER BY FAR_FECHA_COMPRA "
pub_cadena = "SELECT DISTINCT FAR_CODCIA, FAR_TIPMOV, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_MONEDA, FAR_BRUTO, FAR_IMPTO, FAR_TOT_DESCTO, FAR_EX_IGV, FAR_FECHA_COMPRA, FAR_CODCIA FROM FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? ) AND (FAR_TIPMOV = 10 OR FAR_TIPMOV = 97 OR FAR_TIPMOV = 98) AND FAR_ESTADO <>'E' AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ?)  ORDER BY FAR_FECHA_COMPRA "
If fbtxt.Text <> "" And txtSerie.Text = "" Then
   pub_cadena = "SELECT DISTINCT FAR_CODCIA, FAR_TIPMOV, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_FECHA_COMPRA, FAR_CODCIA, FAR_MONEDA, FAR_BRUTO, FAR_IMPTO, FAR_TOT_DESCTO, FAR_EX_IGV  FROM FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND ( FAR_TIPMOV = 10 OR FAR_TIPMOV = 97 OR FAR_TIPMOV = 98 ) AND FAR_ESTADO<>'E' AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?  AND (FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ?) ORDER BY FAR_FECHA_COMPRA "
   PS_REP02(7) = ""
ElseIf fbtxt.Text <> "" And txtSerie.Text <> "" Then
   pub_cadena = "SELECT DISTINCT FAR_CODCIA, FAR_TIPMOV, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_FECHA_COMPRA, FAR_CODCIA, FAR_MONEDA, FAR_BRUTO, FAR_IMPTO, FAR_TOT_DESCTO, FAR_EX_IGV  FROM FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND ( FAR_TIPMOV = 10 OR FAR_TIPMOV = 97 OR FAR_TIPMOV = 98 ) AND FAR_ESTADO<>'E' AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?  AND (FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ?) AND FAR_NUMSER = ? ORDER BY FAR_FECHA_COMPRA "
   PS_REP02(7) = ""
   PS_REP02(8) = ""
ElseIf fbtxt.Text = "" And txtSerie.Text <> "" Then
   pub_cadena = "SELECT DISTINCT FAR_CODCIA, FAR_TIPMOV, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_FECHA_COMPRA, FAR_CODCIA, FAR_MONEDA, FAR_BRUTO, FAR_IMPTO, FAR_TOT_DESCTO, FAR_EX_IGV  FROM FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND ( FAR_TIPMOV = 10 OR FAR_TIPMOV = 97 OR FAR_TIPMOV = 98  ) AND FAR_ESTADO<>'E' AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?  AND FAR_NUMSER = ? AND (FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ? OR FAR_FBG = ?) ORDER BY FAR_FECHA_COMPRA "
End If
WS_SIGNO = 1
  
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = ""
PS_REP02(1) = ""
PS_REP02(2) = ""
PS_REP02(3) = ""
PS_REP02(4) = ""
PS_REP02(5) = 0
PS_REP02(6) = 0
PS_REP02(7) = 0
PS_REP02(8) = 0
PS_REP02(9) = 0
PS_REP02(10) = 0

Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

If checia.Visible And checia.Value = 1 Then
 If Trim(par_llave!par_art_cias) <> "" Then
      nn = 1
    For m_ind = 1 To 15
        ws_codcia = Mid(par_llave!par_art_cias, nn, 2)
        If Trim(ws_codcia) = "" Then Exit For
        PS_REP02(m_ind - 1) = ws_codcia
        nn = nn + 2
    Next m_ind
 End If
Else
 PS_REP02(0) = LK_CODCIA
End If

' el PS_REP1(0) ESTA MAS ABAJO
PS_REP02(5) = wsFECHA1
PS_REP02(6) = wsFECHA2
PS_REP02(7) = TD_F
PS_REP02(8) = TD_B
PS_REP02(9) = TD_N
PS_REP02(10) = TD_D

If fbtxt.Text <> "" And txtSerie.Text = "" Then
 PS_REP02(7) = TD_F
 PS_REP02(8) = TD_B
 PS_REP02(9) = TD_N
 PS_REP02(10) = TD_D
ElseIf fbtxt.Text <> "" And txtSerie.Text <> "" Then
 PS_REP02(7) = TD_F
 PS_REP02(8) = TD_B
 PS_REP02(9) = TD_N
 PS_REP02(10) = TD_D
 PS_REP02(11) = txtSerie.Text
ElseIf fbtxt.Text = "" And txtSerie.Text <> "" Then
 PS_REP02(7) = txtSerie.Text
 PS_REP02(8) = TD_F
 PS_REP02(9) = TD_B
 PS_REP02(10) = TD_N
 PS_REP02(11) = TD_D
End If

DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep02.Requery
If llave_rep02.EOF Then
  GoTo OTRO_DOCUMENTO
End If
FrmImp2.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.max = llave_rep02.RowCount

WFLAG = ""
SQ_OPER = 1
pu_cp = "C"
pu_codcia = llave_rep02!FAR_CODCIA
wnumfac = llave_rep02!far_numfac
wfecha = llave_rep02!FAR_fecha_compra 'llave_rep02!far_fecha
wserie = llave_rep02!far_numser

xcuenta = 0
WFLAG = "A"
wflag_numfac = "A"
AWQ_DESCTOS = 0
AWQ_GASTOS = 0
AWQ_FLETES = 0
w_exo = 0
If llave_rep02.EOF Then GoTo CANCELA

wq_fecha = llave_rep02!FAR_fecha_compra
ww_fbg = llave_rep02!far_fbg
ww_numser = llave_rep02!far_numser
ww_serini = llave_rep02!far_numser
ww_numfac = llave_rep02!far_numfac
ww_numini = llave_rep02!far_numfac
ww_numfin = llave_rep02!far_numfac
wnumfac = llave_rep02!far_numfac

Do Until llave_rep02.EOF
  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  If wfecha <> llave_rep02!FAR_fecha_compra Then '
        GoSub IMP_LINEA
        wq_fecha = llave_rep02!FAR_fecha_compra
        wfecha = llave_rep02!FAR_fecha_compra
        ww_fbg = llave_rep02!far_fbg
        ww_numser = llave_rep02!far_numser
        ww_serini = llave_rep02!far_numser
        ww_numfac = llave_rep02!far_numfac
        ww_numini = llave_rep02!far_numfac
        wnumfac = llave_rep02!far_numfac
  End If
 
  If ww_fbg <> llave_rep02!far_fbg Then
        GoSub IMP_LINEA
        ww_fbg = llave_rep02!far_fbg
        ww_numser = llave_rep02!far_numser
        ww_serini = llave_rep02!far_numser
        ww_numfac = llave_rep02!far_numfac
        ww_numini = llave_rep02!far_numfac
        wnumfac = llave_rep02!far_numfac
  End If
  If ww_numser <> llave_rep02!far_numser Then
        GoSub IMP_LINEA
        ww_numser = llave_rep02!far_numser
        ww_serini = llave_rep02!far_numser
        ww_numfac = llave_rep02!far_numfac
        ww_numini = llave_rep02!far_numfac
        wnumfac = llave_rep02!far_numfac
  End If
  ws_tc = 1
  If llave_rep02!FAR_MONEDA = "D" Then
    ws_tc = JALAR(llave_rep02!FAR_fecha_compra)
    If ws_tc <= 0 Then
        MsgBox "Falta Ingresar el Tipo de Cambio del día : " & Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy"), 48, Pub_Titulo
        GoTo CANCELA
    End If
  End If
  If llave_rep02!FAR_TIPMOV = 97 Then
     WS_SIGNO = -1
  Else
     WS_SIGNO = 1
  End If
  
  wq_bruto = Format((val(llave_rep02!FAR_BRUTO) - val(llave_rep02!FAR_TOT_DESCTO)) * ws_tc * WS_SIGNO, "0.000")
  WQ_IMPTO = Format(val(llave_rep02!far_IMPTO) * WS_SIGNO * ws_tc, "0.000")
  WQ_TOTAL = Format((val(llave_rep02!FAR_BRUTO) + val(llave_rep02!far_IMPTO)) * WS_SIGNO * ws_tc, "0.000")
  ts_suma = val(wq_bruto) + val(WQ_IMPTO)
  If ts_suma <> val(WQ_TOTAL) Then
      wq_bruto = val(wq_bruto) + (val(WQ_TOTAL) - ts_suma) ' (Val(wq_bruto) + Val(WQ_IMPTO)) - Val(WQ_TOTAL)
  End If
  
  S_VALOR_PRECIO = S_VALOR_PRECIO + wq_bruto
  s_valor_igv = s_valor_igv + WQ_IMPTO
  S_VALOR_VENTA = S_VALOR_VENTA + WQ_TOTAL
  
  SS_VALOR_PRECIO = SS_VALOR_PRECIO + wq_bruto
  SS_VALOR_IGV = SS_VALOR_IGV + WQ_IMPTO
  SS_VALOR_VENTA = SS_VALOR_VENTA + WQ_TOTAL
  ww_numfin = llave_rep02!far_numfac
  wnumfac = wnumfac + 1
  
SALTARIN:
 llave_rep02.MoveNext
Loop
GoSub IMP_LINEA
    
  xcuenta = c1 + 1
  If WCONTROL = 1 Then
   If cheasiento.Value = 1 Then
   End If
  End If
OTRO_DOCUMENTO:
If WCONTROL >= 1 Then
Else
  GoTo NCREDITO
End If

MOSTRAR:

  f1 = f1 + 2
  xl.Cells(f1, 1) = "Total General = "
  xl.Worksheets(1).Rows(f1).RowHeight = 20
  xl.Cells(f1, 8) = Format(SS_VALOR_PRECIO, "0.000")
  xl.Cells(f1, 9) = Format(SS_VALOR_IGV, "0.000")
  xl.Cells(f1, 10) = Format(SS_VALOR_VENTA, "0.000")


' Ordenando la información para asintos.
'---------------------------------------

  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'" & wsTexto & " -  DEL " & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub

CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo RESU_REGISTRO.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\RESU_REGISTRO.xls", 0, True, 4, WPAS, WPAS
Return



FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
 GoTo CANCELA
 Resume Next
Exit Sub

IMP_LINEA:
f1 = f1 + 1
xl.Cells(f1, 1) = wq_fecha
If ww_fbg = "F" Then
 xl.Cells(f1, 2) = "'01"
ElseIf ww_fbg = "B" Then
 xl.Cells(f1, 2) = "'03"
ElseIf ww_fbg = "G" Then
 xl.Cells(f1, 2) = ""
ElseIf ww_fbg = "N" Then
 xl.Cells(f1, 2) = "'07"
ElseIf ww_fbg = "D/." Then
 xl.Cells(f1, 2) = "'08"
End If
xl.Cells(f1, 3) = "Nº "
xl.Cells(f1, 4) = ww_serini
xl.Cells(f1, 5) = ww_numini
xl.Cells(f1, 6) = "'al Nº"
xl.Cells(f1, 7) = ww_numfin
xl.Cells(f1, 8) = Format(S_VALOR_PRECIO, "0.000")
xl.Cells(f1, 9) = Format(s_valor_igv, "0.000")
xl.Cells(f1, 10) = Format(S_VALOR_VENTA, "0.000")

S_VALOR_PRECIO = 0
s_valor_igv = 0
S_VALOR_VENTA = 0

Return

End Sub

Public Sub CAJA_GENERAL()
FrmImp2.Pantalla.Enabled = False
Dim ww_muestra As String * 35
Dim ww_descrip As String
Dim WIMPORTE As Currency
Dim wnumfac As Currency
Dim wusu As String * 1
Dim ww_signo
Dim ww_sec
Dim ww_fbg
Dim WMONTO_CREDITO As Currency
Dim WMONTO_CONTADO As Currency
Dim ACU_CREDITO As Currency
Dim ACU_CONTADO As Currency
Dim COBRA_EF      As Currency
Dim COBRA_CH   As Currency
Dim COBRA_NC  As Currency
Dim COBRA_LE   As Currency
Dim COBRA_EF2      As Currency
Dim COBRA_CH2   As Currency
Dim COBRA_NC2  As Currency
Dim COBRA_LE2   As Currency
Dim OTRO_INGRE As Currency

Dim wsFECHA1
Dim WDOLAR_CRED As Currency
Dim WSOLES_CRED As Currency
Dim TT_TOTAL_SOLES As Currency
Dim TT_TOTAL_DOLAR  As Currency
Dim WSOLES_DOLAR As Currency
Dim TT_DOLAR_SOLES As Currency
Dim WSOLES_DOLAR_CRED As Currency
Dim WSOLES_DOLAR_TOT As Currency
Dim SOLES_TOTAL_COBRA As Currency
Dim DOLAR_TOTAL_COBRA As Currency
Dim TOTAL_COBRA_SD As Currency
Dim WS_TIPO_C As Currency
Dim TOTAL_FACTURAS_UM As Currency
Dim TOTAL_FACTURAS_LITROS As Currency
Dim TOTAL_BOLETAS_UM As Currency
Dim TOTAL_BOLETAS_LITROS As Currency

If Right(txtFecha.Text, 2) = "__" Then
     wsFECHA1 = Left(txtFecha.Text, 8)
Else
     wsFECHA1 = Trim(txtFecha.Text)
End If
If Not IsDate(wsFECHA1) Then
 FrmImp2.Pantalla.Enabled = True
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If


FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL

pub_cadena = "SELECT SUT_SECUENCIA, SUT_DESCRIPCION FROM SUB_TRANSA WHERE SUT_SECUENCIA = ? AND SUT_CODTRA = 2401"
Set PS_REP04 = CN.CreateQuery("", pub_cadena)
PS_REP04(0) = 0
Set llave_rep04 = PS_REP04.OpenResultset(rdOpenKeyset, rdConcurValues)


pub_cadena = "SELECT FAR_FBG,(FAR_CANTIDAD/FAR_EQUIV) AS UM,(FAR_CANTIDAD/FAR_EQUIV)*FAR_LITRO AS LITROS " & _
" FROM FACART WHERE FAR_CODCIA=? AND " & _
"FAR_FECHA_COMPRA = ? AND FAR_MONEDA = ? AND FAR_TIPMOV = 10   AND FAR_ESTADO <> 'E' " & _
"ORDER BY FAR_FBG, FAR_NUMSER,  FAR_NUMFAC"

Set PS_FACART = CN.CreateQuery("", pub_cadena)
PS_FACART(0) = LK_CODCIA
PS_FACART(1) = LK_FECHA_DIA
PS_FACART(2) = Left(cmdmoneda.Text, 1)
Set llave_FACART = PS_FACART.OpenResultset(rdOpenKeyset, rdConcurValues)





f1 = 5
pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND " & _
"ALL_FECHA_SUNAT = ? AND ALL_MONEDA_CAJA = ? AND ALL_TIPMOV = 10   AND all_flag_ext <> 'E' AND " & _
"(ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY ALL_FBG, ALL_SIGNO_CAJA DESC, ALL_SECUENCIA, ALL_NUMSER,  ALL_NUMFAC"


Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = LK_FECHA_DIA
PS_REP01(3) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = ""
PS_REP01(2) = wsFECHA1
PS_REP01(3) = Left(cmdmoneda.Text, 1)
llave_rep01.Requery
If llave_rep01.EOF Then
  MsgBox "No Existe Ventas", 48, Pub_Titulo
  GoTo CANCELA
  Exit Sub
End If
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.max = llave_rep01.RowCount
f1 = 6
ww_signo = llave_rep01!ALL_SIGNO_CAJA
ww_sec = llave_rep01!all_SECUENCIA
ww_fbg = llave_rep01!ALL_FBG
FrmImp2.ProgBar.Value = 0
Do Until llave_rep01.EOF
'    FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
    If ww_fbg <> llave_rep01!ALL_FBG Then
        GoSub IMP_LINEA
        ww_fbg = llave_rep01!ALL_FBG
        ww_sec = llave_rep01!all_SECUENCIA
        ww_signo = llave_rep01!ALL_SIGNO_CAJA
    End If
    If ww_signo <> llave_rep01!ALL_SIGNO_CAJA Then
        GoSub IMP_LINEA
        ww_signo = llave_rep01!ALL_SIGNO_CAJA
        ww_sec = llave_rep01!all_SECUENCIA
    End If
    If ww_sec <> llave_rep01!all_SECUENCIA Then
        GoSub IMP_LINEA
        ww_sec = llave_rep01!all_SECUENCIA
    End If
    TT_TOTAL_SOLES = TT_TOTAL_SOLES + val(llave_rep01!ALL_IMPORTE_AMORT)
    If llave_rep01!ALL_SIGNO_CAJA = 0 Then
      WMONTO_CREDITO = WMONTO_CREDITO + val(llave_rep01!ALL_IMPORTE_AMORT)
      ACU_CREDITO = ACU_CREDITO + val(llave_rep01!ALL_IMPORTE_AMORT)
    Else
      WMONTO_CONTADO = WMONTO_CONTADO + val(llave_rep01!ALL_IMPORTE_AMORT)
      ACU_CONTADO = ACU_CONTADO + val(llave_rep01!ALL_IMPORTE_AMORT)
    End If
    llave_rep01.MoveNext
Loop
GoSub IMP_LINEA
f1 = f1 + 1
xl.Cells(f1, 1) = "TOTAL VENTA = "
xl.Cells(f1, 2) = ACU_CONTADO
xl.Cells(f1, 3) = ACU_CREDITO
wranF = "B" & f1 & ":C" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1

' SELECT PARA COBRANZAS EFECTIVO
pub_cadena = "SELECT ALL_CONCEPTO,ALL_FBG ,ALL_SIGNO_CAJA,ALL_TIPDOC, ALL_CODTRA, ALL_MONEDA_CAJA , ALL_FECHA_SUNAT, ALL_MONEDA_CLI, ALL_IMPORTE_AMORT, ALL_TIPO_CAMBIO  FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_MONEDA_CAJA = ? AND (ALL_FBG = ? OR ALL_FBG = 'N') AND (ALL_CODTRA = 2770 OR ALL_CODTRA = 2774 or all_codtra = 2412) AND ALL_TIPMOV <> 10 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = LK_FECHA_DIA
PS_REP01(3) = 0
PS_REP01(4) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = ""
PS_REP01(2) = wsFECHA1
PS_REP01(3) = Left(cmdmoneda.Text, 1)
PS_REP01(4) = "F"
llave_rep01.Requery
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Min = 0
If Not llave_rep01.EOF Then FrmImp2.ProgBar.max = llave_rep01.RowCount
FrmImp2.ProgBar.Value = 0
COBRA_EF = 0
COBRA_CH = 0
COBRA_NC = 0
COBRA_LE = 0
FrmImp2.ProgBar.Value = 0
Do Until llave_rep01.EOF
    FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
    WIMPORTE = llave_rep01!ALL_IMPORTE_AMORT
    If llave_rep01!ALL_SIGNO_CAJA = 1 Then
      COBRA_EF = COBRA_EF + WIMPORTE
    ElseIf llave_rep01!ALL_SIGNO_CAJA = 0 And llave_rep01!ALL_TIPDOC = "CH" Then
  '  F1 = F1 + 1
  '  xl.Cells(F1, 4) = WIMPORTE
      COBRA_CH = COBRA_CH + WIMPORTE
    ElseIf llave_rep01!ALL_SIGNO_CAJA = 0 And llave_rep01!ALL_CODTRA = 2412 And Left(Trim(llave_rep01!all_concepto), 1) = "F" Then
      COBRA_NC = COBRA_NC + WIMPORTE
    ElseIf llave_rep01!ALL_SIGNO_CAJA = 0 And llave_rep01!ALL_TIPDOC = "LE" Then
      COBRA_LE = COBRA_LE + WIMPORTE
    End If
    
llave_rep01.MoveNext
Loop

f1 = f1 + 2
xl.Cells(f1, 1) = "COBRANZA x FACTURAS EFECTIVO"
xl.Cells(f1, 4) = COBRA_EF
f1 = f1 + 1
xl.Cells(f1, 1) = "COBRANZA x FACTURAS CHEQUE"
xl.Cells(f1, 5) = COBRA_CH
f1 = f1 + 1
xl.Cells(f1, 1) = "COBRANZA x FACTURAS NOTA ABONO"
xl.Cells(f1, 5) = COBRA_NC
f1 = f1 + 1
xl.Cells(f1, 1) = "COBRANZA x FACTURAS LETRAS"
xl.Cells(f1, 5) = COBRA_LE

' BOLETAS
PS_REP01(4) = "B"
llave_rep01.Requery
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Min = 0
If Not llave_rep01.EOF Then FrmImp2.ProgBar.max = llave_rep01.RowCount
FrmImp2.ProgBar.Value = 0
COBRA_EF2 = 0
COBRA_CH2 = 0
COBRA_NC2 = 0
COBRA_LE2 = 0
FrmImp2.ProgBar.Value = 0
Do Until llave_rep01.EOF
    FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
    WIMPORTE = llave_rep01!ALL_IMPORTE_AMORT
    If llave_rep01!ALL_SIGNO_CAJA = 1 Then
      COBRA_EF2 = COBRA_EF2 + WIMPORTE
    ElseIf llave_rep01!ALL_SIGNO_CAJA = 0 And llave_rep01!ALL_TIPDOC = "CH" Then
      COBRA_CH2 = COBRA_CH2 + WIMPORTE
    ElseIf llave_rep01!ALL_SIGNO_CAJA = 0 And llave_rep01!ALL_CODTRA = 2412 And Left(Trim(llave_rep01!all_concepto), 1) = "B" Then
      COBRA_NC2 = COBRA_NC2 + WIMPORTE
    ElseIf llave_rep01!ALL_SIGNO_CAJA = 0 And llave_rep01!ALL_TIPDOC = "LE" Then
      COBRA_LE2 = COBRA_LE2 + WIMPORTE
    End If
llave_rep01.MoveNext
Loop



' SELECT OTROS INGRESOS
pub_cadena = "SELECT ALL_IMPORTE,ALL_CONCEPTO,ALL_FBG ,ALL_SIGNO_CAJA,ALL_TIPDOC, ALL_CODTRA, ALL_MONEDA_CAJA , ALL_FECHA_SUNAT, ALL_MONEDA_CLI, ALL_IMPORTE_AMORT, ALL_TIPO_CAMBIO  FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_FECHA_SUNAT = ? AND ALL_MONEDA_CAJA = ? AND (ALL_CODTRA = 5350) AND ALL_TIPMOV <> 10 AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_SUNAT"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = LK_FECHA_DIA
PS_REP01(3) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = ""
PS_REP01(2) = wsFECHA1
PS_REP01(3) = Left(cmdmoneda.Text, 1)
llave_rep01.Requery
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Min = 0
If Not llave_rep01.EOF Then FrmImp2.ProgBar.max = llave_rep01.RowCount
FrmImp2.ProgBar.Value = 0
OTRO_INGRE = 0
FrmImp2.ProgBar.Value = 0
Do Until llave_rep01.EOF
   OTRO_INGRE = OTRO_INGRE + val(llave_rep01!ALL_IMPORTE)
   llave_rep01.MoveNext
Loop


f1 = f1 + 1
xl.Cells(f1, 1) = "COBRANZA x BOLETAS EFECTIVO"
xl.Cells(f1, 4) = COBRA_EF2
f1 = f1 + 1
xl.Cells(f1, 1) = "COBRANZA x BOLETAS CHEQUE"
xl.Cells(f1, 5) = COBRA_CH2
f1 = f1 + 1
xl.Cells(f1, 1) = "COBRANZA x BOLETAS NOTA ABONO"
xl.Cells(f1, 5) = COBRA_NC2
f1 = f1 + 1
xl.Cells(f1, 1) = "COBRANZA x BOLETAS LETRAS"
xl.Cells(f1, 5) = COBRA_LE2
f1 = f1 + 1
xl.Cells(f1, 1) = "TOTAL COBRANZAS :"
xl.Cells(f1, 4) = COBRA_EF + COBRA_EF2
xl.Cells(f1, 5) = (COBRA_CH + COBRA_CH2) + (COBRA_NC + COBRA_NC2) + (COBRA_LE + COBRA_LE2)
wranF = "D" & f1 & ":E" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
f1 = f1 + 1
wranF = "A" & f1 & ":E" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1

f1 = f1 + 1
xl.Cells(f1, 1) = "TOTAL INGRESO EFECTIVO (+):"
xl.Cells(f1, 2) = (COBRA_EF + COBRA_EF2) + ACU_CONTADO
f1 = f1 + 1
xl.Cells(f1, 1) = "TOTAL INGRESOS CHEQUE  (+):"
xl.Cells(f1, 2) = (COBRA_CH + COBRA_CH2) + (COBRA_NC + COBRA_NC2)
f1 = f1 + 1
xl.Cells(f1, 1) = "TOTAL OTROS INGRESOS   (+):"
xl.Cells(f1, 2) = OTRO_INGRE
f1 = f1 + 1
xl.Cells(f1, 1) = "TOTAL GASTOS           (-):"
xl.Cells(f1, 2) = 0
f1 = f1 + 1
xl.Cells(f1, 1) = "TOTAL NOTA ABONO       (-):"
xl.Cells(f1, 2) = (COBRA_NC + COBRA_NC2)
f1 = f1 + 1
xl.Cells(f1, 1) = "TOTAL LETRAS           (-):"
xl.Cells(f1, 2) = (COBRA_LE + COBRA_LE2)
f1 = f1 + 1
xl.Cells(f1, 1) = "TOTAL A DEPOSITAR      (=):"
xl.Cells(f1, 2) = ((COBRA_EF + COBRA_EF2) + ACU_CONTADO) + (COBRA_CH + COBRA_CH2 + (COBRA_NC + COBRA_NC2)) - (COBRA_NC + COBRA_NC2) - (COBRA_LE + COBRA_LE2) + OTRO_INGRE

f1 = f1 + 1
xl.Cells(f1, 3) = "TOT. DEPOSITO :"
xl.Cells(f1, 4) = ((COBRA_EF + COBRA_EF2) + ACU_CONTADO) + (COBRA_CH + COBRA_CH2 + (COBRA_NC + COBRA_NC2)) - (COBRA_NC + COBRA_NC2) - (COBRA_LE + COBRA_LE2) + OTRO_INGRE
wranF = "A" & f1 & ":E" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
f1 = f1 + 1
wranF = "A" & f1 & ":E" & f1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1



'TOTALES LITROS Y UM
f1 = f1 + 1
PS_FACART(0) = LK_CODCIA
PS_FACART(1) = wsFECHA1
PS_FACART(2) = Left(cmdmoneda.Text, 1)
llave_FACART.Requery

TOTAL_FACTURAS_UM = 0
TOTAL_FACTURAS_LITROS = 0

TOTAL_BOLETAS_UM = 0
TOTAL_BOLETAS_LITROS = 0

While Not llave_FACART.EOF
 If llave_FACART!far_fbg = "F" Then
  TOTAL_FACTURAS_UM = TOTAL_FACTURAS_UM + llave_FACART!UM
  TOTAL_FACTURAS_LITROS = TOTAL_FACTURAS_LITROS + llave_FACART!LITROS
 ElseIf llave_FACART!far_fbg = "B" Then
  TOTAL_BOLETAS_UM = TOTAL_BOLETAS_UM + llave_FACART!UM
  TOTAL_BOLETAS_LITROS = TOTAL_BOLETAS_LITROS + llave_FACART!LITROS
  
 End If
 llave_FACART.MoveNext
Wend

xl.Cells(f1, 3) = "U.M. VENDIDAS:"
xl.Cells(f1 + 1, 2) = "FAC."
xl.Cells(f1 + 1, 3) = Format(TOTAL_FACTURAS_UM, "##,##0.000")
xl.Cells(f1 + 2, 2) = "BOL."
xl.Cells(f1 + 2, 3) = Format(TOTAL_BOLETAS_UM, "##,##0.000")
xl.Cells(f1 + 3, 2) = "TOTAL :"
xl.Cells(f1 + 3, 3) = Format(TOTAL_BOLETAS_UM + TOTAL_FACTURAS_UM, "##,##0.000")

xl.Cells(f1, 5) = "LITROS VENDIDOS:"
xl.Cells(f1 + 1, 4) = "FAC."
xl.Cells(f1 + 1, 5) = Format(TOTAL_FACTURAS_LITROS, "##,##0.000")
xl.Cells(f1 + 2, 4) = "BOL."
xl.Cells(f1 + 2, 5) = Format(TOTAL_BOLETAS_LITROS, "##,##0.000")
xl.Cells(f1 + 3, 4) = "TOTAL :"
xl.Cells(f1 + 3, 5) = Format(TOTAL_BOLETAS_LITROS + TOTAL_FACTURAS_LITROS, "##,##0.000")
wranF = "B" & f1 + 3 & ":E" & f1 + 3
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
DoEvents
FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
DoEvents
xcuenta = 1
xl.APPLICATION.Visible = True
xl.Cells(2, 2) = "CTA. CTE. " & UCase(Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia")))))
xl.Cells(4, 2) = "DEL DIA : " & Format(wsFECHA1, "dd mmm yyyy")
If Left(cmdmoneda.Text, 1) = "S" Then
 xl.Cells(3, 2) = "RESUMEN DE CAJA DE SOLES (S/.)"
Else
 xl.Cells(3, 2) = "RESUMEN DE CAJA DE DOLARES (US$.)"
End If
If checia.Value = 1 Then
  ' xl.Cells(2, 1) = "Almacenes 01 - 02 "
End If
xl.Cells(1, 1) = Now
xl.DisplayAlerts = False
xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
xl.Worksheets("Hoja1").Protect PUB_CLAVE
xl.APPLICATION.Visible = True
DoEvents
FrmImp2.lblproceso.Visible = False
FrmImp2.ProgBar.Visible = False
Set xl = Nothing
Screen.MousePointer = 0
FrmImp2.Pantalla.Enabled = True
FrmImp2.Pantalla.Caption = "Por &Pantalla"
FrmImp2.lblproceso.Visible = False

Exit Sub

IMP_LINEA:
f1 = f1 + 1
PS_REP04(0) = ww_sec
llave_rep04.Requery
ww_descrip = "VENTAS"
If Not llave_rep04.EOF Then
  ww_descrip = llave_rep04!sut_descripcion
End If
ww_descrip = UCase(ww_descrip)
If ww_fbg = "F" Then
  ww_muestra = "VENTA X FACTURAS " & ww_descrip
ElseIf ww_fbg = "B" Then
  ww_muestra = "VENTA X BOLETAS " & ww_descrip
ElseIf ww_fbg = "G" Then
  ww_muestra = "GUIAS"
Else
  ww_muestra = "PENDIENTES"
End If
xl.Cells(f1, 1) = ww_muestra & " :"
xl.Cells(f1, 2) = Format(WMONTO_CONTADO, "0.000")
xl.Cells(f1, 3) = Format(WMONTO_CREDITO, "0.000")
xl.Cells(f1, 4) = ""
xl.Cells(f1, 5) = ""
WMONTO_CONTADO = 0
WMONTO_CREDITO = 0
Return


IMP_LINEA2:
f1 = f1 + 1
If ww_fbg = "F" And ww_sec = 1 Then
  ww_muestra = "COBRANZA x FACTURAS EFECTIVO"
ElseIf ww_fbg = "F" And ww_sec = 2 Then
  ww_muestra = "COBRANZA x FACTURAS CHEQUE"
ElseIf ww_fbg = "F" And ww_sec = 3 Then
  ww_muestra = "COBRANZA x FACTURAS NOTA ABONO"
ElseIf ww_fbg = "F" And ww_sec = 4 Then
  ww_muestra = "COBRANZA x FACTURAS LETRAS"
ElseIf ww_fbg = "B" And ww_sec = 1 Then
  ww_muestra = "COBRANZA x BOLETAS EFECTIVO"
ElseIf ww_fbg = "B" And ww_sec = 2 Then
  ww_muestra = "COBRANZA x BOLETAS CHEQUE"
ElseIf ww_fbg = "B" And ww_sec = 3 Then
  ww_muestra = "COBRANZA x BOLETAS NOTA ABONO"
ElseIf ww_fbg = "B" And ww_sec = 4 Then
  ww_muestra = "COBRANZA x BOLETAS LETRAS"
End If
  
xl.Cells(f1, 1) = ww_muestra & " :"
xl.Cells(f1, 2) = Format(WMONTO_CONTADO, "0.000")
xl.Cells(f1, 3) = Format(WMONTO_CREDITO, "0.000")
xl.Cells(f1, 4) = ""
xl.Cells(f1, 5) = ""
WMONTO_CONTADO = 0
WMONTO_CREDITO = 0
Return


WEXCEL:
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FrmImp2.lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
  DoEvents
  WPAS = "131296"
  'WPAS = PUB_RUTA_OTRO + "CAJA_DET.xls"
  'DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open PUB_RUTA_OTRO & "\CAJA_DET.xls", 0, True, 4, WPAS, WPAS

  'xl.Workbooks.Open , "C:\ADMIN\HERTISA\CAJA_DET.XLS", 0, True, 4, PUB_CLAVE, PUB_CLAVE
Return

Exit Sub

'LLENA_VALOR:
'For I = QJ To WDIF
'  If I >= 3 Then
'    xl.Cells(F1 + 7, I) = Format(LOC_VALOR, "0.000")
'  End If
'Next I
'Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
Exit Sub

End Sub

Public Sub REG_VENTA_NORDI()
'On Error GoTo FINTODO
Dim TD_F As String * 1
Dim TD_B As String * 1
Dim TD_N As String * 1
Dim TD_D As String * 1
Dim w_exo As Currency
Dim ws_codcia As String * 2
Dim FILAS As Integer
Dim nn, m_ind As Integer
Dim AWQ_NETO_ACT_FIJO As Currency
Dim AWQ_CTA_ACT_FIJO As String
Dim AWQ_BRUTO_ACT_FIJO As Currency
Dim WCONTROL As Integer
FILAS = 5
Dim WQ
Dim wfecha
Dim ws_clave
Dim LETRAS(24) As String * 1
Dim wRuta As String
Dim wmonto As Currency
Dim wcodclie As Currency
Dim valor_venta As Currency
Dim WQ_TIPMOV As Integer
Dim descto As Currency
Dim valor_igv As Currency
Dim valor_precio As Currency
Dim S_VALOR_VENTA As Currency
Dim s_descto As Currency
Dim s_valor_igv As Currency
Dim S_VALOR_PRECIO As Currency

Dim t_valor_venta As Currency
Dim t_descto As Currency
Dim t_valor_igv As Currency
Dim t_valor_precio As Currency

Dim wnumfac As Currency
Dim WFLAG As String * 1
Dim wsFECHA1, wsFECHA2
Dim xcuenta As Integer
Dim wq_fecha, wq_codclie, wq_codven, wq_docu, wq_nombre, wq_bruto, wq_gastos, wq_desto, wq_flete, wq_fbg, wq_serie
Dim wq_tot_descto, WQ_IMPTO, WQ_TOTAL, wq_estado, wq_condi
Dim wq_ruc
Dim wflag_numfac
Dim WQ_BRUTO_D, WQ_GASTOS_D, WQ_FLETE_D, WQ_IMPTO_D, WQ_TOTAL_D As Currency
Dim wserie As String * 3
Dim AWQ_BRUTO As Currency
Dim AWQ_DESCTOS As Currency
Dim AWQ_GASTOS As Currency
Dim AWQ_FLETES As Currency
Dim AWQ_IMPTO As Currency
Dim AWQ_NETO As Currency
Dim AWQ_NETO_CRED  As Currency
Dim AWQ_NETO_CONT   As Currency
Dim AWQ_COSTO_VENTA As Currency
Dim WS_SIGNO As Integer
Dim ws_tc As Currency
Dim wsTexto As String
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) <> cop_llave!cop_fecha_proceso Then cheasiento.Value = 0
If CDate(wsFECHA2) <> cop_llave!cop_fecha_proceso2 Then cheasiento.Value = 0
GoSub WEXCEL
If LK_EMP = "PIU" Then
  max.Text = 90000
End If
pub_cadena = ""
xcuenta = 0
TD_F = ""
TD_B = ""
TD_N = ""
TD_D = ""
wsTexto = "TIPO: "
For fila = 0 To 3
 txttp.ListIndex = fila
 If fila = 0 And txttp.Selected(fila) Then
   TD_F = "F"
   wsTexto = wsTexto + "- FACT."
 End If
 If fila = 1 And txttp.Selected(fila) Then
   TD_B = "B"
   wsTexto = wsTexto + "- BOLE."
 End If
 If fila = 2 And txttp.Selected(fila) Then
   TD_N = "N"
   wsTexto = wsTexto + "- NCRE."
 End If
 If fila = 3 And txttp.Selected(fila) Then
   TD_D = "D"
   wsTexto = wsTexto + "- NDEB."
 End If
Next fila

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
WCONTROL = 0
f1 = 5  'Fila Inicial
t_valor_venta = 0
t_descto = 0
t_valor_igv = 0
t_valor_precio = 0

'NCREDITO: ' empieza

'WCONTROL = WCONTROL + 1
wmonto = 0
wcodclie = 0
valor_venta = 0
descto = 0
valor_igv = 0
valor_precio = 0
wnumfac = 0
WFLAG = ""
xcuenta = 0
wflag_numfac = ""
wserie = ""
AWQ_BRUTO = 0
AWQ_DESCTOS = 0
AWQ_GASTOS = 0
AWQ_FLETES = 0
AWQ_IMPTO = 0
AWQ_NETO = 0
AWQ_NETO_CRED = 0
AWQ_NETO_CONT = 0
AWQ_COSTO_VENTA = 0
AWQ_NETO_ACT_FIJO = 0
AWQ_BRUTO_ACT_FIJO = 0



f1 = 5
FILAS = 0
FILAS = FILAS + 1
f1 = f1 + 1

NCREDITO:


S_VALOR_VENTA = 0
s_descto = 0
s_valor_igv = 0
S_VALOR_PRECIO = 0

WCONTROL = WCONTROL + 1

If WCONTROL = 1 Then
  If TD_F = "F" Or TD_B = "B" Then
  Else
    GoTo OTRO_DOCUMENTO
  End If
  pub_cadena = "SELECT * FROM  FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND FAR_TIPMOV = 10 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = ? OR FAR_FBG = ? ) ORDER BY  FAR_FECHA_COMPRA, FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC,FAR_NUMSEC"
  WS_SIGNO = 1
ElseIf WCONTROL = 2 Then
  If TD_N = "N" Then
  Else
    GoTo OTRO_DOCUMENTO
  End If
  pub_cadena = "SELECT *  FROM FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND FAR_TIPMOV = 97 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?  AND FAR_CP = 'C'  ORDER BY FAR_TIPMOV, FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC " ' FAR_TIPMOV, FAR_FBG DESC , FAR_FECHA_COMPRA,  FAR_NUMSER,FAR_NUMFAC"
  WS_SIGNO = -1
ElseIf WCONTROL = 3 Then
  If TD_D = "D" Then
  Else
    GoTo OTRO_DOCUMENTO
  End If
  pub_cadena = "SELECT *  FROM FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? OR FAR_CODCIA= ? OR FAR_CODCIA= ? ) AND FAR_TIPMOV = 98 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?   AND FAR_CP = 'C' ORDER BY FAR_TIPMOV, FAR_FBG DESC , FAR_FECHA_COMPRA,  FAR_NUMSER,FAR_NUMFAC"
  WS_SIGNO = 1
End If

If Trim(txtSerie.Text) <> "" Then
If WCONTROL = 1 Then
  pub_cadena = "SELECT * FROM  FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND FAR_TIPMOV = 10 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = ? OR FAR_FBG = ? )AND FAR_NUMSER = ?  ORDER BY FAR_TIPMOV, FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC"
  WS_SIGNO = 1
End If
End If

If Trim(fbtxt.Text) <> "" And Trim(txtSerie.Text) <> "" Then
If WCONTROL = 1 Then
  pub_cadena = "SELECT * FROM  FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND FAR_TIPMOV = 10 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = ? OR FAR_FBG = ? )AND FAR_NUMSER = ?  ORDER BY FAR_TIPMOV, FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC"
  WS_SIGNO = 1
End If
End If

If Right(max.Text, 1) = "D" Or Right(max.Text, 1) = "S" Then
If WCONTROL = 1 Then
  pub_cadena = "SELECT * FROM  FACART WHERE ( FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA = ? OR FAR_CODCIA= ? )AND FAR_TIPMOV = 10 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = ? OR FAR_FBG = ? )AND FAR_MONEDA = ?  ORDER BY FAR_TIPMOV, FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC"
  WS_SIGNO = 1
End If
End If




Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = ""
PS_REP02(1) = ""
PS_REP02(2) = ""
PS_REP02(3) = ""
PS_REP02(4) = ""
PS_REP02(5) = wsFECHA1
PS_REP02(6) = wsFECHA2
If WCONTROL = 1 Then
 PS_REP02(7) = TD_F
 PS_REP02(8) = TD_B
End If


Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

PS_REP02(0) = ""
PS_REP02(1) = ""
PS_REP02(2) = ""
PS_REP02(3) = ""
PS_REP02(4) = ""
If checia.Visible And checia.Value = 1 Then
   If Trim(par_llave!par_art_cias) <> "" Then
     nn = 1
     For m_ind = 1 To 15
         ws_codcia = Mid(par_llave!par_art_cias, nn, 2)
         If Trim(ws_codcia) = "" Then Exit For
         PS_REP02(m_ind - 1) = ws_codcia
         nn = nn + 2
     Next m_ind
   End If
Else
PS_REP02(0) = LK_CODCIA
End If


' el PS_REP1(0) ESTA MAS ABAJO
PS_REP02(5) = wsFECHA1
PS_REP02(6) = wsFECHA2

If WCONTROL = 1 Then
PS_REP02(7) = TD_F
PS_REP02(8) = TD_B
''If fbtxt.Text = "F" Or fbtxt.Text = "B" Then
''   PS_REP02(7) = fbtxt.Text
''   PS_REP02(8) = fbtxt.Text
''End If
End If

If Trim(txtSerie.Text) <> "" Then
   PS_REP02(9) = txtSerie.Text
End If
If Right(max.Text, 1) = "S" Or Right(max.Text, 1) = "D" Then
   PS_REP02(9) = Right(max.Text, 1)
End If


DoEvents
FrmImp2.lblproceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep02.Requery
If llave_rep02.EOF Then
  GoTo OTRO_DOCUMENTO
End If
FrmImp2.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.max = llave_rep02.RowCount

WFLAG = ""
SQ_OPER = 1
pu_cp = "C"
pu_codcia = llave_rep02!FAR_CODCIA
wnumfac = llave_rep02!far_numfac
wfecha = llave_rep02!far_fbg 'llave_rep02!far_fecha
wserie = llave_rep02!far_numser
wq_fecha = llave_rep02!FAR_fecha_compra
WQ_TIPMOV = llave_rep02!FAR_TIPMOV
wq_fbg = Trim(llave_rep02!far_fbg)

xcuenta = 0
WFLAG = "A"
wflag_numfac = "A"
AWQ_DESCTOS = 0
AWQ_GASTOS = 0
AWQ_FLETES = 0
w_exo = 0
If llave_rep02.EOF Then GoTo CANCELA
Do Until llave_rep02.EOF
  If Trim(txtSerie.Text) <> "" Then
      If Trim(txtSerie.Text) <> Trim(llave_rep02!far_numser) Then GoTo SALTARIN
  End If
  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  
  If wnumfac = val(llave_rep02!far_numfac) And val(wserie) = val(llave_rep02!far_numser) And wq_fbg = llave_rep02!far_fbg And WQ_TIPMOV = llave_rep02!FAR_TIPMOV Then
  Else
     GoSub IMPRI_FAC
     wflag_numfac = ""
  End If
  wnumfac = llave_rep02!far_numfac
  wq_fecha = Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy")
  wq_codclie = llave_rep02!far_codclie
  wq_codven = llave_rep02!FAR_CODVEN
  wq_fbg = Trim(llave_rep02!far_fbg)
  WQ_TIPMOV = llave_rep02!FAR_TIPMOV
  wq_serie = "'" & llave_rep02!far_numser
  wserie = llave_rep02!far_numser
  wq_docu = "'" & llave_rep02!far_numfac
  wq_nombre = ""
  ws_tc = 1
  If llave_rep02!FAR_MONEDA = "D" Then
    ws_tc = JALAR(llave_rep02!FAR_fecha_compra)
    If ws_tc <= 0 Then
        MsgBox "Falta Ingresar el Tipo de Cambio del día : " & Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy"), 48, Pub_Titulo
        GoTo CANCELA
    End If
    If llave_rep02!far_estado <> "E" Then
       WQ_BRUTO_D = Format((val(llave_rep02!FAR_BRUTO) - val(llave_rep02!FAR_TOT_DESCTO)) * WS_SIGNO, "0.000")
       WQ_GASTOS_D = val(llave_rep02!FAR_GASTOS) * WS_SIGNO
       WQ_FLETE_D = val(Nulo_Valor0(llave_rep02!FAR_TOT_FLETE)) * WS_SIGNO
       WQ_IMPTO_D = Format(val(llave_rep02!far_IMPTO) * WS_SIGNO, "0.000")
       WQ_TOTAL_D = val(WQ_BRUTO_D) + val(WQ_IMPTO_D)
    End If
  End If
  
  wq_bruto = Format((val(llave_rep02!FAR_BRUTO) - val(llave_rep02!FAR_TOT_DESCTO)) * ws_tc * WS_SIGNO, "0.000")
  wq_gastos = val(llave_rep02!FAR_GASTOS) * WS_SIGNO * ws_tc
  wq_flete = val(Nulo_Valor0(llave_rep02!FAR_TOT_FLETE)) * WS_SIGNO * ws_tc
  WQ_IMPTO = Format(val(llave_rep02!far_IMPTO) * WS_SIGNO * ws_tc, "0.000")
  WQ_TOTAL = val(wq_bruto) + val(WQ_IMPTO)  ' (Val(llave_rep02!far_bruto) + Val(llave_rep02!far_impto) - Val(llave_rep02!FAR_TOT_DESCTO) + Val(llave_rep02!FAR_GASTOS)) * WS_SIGNO * WS_TC
  If llave_rep02!FAR_EX_IGV = "A" Then
     w_exo = w_exo + val(llave_rep02!FAR_SUBTOTAL)
  End If
  If llave_rep02!far_estado = "E" Then
     wq_bruto = 0
     wq_gastos = 0
     wq_flete = 0
     WQ_IMPTO = 0
     WQ_TOTAL = 0
     w_exo = 0
     WQ_BRUTO_D = 0
     WQ_GASTOS_D = 0
     WQ_FLETE_D = 0
     WQ_IMPTO_D = 0
     WQ_TOTAL_D = 0
 End If
  
 If wq_bruto = 0 Then WQ_TOTAL = 0
 wq_estado = llave_rep02!far_estado
 If wq_estado <> "E" Then
    If Left(UCase(llave_rep02!far_subtra), 1) <> "A" Then
     AWQ_COSTO_VENTA = AWQ_COSTO_VENTA + ((llave_rep02!FAR_COSPRO * llave_rep02!far_cantidad)) * WS_SIGNO
    End If
 End If
 wflag_numfac = "A"
 WFLAG = "A"
 pu_codcia = llave_rep02!FAR_CODCIA
SALTARIN:
 llave_rep02.MoveNext
Loop
If wflag_numfac = "A" Then
    GoSub IMPRI_FAC
    wflag_numfac = ""
End If
If WFLAG = "A" Then
    f1 = f1 + 1
    FILAS = FILAS + 1
    GoSub TOTAL_DIA
End If
  xcuenta = c1 + 1
  If WCONTROL = 1 Then
   If cheasiento.Value = 1 Then
   End If
  End If
OTRO_DOCUMENTO:
If WCONTROL >= 3 Or Trim(fbtxt.Text) <> "" Or Trim(txtSerie.Text) <> "" Or Right(max.Text, 1) = "S" Or Right(max.Text, 1) = "D" Then
Else
  GoTo NCREDITO
End If

MOSTRAR:
   If cheasiento.Value = 1 Then
    FrmImp2.lblproceso.Caption = "Procesando Pase de Contabilidad . . . "
    DoEvents
    GoSub PASE_CONTAB
   End If


  f1 = f1 + 2
  xl.Cells(f1, 1) = "Total General = "
  xl.Worksheets(1).Rows(f1).RowHeight = 20
  xl.Cells(f1, 8) = t_valor_venta
  xl.Cells(f1, 9) = t_descto
  xl.Cells(f1, 11) = t_valor_igv
  xl.Cells(f1, 12) = t_valor_precio
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  If checia.Visible And checia.Value = 1 Then
    xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  Else
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  End If
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'" & wsTexto & " DEL " & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub

IMPRI_FAC:
       FILAS = FILAS + 1
       f1 = f1 + 1
       pu_codclie = wq_codclie
       LEER_CLI_LLAVE
       If Not cli_llave.EOF Then
         wq_codclie = cli_llave!cli_codclie 'xl.Cells(F1, 2)
         wq_nombre = Trim(cli_llave!CLI_NOMBRE)
         wq_ruc = Trim(cli_llave!cli_ruc_esposo)
       End If
     xl.Cells(f1, 1) = "'" & wq_fecha
     xl.Cells(f1, 3) = wq_ruc
     xl.Cells(f1, 2) = wq_nombre
     If wq_fbg = "F" Then
        wq_condi = "01"
     ElseIf wq_fbg = "B" Then
        wq_condi = "03"
     ElseIf wq_fbg = "N" Then
        wq_condi = "07"
     ElseIf wq_fbg = "D" Then
        wq_condi = "08"
     End If
     xl.Cells(f1, 4) = "'" & wq_condi
     xl.Cells(f1, 5) = wq_fbg
     xl.Cells(f1, 6) = wq_serie
     xl.Cells(f1, 7) = wq_docu
     If wq_estado = "E" Then
         xl.Cells(f1, 2) = "[ANULADO] " & wq_nombre
     Else
         xl.Cells(f1, 2) = wq_nombre
     End If
     If wq_estado <> "E" Then
      If True Then 'If Left(cli_llave!CLI_CUENTA_CONTAB, 2) <> "12" Then
       xl.Cells(f1, 8) = Format(val(wq_bruto) - w_exo, "0.000")
       xl.Cells(f1, 9) = val(w_exo)
       xl.Cells(f1, 11) = val(WQ_IMPTO)
       xl.Cells(f1, 12) = val(WQ_TOTAL)
      Else
       xl.Cells(f1, 8) = Format(val(wq_bruto) - w_exo, "0.000")
       xl.Cells(f1, 9) = val(w_exo)
       xl.Cells(f1, 11) = val(WQ_IMPTO)
       xl.Cells(f1, 12) = val(WQ_TOTAL)
       xl.Cells(f1, 15) = Trim(cli_llave!CLI_CUENTA_CONTAB)
       AWQ_IMPTO = AWQ_IMPTO + val(WQ_IMPTO)
       AWQ_NETO = AWQ_NETO + val(WQ_TOTAL)
       AWQ_BRUTO = AWQ_BRUTO + wq_bruto
       AWQ_DESCTOS = wq_desto
       AWQ_GASTOS = wq_gastos
       AWQ_FLETES = wq_flete
       If wq_condi = "CRED" Then
         xl.Cells(f1, 16) = -1
         AWQ_NETO_CRED = AWQ_NETO_CRED + val(WQ_TOTAL)
       Else
         xl.Cells(f1, 16) = 0
         AWQ_NETO_CONT = AWQ_NETO_CONT + val(WQ_TOTAL)
       End If
     End If
     End If
     If ws_tc > 1 And Right(max.Text, 1) = "D" Then
       xl.Cells(f1, 17) = Format(val(WQ_BRUTO_D) - w_exo, "0.000")
       xl.Cells(f1, 18) = val(WQ_IMPTO_D)
       xl.Cells(f1, 19) = val(WQ_TOTAL_D)
       xl.Cells(f1, 20) = " " & ws_tc
    End If
     
     
     S_VALOR_VENTA = S_VALOR_VENTA + (wq_bruto - w_exo)
     s_descto = s_descto + val(w_exo)
     s_valor_igv = s_valor_igv + val(WQ_IMPTO)
     S_VALOR_PRECIO = S_VALOR_PRECIO + val(WQ_TOTAL)
     
     
     
     t_valor_venta = t_valor_venta + (wq_bruto - w_exo)
     t_descto = t_descto + val(w_exo)
     t_valor_igv = t_valor_igv + val(WQ_IMPTO)
     t_valor_precio = t_valor_precio + val(WQ_TOTAL)
     
     If FILAS >= val(max.Text) Then
        f1 = f1 + 1
        xl.Cells(f1, 1) = "VAN ... "
        xl.Worksheets(1).Rows(f1).RowHeight = 20
        xl.Cells(f1, 8) = t_valor_venta
        xl.Cells(f1, 9) = t_descto
        xl.Cells(f1, 11) = t_valor_igv
        xl.Cells(f1, 12) = t_valor_precio
        'F1 = F1 + 1
        wranF = "A" & f1
        xl.APPLICATION.Range(wranF).Select
        On Error Resume Next
        xl.APPLICATION.ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
        FILAS = 1
        f1 = f1 + 1
        xl.Cells(f1, 1) = "VIENEN ... "
        xl.Worksheets(1).Rows(f1).RowHeight = 20
        xl.Cells(f1, 8) = t_valor_venta
        xl.Cells(f1, 9) = t_descto
        xl.Cells(f1, 11) = t_valor_igv
        xl.Cells(f1, 12) = t_valor_precio
     End If

Return

CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\REGVENTA.xls", 0, True, 4, WPAS, WPAS
Return

TOTAL_DIA:
  If wfecha = "F" Then
    xl.Cells(f1, 1) = "Total Ventas      = "
  ElseIf wfecha = "B" Then
    xl.Cells(f1, 1) = "Total Boletas     = "
  ElseIf wfecha = "N" Then
    xl.Cells(f1, 1) = "Total N.Creditos  = "
  ElseIf wfecha = "D" Then
    xl.Cells(f1, 1) = "Total N.Debito    = "
  End If
  xl.Worksheets(1).Rows(f1).RowHeight = 20
  xl.Cells(f1, 2) = ""
  xl.Cells(f1, 3) = ""
  xl.Cells(f1, 7) = ""
  xl.Cells(f1, 8) = S_VALOR_VENTA
  xl.Cells(f1, 9) = s_descto
  xl.Cells(f1, 11) = s_valor_igv
  xl.Cells(f1, 12) = S_VALOR_PRECIO
Return

PASE_CONTAB:
Dim wcta As String
Dim wcta_clientes As Currency
Dim PS_CONTAB1 As rdoQuery
Dim contab_llave As rdoResultset
Dim ws_nro_voucher As Integer
Dim ws_nro_sec As Integer
Dim ws_glosa As String
Dim wsq_fecha As String
Dim wdh As String * 1
Dim wscodcia As String * 2
Dim wsq_fecha2
wscodcia = LK_CODCIA
ws_glosa = "Registro de Venta"
If LK_EMP_PTO = "A" Then
 wscodcia = "00"
 ws_glosa = "Registro de Venta - " & Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
End If
wsq_fecha = Format(cop_llave!cop_fecha_proceso, "yyyy/mm/dd")
wsq_fecha2 = Format(cop_llave!cop_fecha_proceso2, "yyyy/mm/dd")
pub_cadena = "DELETE COMOV  WHERE  COV_FLAG_AUTOMATICA = '3' AND COV_CODUSU = '" & LK_CODCIA & "' AND (COV_FECHA_VOUCHER >=  ' " & wsq_fecha & "' AND COV_FECHA_VOUCHER <=  ' " & wsq_fecha2 & "')"
CN.Execute pub_cadena, rdExecDirect

FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.max = 10
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
PSCOV_VOUCHER(0) = wscodcia
PSCOV_VOUCHER(1) = cop_llave!cop_fecha_proceso
PSCOV_VOUCHER(2) = cop_llave!cop_fecha_proceso2
cov_voucher.Requery
If cov_voucher.EOF Then
 ws_nro_voucher = 0
Else
 cov_voucher.MoveLast
 ws_nro_voucher = cov_voucher!COV_NRO_VOUCHER
End If
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
wran1 = "A" & 6 & ":Q" & f1
xl.APPLICATION.Worksheets("Hoja1").Range(wran1).Sort Key1:=xl.APPLICATION.Worksheets("Hoja1").Range("N6")
fila = 6
' xl.Application.Visible = True
wcta = Trim(xl.Cells(fila, 14))
wcta_clientes = 0
wdh = "D"
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
For fila = 6 To f1
  If val(xl.Cells(fila, 15)) = 0 Then GoTo OTRITO
  'If Trim(xl.Cells(fila, 5)) = "N" Then
  '   GoTo OTRITO
  'End If
  If val(xl.Cells(fila, 14)) = 0 Then GoTo OTRITO
   If wcta <> Trim(xl.Cells(fila, 14)) Then
     GoSub Graba
     wcta_clientes = 0
     wcta = Trim(xl.Cells(fila, 14))
     wcta_clientes = wcta_clientes + val(xl.Cells(fila, 11))
  Else
     wcta_clientes = wcta_clientes + val(xl.Cells(fila, 11))
  End If
OTRITO:
Next fila
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
GoSub Graba
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1

If AWQ_NETO_ACT_FIJO <> 0 Then
 SQ_OPER = 1
 PUB_SECUENCIA = 2
 PUB_CODTRA = 2401
 PUB_CODCIA = wscodcia
 LEER_CNT_LLAVE
 If cnt_llave.EOF Then
  MsgBox "Error de Dato de Transaccion , Consulte a su Proveedor.", 48, Pub_Titulo
  End
 End If
 If Trim(cnt_llave!CNT_CTA1) <> "" Then
  wcta = cnt_llave!CNT_CTA1  'Ventas a costo
  wdh = cnt_llave!CNT_DH1
  wcta_clientes = AWQ_BRUTO_ACT_FIJO
  GoSub Graba
 End If
  wcta = AWQ_CTA_ACT_FIJO
  wdh = "D"
  wcta_clientes = AWQ_NETO_ACT_FIJO
  GoSub Graba
End If

SQ_OPER = 1
PUB_SECUENCIA = 24
PUB_CODTRA = 2401
PUB_CODCIA = wscodcia
LEER_CNT_LLAVE
If cnt_llave.EOF Then
 MsgBox "Error de Dato de Transaccion , Consulte a su Proveedor.", 48, Pub_Titulo
 'End
  GoTo CANCELA
End If
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1

If Trim(cnt_llave!CNT_CTA1) <> "" Then
 wcta = cnt_llave!CNT_CTA1 'Ventas Brutas
 wdh = cnt_llave!CNT_DH1
 wcta_clientes = AWQ_BRUTO
 GoSub Graba
End If
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
If Trim(cnt_llave!CNT_CTA2) <> "" Then
 wcta = cnt_llave!CNT_CTA2 'Impuesto
 wdh = cnt_llave!CNT_DH2
 wcta_clientes = AWQ_IMPTO
 GoSub Graba
End If

If AWQ_NETO_CONT <> 0 Then
 SQ_OPER = 1
 PUB_SECUENCIA = 24
 PUB_CODTRA = 2401
 PUB_CODCIA = wscodcia
 LEER_CNT_LLAVE
 If cnt_llave.EOF Then
  MsgBox "Error de Dato de Transaccion , Consulte a su Proveedor.", 48, Pub_Titulo
  End
 End If
 If Trim(cnt_llave!CNT_CTA1) <> "" Then
  wcta = cnt_llave!CNT_CTA1  'Ventas a costo
  wdh = "D" ' ES HABER PERO AL DEBE
  wcta_clientes = AWQ_NETO_CONT
  GoSub Graba
 End If
End If
SQ_OPER = 1
PUB_SECUENCIA = 24
PUB_CODTRA = 2401
PUB_CODCIA = wscodcia
LEER_CNT_LLAVE
If cnt_llave.EOF Then
 MsgBox "Error de Dato de Transaccion , Consulte a su Proveedor.", 48, Pub_Titulo
 End
End If
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
ws_nro_voucher = ws_nro_voucher + 1
If Trim(cnt_llave!CNT_CTA3) <> "" Then
 wcta = cnt_llave!CNT_CTA3  'Ventas a costo
 wdh = cnt_llave!CNT_DH3
 wcta_clientes = AWQ_COSTO_VENTA
 GoSub Graba
End If
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
If Trim(cnt_llave!CNT_CTA4) <> "" Then
 wcta = cnt_llave!CNT_CTA4  'Ventas a costo
 wdh = cnt_llave!CNT_DH4
 wcta_clientes = AWQ_COSTO_VENTA
 GoSub Graba
End If

FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1

Return

Graba:
     If wcta_clientes = 0 Then Return
     ws_nro_sec = ws_nro_sec + 1
     cov_voucher.AddNew
     cov_voucher!COV_CODCIA = wscodcia
     cov_voucher!COV_FECHA_VOUCHER = cop_llave!cop_fecha_proceso2
     cov_voucher!COV_NRO_MOV = ws_nro_sec
     cov_voucher!COV_NRO_VOUCHER = ws_nro_voucher
     cov_voucher!COV_NUMTAB = 0
     cov_voucher!COV_CODCTA = wcta
     cov_voucher!COV_DH = wdh
     cov_voucher!COV_IMPORTE = wcta_clientes
     cov_voucher!COV_ESTADO = " "
     cov_voucher!COV_CODUSU = LK_CODCIA
     cov_voucher!cov_flag_automatica = "3"
     cov_voucher!COV_glosa = ws_glosa
     cov_voucher!COV_FECHA_doc = LK_FECHA_DIA
     cov_voucher.Update
Return


LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"

Return

FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
 GoTo CANCELA
 Resume Next
End Sub

Public Sub LISTA_PRECIOS()
'On Error GoTo FINTODO
Dim nroitem As Integer
Dim CADENITA2 As String
Dim wSTOCK As Currency
Dim WSCODART As Currency
Dim ws_clave As String
Dim LETRAS(24) As String * 1
Dim WSFECHA As Date
Dim acu_val_ingresos As Currency
Dim acu_val_salidas As Currency
Dim acu_cant_dia As Currency
Dim acu_saldo As Currency
Dim acu_stock As Currency
Dim wsfile As String
Dim walterno As String
Dim wdnombre As String
Dim WD_COSPRO As Currency
Dim CADENITA  As String
Dim wwfami As Integer
If val(Right(listac.Text, 6)) = 0 Then
  MsgBox "Seleccionar Lista de Descto", 48, Pub_Titulo
  Exit Sub
End If
If val(Right(listac.Text, 6)) = 0 Then
  MsgBox "Seleccionar Tipo de Descto", 48, Pub_Titulo
  Exit Sub
End If

'If Val(Right(listVen.Text, 6)) = 0 Then
'  GoTo CANCELA
'End If
wsfile = ""
Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
CADENITA2 = ""
For fila = 0 To listacal.ListCount - 1
 listacal.ListIndex = fila
 If listacal.Selected(fila) Then
   CADENITA2 = CADENITA2 + "ART_CALIDAD = " & Trim(Right(listacal.Text, 8)) & " OR "
 End If
Next fila
If CADENITA2 <> "" Then
  CADENITA2 = "(" & Mid(CADENITA2, 1, Len(CADENITA2) - 4) & ")"
End If

CADENITA = ""
For fila = 0 To vendmulti.ListCount - 1
 vendmulti.ListIndex = fila
 If vendmulti.Selected(fila) Then
   CADENITA = CADENITA + "ART_FAMILIA = " & Trim(Right(vendmulti.Text, 8)) & " OR "
 End If
Next fila
If CADENITA <> "" Then
  CADENITA = "(" & Mid(CADENITA, 1, Len(CADENITA) - 4) & ")"
End If

If CADENITA <> "" Then
  If CADENITA2 <> "" Then
    pub_cadena = "SELECT ART_MONEDA, ART_FAMILIA, ART_NUMERO, ART_KEY, ART_ALTERNO, ART_NOMBRE, ARM_STOCK, ARM_COSPRO, ARM_SALDO_S FROM ARTI, ARTICULO  WHERE (ARM_CODART = ART_KEY) AND (ARM_CODCIA = ART_CODCIA) AND ART_CODCIA = ? AND ART_KEY <> 0 AND " & CADENITA & " AND " & CADENITA2 & " AND ART_CALIDAD = 1 ORDER BY ART_FAMILIA, ART_SUBFAM, ART_NUMERO, ART_NOMBRE "
  Else
    pub_cadena = "SELECT ART_MONEDA, ART_FAMILIA, ART_NUMERO, ART_KEY, ART_ALTERNO, ART_NOMBRE, ARM_STOCK, ARM_COSPRO, ARM_SALDO_S FROM ARTI, ARTICULO  WHERE (ARM_CODART = ART_KEY) AND (ARM_CODCIA = ART_CODCIA) AND ART_CODCIA = ? AND ART_KEY <> 0 AND " & CADENITA & " AND ART_CALIDAD = 1 ORDER BY ART_FAMILIA, ART_SUBFAM, ART_NUMERO,  ART_NOMBRE "
  End If
Else
  If CADENITA2 <> "" Then
    pub_cadena = "SELECT ART_MONEDA, ART_FAMILIA, ART_NUMERO, ART_KEY, ART_ALTERNO, ART_NOMBRE, ARM_STOCK, ARM_COSPRO, ARM_SALDO_S FROM ARTI, ARTICULO  WHERE (ARM_CODART = ART_KEY) AND (ARM_CODCIA = ART_CODCIA) AND ART_CODCIA = ? AND ART_KEY <> 0 AND " & CADENITA2 & " AND ART_CALIDAD = 1 ORDER BY ART_FAMILIA, ART_SUBFAM ART_NUMERO, ART_NOMBRE "
  Else
    pub_cadena = "SELECT ART_MONEDA, ART_FAMILIA, ART_NUMERO, ART_KEY, ART_ALTERNO, ART_NOMBRE, ARM_STOCK, ARM_COSPRO, ARM_SALDO_S FROM ARTI, ARTICULO  WHERE (ARM_CODART = ART_KEY) AND (ARM_CODCIA = ART_CODCIA) AND ART_CODCIA = ? AND ART_KEY <> 0 AND ART_CALIDAD = 1  ORDER BY ART_FAMILIA, ART_SUBFAM, ART_NUMERO, ART_NOMBRE "
  End If
End If

'PRUEBA X ARTI
' pub_cadena = "SELECT ART_FAMILIA, ART_KEY, ART_ALTERNO, ART_NOMBRE, ARM_STOCK, ARM_COSPRO, ARM_SALDO_S FROM ARTI, ARTICULO  WHERE (ARM_CODART = ART_KEY) AND (ARM_CODCIA = ART_CODCIA) AND ART_CODCIA = ? AND ART_KEY = 4991 ORDER BY ART_FAMILIA, ART_NOMBRE "

Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
'Debug.Print pub_cadena
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

'pub_cadena = "SELECT FAR_FECHA_COMPRA, FAR_CANTIDAD, FAR_SIGNO_ARM, FAR_COSPRO, FAR_CODART FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA_COMPRA >= ?  AND FAR_CODART = ? AND FAR_ESTADO <> 'E' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC , FAR_NUMOPER2"

pub_cadena = "SELECT * FROM CLIDSCTO WHERE CLD_CODCIA = ? AND CLD_TIPODSCTO = ? AND CLD_CODART = ? AND CLD_LISTADSCTO = ? "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = 0
PS_REP01(3) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
DoEvents
ws_clave = PUB_CLAVE
GoSub WEXCEL
FrmImp2.ProgBar.Visible = True
DoEvents
'xl.Worksheets(1).Activate
'GoSub LETRAS

xcuenta = 0
xl.Cells(1, 1) = Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia")))
xl.Cells(2, 1) = "'LISTADO DE PRECIOS CON DESCTO."
xl.Cells(3, 1) = "Lista : " & Left(listac.Text, 20)
xl.Cells(4, 1) = "Tipo  : " & Left(listat.Text, 20)


f1 = 6 'Fila Inicial
PS_REP02(0) = LK_CODCIA
llave_rep02.Requery
If llave_rep02.RowCount <> 0 Then
 FrmImp2.ProgBar.Min = 0
 FrmImp2.ProgBar.Value = 0
 FrmImp2.ProgBar.max = llave_rep02.RowCount
End If

FrmImp2.lblproceso.Visible = True
FrmImp2.lblproceso.Caption = "Procesando . . . "
DoEvents

acu_val_ingresos = 0
acu_val_salidas = 0
acu_cant_dia = 0
wSTOCK = 0
WD_COSPRO = 0
acu_saldo = 0
wwfami = -1
Do Until llave_rep02.EOF
        If llave_rep02!ART_KEY = 0 Then GoTo SIGUER
         FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
         If wwfami <> llave_rep02!art_numero Then
           f1 = f1 + 1
           xl.Cells(f1, 1) = ""
           SQ_OPER = 1
           PUB_CODCIA = LK_CODCIA
           PUB_TIPREG = 130
           PUB_NUMTAB = llave_rep02!art_numero
           LEER_TAB_LLAVE
           If tab_llave.EOF Then
        '       MsgBox "El Producto : " & llave_rep02!art_alterno & " " & llave_rep02!art_nombre & "  Definir Linea nuevamente", 48, Pub_Titulo
              xl.Cells(f1, 2) = "Linea: "
           Else
             xl.Cells(f1, 2) = "Linea: " & Trim(tab_llave!tab_NOMLARGO)
           End If
           wwfami = llave_rep02!art_numero
        End If
        'If llave_rep02!ART_KEY = 77854 Then Stop
        PUB_CODART = llave_rep02!ART_KEY
        SQ_OPER = 2
        pu_codcia = LK_CODCIA
        LEER_PRE_LLAVE
        Do Until pre_mayor.EOF
          If pre_mayor!PRE_FLAG_UNIDAD = "A" Then
            Exit Do
          End If
          pre_mayor.MoveNext
        Loop
        walterno = llave_rep02!ART_ALTERNO
        wdnombre = llave_rep02!art_nombre
        acu_cant_dia = 0
        acu_val_ingresos = 0
        acu_val_salidas = 0
        acu_cant_dia = 0
        WD_COSPRO = 0
        wSTOCK = 0
        WSCODART = llave_rep02!ART_KEY
        WD_COSPRO = 0 ' Val(llave_rep01!FAR_COSPRO)
        'CLD_TIPODSCTO = ? AND CLD_CODART = ? AND CLD_LISTADSCTO = ?
        f1 = f1 + 1
        xl.Cells(f1, 1) = "'" & Trim(walterno)
        xl.Cells(f1, 2) = Trim(wdnombre)
        xl.Cells(f1, 3) = Trim(pre_mayor!pre_unidad)
        If llave_rep02!ART_MONEDA = "S" Then
              xl.Cells(4, 3) = "Moneda : S/."
              wSTOCK = pre_mayor!PRE_PRE1
              xl.Cells(f1, 4) = Format(wSTOCK, "0.000")
        Else
              xl.Cells(4, 3) = "Moneda : US$."
              wSTOCK = pre_mayor!pre_pre11
              xl.Cells(f1, 4) = Format(wSTOCK, "0.000")
        End If
        c1 = 5
        WD_COSPRO = 0
        For nroitem = 0 To listat.ListCount - 1
            If Not listat.Selected(nroitem) Then GoTo saltanro
            PS_REP01(0) = LK_CODCIA
            PS_REP01(1) = Right(listat.List(nroitem), 6) ' Right(listat.Text, 6)
            PS_REP01(2) = llave_rep02!ART_KEY
            PS_REP01(3) = Right(listac.Text, 6)
            llave_rep01.Requery
            If Not llave_rep01.EOF Then
                WD_COSPRO = val(llave_rep01!cld_desto1)
            End If
            
            xl.Cells(f1, c1) = val(WD_COSPRO)
            xl.Cells(5, c1) = "DESC.(%)"
            xl.Cells(6, c1) = Left(listat.List(nroitem), 7)
            c1 = c1 + 1
            xl.Cells(f1, c1) = Format(val(wSTOCK) - (val(wSTOCK) * (WD_COSPRO / 100)), "0.000")
            xl.Cells(5, c1) = "PRECIO"
            xl.Cells(6, c1) = Left(listat.List(nroitem), 7)
            c1 = c1 + 1
saltanro:
         Next nroitem
         
        acu_saldo = acu_saldo + val(xl.Cells(f1, 4))
SIGUER:
llave_rep02.MoveNext
Loop
f1 = f1 + 1
    'xl.Cells(F1, 2) = "TOTAL GENERAL = "
    'xl.Cells(F1, 6) = Format(acu_saldo, "0.000")
  '  If F1 <> 6 Then
  '      wran1 = "D" & 6
  '      wran2 = "D" & F1 - 1
  '      wranF = "D" & F1
  '      xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  '   End If
  FrmImp2.lblproceso.Caption = "Procesando . . .  un Momento ."
  'xl.Application.Worksheets("HOJA DE COMISIONES x VENDEDOR").Range(wran1).Sort Key1:=xl.Application.Worksheets("HOJA DE COMISIONES x VENDEDOR").Range("O7")
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  GoSub LETRAS
'  xl.Application.Visible = True
  wranF = "A" & 5 & ":" & LETRAS(c1 - 1) & 5
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  wranF = "A" & 7 & ":" & LETRAS(c1 - 1) & 7
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  
 ' xl.Cells(2, 2) = "INFORME DE COMISIONES x VENDEDOR"
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
Exit Sub



LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"
Return

WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblproceso.Caption = "Abriendo , Archivo Comisiones.xls . . . "
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\LISTAPRE.xls", 0, True, 4, WPAS, WPAS

Return

Exit Sub
CANCELA:
  FrmImp2.Pantalla.Enabled = True
  FrmImp2.Pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblproceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
  
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.APPLICATION.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
Exit Sub
End Sub
''===========================================
Private Sub STOCKxUNIDAD_EXCEL(ByVal sWhere As String)
Dim PSRS As rdoQuery
Dim rs As rdoResultset
Dim PSPRE As rdoQuery
Dim RSPRE As rdoResultset

Dim oExcel As Object
Dim iRow As Integer
Dim TotalCantidad As Double
Dim TotalLitros As Double
Dim TotalMonto As Double
Dim CodigoTMP As Double
Dim CADENITA2 As String
Dim CADENITA As String
Dim SQL As String
Dim iCount As Integer
Dim grupo As String
Dim Codigo1 As String
Dim Cantidad As Double
Dim CantidadEquivalencia As Double
Dim CantidadTMP As Double
On Error GoTo ErrorGrave

Dim wsFECHA1, wsFECHA2
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo ErrorGrave
End If

CADENITA2 = ""
For fila = 0 To listacal.ListCount - 1
 listacal.ListIndex = fila
 If listacal.Selected(fila) Then
   CADENITA2 = CADENITA2 + "ART_CALIDAD = " & Trim(Right(listacal.Text, 8)) & " OR "
 End If
Next fila
If CADENITA2 <> "" Then
  CADENITA2 = "(" & Mid(CADENITA2, 1, Len(CADENITA2) - 4) & ")"
End If

CADENITA = ""
For fila = 0 To vendmulti.ListCount - 1
 vendmulti.ListIndex = fila
 If vendmulti.Selected(fila) Then
   CADENITA = CADENITA + "ART_FAMILIA = " & Trim(Right(vendmulti.Text, 8)) & " OR "
 End If
Next fila
If CADENITA <> "" Then
  CADENITA = " AND (" & Mid(CADENITA, 1, Len(CADENITA) - 4) & ")"
End If


SQL = "SELECT ARTI.ART_NOMBRE, ARTI.ART_CALIDAD, ARTI.ART_KEY, ARTI.ART_FAMILIA, ARTI.ART_ALTERNO, ARTI.ART_FLAG_STOCK, ARTICULO.ARM_STOCK, TABLAS.TAB_TIPREG, TABLAS.TAB_NOMLARGO "
SQL = SQL & "FROM ARTI ARTI INNER JOIN ARTICULO ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN TABLAS TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB "
SQL = SQL & "WHERE ARTI.ART_CALIDAD = 1 AND TABLAS.TAB_TIPREG = 122 AND ARTI.ART_FLAG_STOCK <> '' " & CADENITA
SQL = SQL & " ORDER BY ARTI.ART_FAMILIA ASC, ARTI.ART_ALTERNO Asc"

    If oExcel Is Nothing Then
       Set oExcel = CreateObject("Excel.Application")
    End If
    DoEvents
    oExcel.Workbooks.Open PUB_RUTA_OTRO & "STOCKxUNIDADxls.xls", 0, True, 4 ', PUB_CLAVE, PUB_CLAVE

    oExcel.Cells(1, 1) = Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia")))
    oExcel.Cells(2, 1) = "'STOCK AL " & Format(wsFECHA1, "dd/mm/yyyy")

    
    Set PSRS = CN.CreateQuery("", SQL)
    Set rs = PSRS.OpenResultset(rdOpenKeyset, rdConcurValues)
    
    SQL = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ?  ORDER BY PRE_SECUENCIA DESC" 'AND PRE_SECUENCIA = ?
    Set PSPRE = CN.CreateQuery("", SQL)
    PSPRE(0) = ""
    PSPRE(1) = 0
    Set RSPRE = PSPRE.OpenResultset(rdOpenKeyset, rdConcurValues)
    iRow = 7
    
    ProgBar.max = IIf(rs.RowCount > 0, rs.RowCount, 1)
    
    Do While Not rs.EOF
        With oExcel
        iCount = iCount + 1
        ProgBar.Value = iCount
        iRow = iRow + 1
        If grupo <> Trim(rs("TAB_NOMLARGO")) Then
            .Cells(iRow + 1, 1) = Trim(rs("TAB_TIPREG")) & " - " & Trim(rs("TAB_NOMLARGO"))
            iRow = iRow + 2
        End If
        
        If Trim(Codigo1) <> Trim(rs("ART_ALTERNO")) Then
            .Cells(iRow, 1) = Trim(rs("ART_ALTERNO"))
            .Cells(iRow, 2) = Trim(rs("ART_NOMBRE"))
        End If
        Codigo1 = rs("ART_ALTERNO")
        grupo = Trim(rs("TAB_NOMLARGO"))
        Cantidad = rs("ARM_STOCK")
        .Cells(iRow, 4) = Cantidad
        PSPRE(0) = LK_CODCIA
        PSPRE(1) = rs("ART_KEY")
        RSPRE.Requery
        
        Do While Not RSPRE.EOF
            If RSPRE.RowCount = 1 Then
                .Cells(iRow, 3) = RSPRE("PRE_EQUIV")
                .Cells(iRow, 6) = CantidadEquivalencia
                .Cells(iRow, 7) = Trim(RSPRE("PRE_UNIDAD"))
            End If
            CantidadEquivalencia = Int((Cantidad - CantidadTMP) / RSPRE("PRE_EQUIV"))
            If RSPRE("PRE_SECUENCIA") = 0 Then
                .Cells(iRow, 5) = Trim(RSPRE("PRE_UNIDAD"))
                .Cells(iRow, 8) = CantidadEquivalencia
                .Cells(iRow, 9) = Trim(RSPRE("PRE_UNIDAD"))
            Else
                .Cells(iRow, 3) = RSPRE("PRE_EQUIV")
                .Cells(iRow, 6) = CantidadEquivalencia
                .Cells(iRow, 7) = Trim(RSPRE("PRE_UNIDAD"))
            End If
            CantidadTMP = CantidadEquivalencia * RSPRE("PRE_EQUIV")
            RSPRE.MoveNext
        Loop
        End With
        CantidadEquivalencia = 0
        CantidadTMP = 0
        rs.MoveNext
    Loop
    
        
    'cmdmostrar.Enabled = True
    oExcel.DisplayAlerts = False
    oExcel.Visible = True
    
ErrorGrave:
    ProgBar.Value = 0
    ProgBar.Visible = False
    Set oExcel = Nothing
    
End Sub

