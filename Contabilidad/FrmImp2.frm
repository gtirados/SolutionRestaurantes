VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmImp2 
   BackColor       =   &H00FAEFDA&
   Caption         =   "Reportes"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8445
   Icon            =   "FrmImp2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox cheasiento 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Pasar a Contabilidad"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4200
      TabIndex        =   55
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FAEFDA&
      Height          =   615
      Left            =   3735
      TabIndex        =   20
      Top             =   15
      Width           =   4455
      Begin VB.Label lblreporte 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   21
         Top             =   225
         Width           =   4140
      End
   End
   Begin VB.Frame fradeudas 
      BackColor       =   &H00FAEFDA&
      Height          =   1215
      Left            =   540
      TabIndex        =   31
      Top             =   3495
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Fechas de Nuevo Vcto."
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
         Height          =   255
         Left            =   105
         TabIndex        =   54
         Top             =   120
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox Cheop2 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Deudas por Cobrar Vencidas"
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
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   600
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Cheop1 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Deudas por Cobrar del Dia"
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
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   360
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Cheop3 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Deudas por Cobrar por Vencer"
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
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   840
         Value           =   1  'Checked
         Width           =   2535
      End
   End
   Begin MSMask.MaskEdBox txtfecha 
      Height          =   285
      Left            =   2280
      TabIndex        =   53
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.ComboBox cmdMoneda 
      Height          =   315
      ItemData        =   "FrmImp2.frx":0442
      Left            =   120
      List            =   "FrmImp2.frx":044C
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame fcontab 
      BackColor       =   &H00FAEFDA&
      Height          =   2535
      Left            =   495
      TabIndex        =   42
      Top             =   3975
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CheckBox chenivel 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Nivel 6"
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
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   49
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox chenivel 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Nivel 5"
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
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   48
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chenivel 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Nivel 4"
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
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   47
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chenivel 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Nivel 3"
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
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox chenivel 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Nivel 2"
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
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chenivel 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Nivel 1"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblcontab 
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione los Niveles para impresión"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame fracli 
      BackColor       =   &H00FAEFDA&
      Height          =   2535
      Left            =   3855
      TabIndex        =   34
      Top             =   3795
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtDias2 
         Height          =   285
         Index           =   1
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   38
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtDias2 
         Height          =   285
         Index           =   0
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   37
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtDias1 
         Height          =   285
         Index           =   1
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   36
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtDias1 
         Height          =   285
         Index           =   0
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   35
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblcli 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   41
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblcli 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   40
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblcli 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese 2 Rangos :"
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
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   39
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CheckBox Chesup 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Suprimir las Columnas con 0"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   840
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame frazonas 
      BackColor       =   &H00FAEFDA&
      Height          =   2535
      Left            =   45
      TabIndex        =   32
      Top             =   3630
      Visible         =   0   'False
      Width           =   8415
      Begin VB.OptionButton opzonas 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Zonas"
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
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton opzonas 
         BackColor       =   &H00FAEFDA&
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
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   2175
      End
      Begin VB.OptionButton opzonas 
         BackColor       =   &H00FAEFDA&
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
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
      Begin VB.ListBox zonas 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   3720
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblzonas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zonas :"
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
         Left            =   2040
         TabIndex        =   33
         Top             =   600
         Width           =   540
      End
   End
   Begin VB.ListBox vendmulti 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   -30
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   1575
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CheckBox chestock 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Incluir Valorado"
      Height          =   255
      Left            =   6240
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame FRASTOCK 
      BackColor       =   &H00FAEFDA&
      Height          =   2415
      Left            =   3120
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   5175
      Begin VB.ComboBox CmbCalidad 
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
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1080
         Width           =   2055
      End
      Begin VB.ListBox subfami 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   1560
         Width           =   2775
      End
      Begin VB.ListBox fami 
         Appearance      =   0  'Flat
         Height          =   705
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblcalidad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calidad"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3000
         TabIndex        =   28
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Familia :"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Familia :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   900
      End
   End
   Begin VB.ListBox listVen 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   210
      TabIndex        =   6
      Top             =   1830
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox PROV 
      Height          =   1860
      Left            =   90
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   1215
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton pantalla 
      Caption         =   "Por &Pantalla .."
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cerrar 
      Caption         =   "Ce&rrar"
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
      Left            =   4680
      TabIndex        =   17
      Top             =   3840
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar ProgBar 
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      Top             =   3360
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   10
      Top             =   3600
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
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Left            =   3360
      TabIndex        =   30
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox txtCampo1 
      Height          =   285
      Left            =   3240
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label lblfecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha :"
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
      Left            =   1680
      TabIndex        =   52
      Top             =   600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblmoneda 
      BackStyle       =   0  'Transparent
      Caption         =   "Moneda :"
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
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblcampo2 
      AutoSize        =   -1  'True
      Caption         =   "Campo1"
      Height          =   195
      Left            =   4800
      TabIndex        =   24
      Top             =   1800
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblcampo1 
      AutoSize        =   -1  'True
      Caption         =   "Campo1"
      Height          =   195
      Left            =   3240
      TabIndex        =   23
      Top             =   1800
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblstock 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedores"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   180
      TabIndex        =   22
      Top             =   870
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblProceso 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Procesando ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Top             =   3600
      Visible         =   0   'False
      Width           =   3375
   End
End
Attribute VB_Name = "FrmImp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xl As Object
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset
Dim PS_REP03 As rdoQuery
Dim llave_rep03 As rdoResultset
Dim PS_REP04 As rdoQuery
Dim llave_rep04 As rdoResultset
Dim wranF, wran1, wran2, WPAS
Dim C1 As Integer
Dim F1 As Integer
Dim xcuenta As Integer
Dim i As Integer
Dim Mensaje, titulo, valorpred As String
Dim Wfile  As String
Dim WFORM  As String


Private Sub cerrar_Click()
Unload FrmImp2
End Sub

Private Sub cmdMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If txtfecha.Visible = True Then txtfecha.SetFocus
End If
End Sub

Private Sub cmdultima_Click()
Dim var
Dim nom
If Trim((listVen.Text)) = "" Then Exit Sub
lblProceso.Visible = True
DoEvents
nom = Trim(Left(listVen.Text, 3))
var = CONS_ADMIN & "OFFICE\PLANVEN" & nom & ".XLS"

If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
End If
On Error GoTo no_existe
xl.Workbooks.Open var, 0, False, 4
xl.Application.Visible = True
On Error GoTo 0
Set xl = Nothing
lblProceso.Visible = False
Exit Sub
no_existe:
If Err.Number = 1004 Then MsgBox "Planilla no emitida ó ya Procesada.", 48, Pub_Titulo
Set xl = Nothing
lblProceso.Visible = False
Exit Sub

End Sub

Private Sub fami_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If pantalla.Enabled Then pantalla.SetFocus
End If
End Sub

Private Sub Form_Activate()
If CmbCalidad.ListCount <> 0 Then CmbCalidad.ListIndex = 0
End Sub

Private Sub Form_Load()
CenterMe FrmImp2
Screen.MousePointer = 11
If tra_llave.EOF Then
   Screen.MousePointer = 0
   Exit Sub
End If
Screen.MousePointer = 0
Wfile = Trim(tra_llave(3))
WFORM = Trim(tra_llave(7))
lblreporte.Caption = Trim(tra_llave(1))
If Wfile = "IMP_STOCK" Then
 chestock.Visible = True
 lblstock.Visible = True
 pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CP = 'P'  AND CLI_CODCIA = ? ORDER BY CLI_NOMBRE"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
 PS_REP01(0) = LK_CODCIA
 llave_rep01.Requery
 Do Until llave_rep01.EOF
     PROV.AddItem llave_rep01!cli_nombre & String(25, " ") & llave_rep01!CLI_CODCLIE
     llave_rep01.MoveNext
 Loop
 PROV.Visible = True
 PUB_CODCIA = LK_CODCIA
End If
If Wfile = "HSTOCK" Or Wfile = "STOCKVA" Or Wfile = "STOCKSEM" Then    'CRYSTAL REPORT
 fami.Height = 1830
 PUB_CODCIA = LK_CODCIA
 LLENADOS fami, 122
 FRASTOCK.Left = 2280
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
If Wfile = "IMP_PLANILLA" Or Wfile = "IMP_CONSUPLAN" Or Wfile = "IMP_COMISION" Then
 Dim codi As String * 5
 lblstock.Caption = "Vendedor : "
 lblstock.Visible = True
 pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = ? ORDER BY VEM_CODVEN"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
 PS_REP01(0) = LK_CODCIA
 llave_rep01.Requery
 listVen.Clear
 Do Until llave_rep01.EOF
     codi = llave_rep01!vem_codven
     listVen.AddItem codi & Trim(llave_rep01!VEM_NOMBRE)
     llave_rep01.MoveNext
 Loop
 listVen.Visible = True
 lblcampo1.Caption = "Fecha Inicial : "
 lblcampo1.Visible = True
 'txtCampo1.MaxLength = 10
 txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
 If Wfile = "IMP_PLANILLA" Then
   cmdultima.Visible = True
 End If
 If Wfile <> "IMP_CONSUPLAN" Then
  lblcampo2.Caption = "Fecha Final: "
  lblcampo2.Visible = True
  'txtCampo2.MaxLength = 10
  txtCampo2.Mask = "##/##/####"
  txtCampo2.Visible = True
 End If
 listVen.TabIndex = 0
End If
If Wfile = "VENTA_X_VEND" Or Wfile = "VEN_NEGOCIOS" Or Wfile = "VEN_NEGOCIOS_CLI" Or Wfile = "HER_NEGOCIOS" Or Wfile = "HER_VENDEDOR" Or Wfile = "HER_REGISTRO" Or Wfile = "REGCOMP" Or Wfile = "RESUMEN_TRANSA" Then
 If Wfile = "HER_REGISTRO" Or Wfile = "REGCOMP" Then
   txtCampo1.Text = Format(LK_FECHA_COP1, "dd/mm/yyyy")
   txtCampo2.Text = Format(LK_FECHA_COP2, "dd/mm/yyyy")
   cheasiento.Visible = True
 Else
   txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 End If
 lblcampo1.Caption = "Fecha Inicial : "
 lblcampo1.Visible = True
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
 lblcampo2.Caption = "Fecha Final: "
 lblcampo2.Visible = True
 txtCampo2.Mask = "##/##/####"
 txtCampo2.Visible = True
 If Wfile = "VENTA_X_VEND" Or Wfile = "HER_VENDEDOR" Or Wfile = "VEN_NEGOCIOS" Or Wfile = "VEN_NEGOCIOS_CLI" Then
    lblstock.Caption = "Vendedor : "
    lblstock.Visible = True
    pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = ? ORDER BY VEM_CODVEN"
    Set PS_REP01 = CN.CreateQuery("", pub_cadena)
    Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    PS_REP01(0) = LK_CODCIA
    llave_rep01.Requery
    vendmulti.Clear
    Do Until llave_rep01.EOF
        codi = llave_rep01!vem_codven
        vendmulti.AddItem codi & Trim(llave_rep01!VEM_NOMBRE)
        llave_rep01.MoveNext
    Loop
    vendmulti.Visible = True
    If Wfile = "VENTA_X_VEND" Or Wfile = "VEN_NEGOCIOS" Or Wfile = "VEN_NEGOCIOS_CLI" Then Chesup.Visible = True
 End If
 
End If
If Wfile = "VENTA_X_VEND" Then 'Or Wfile = "HER_VENDEDOR" Then
    fami.Height = 1830
    PUB_CODCIA = LK_CODCIA
    LLENADOS fami, 122
    FRASTOCK.Left = 5980
    FRASTOCK.Visible = True
    subfami.Visible = False
    lblcalidad.Visible = False
    CmbCalidad.Visible = False
End If
If Wfile = "DEDVENC" Then
    pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = ? ORDER BY VEM_CODVEN"
    Set PS_REP01 = CN.CreateQuery("", pub_cadena)
    Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    PS_REP01(0) = LK_CODCIA
    llave_rep01.Requery
    vendmulti.Clear
    Do Until llave_rep01.EOF
        codi = llave_rep01!vem_codven
        vendmulti.AddItem codi & Trim(llave_rep01!VEM_NOMBRE)
        llave_rep01.MoveNext
    Loop
    vendmulti.Visible = True
    fradeudas.Visible = True
    
    lblcampo1.Caption = "Fecha Inicial : "
    lblcampo1.Visible = True
    txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    txtCampo1.Mask = "##/##/####"
    txtCampo1.Visible = True
    lblcampo2.Caption = "Fecha Final: "
    lblcampo2.Visible = True
    txtCampo2.Mask = "##/##/####"
    txtCampo2.Visible = True
ElseIf Wfile = "COMISIONES" Then
    pantalla.TabIndex = 0
    lblcampo1.Caption = "Fecha Inicial : "
    lblcampo1.Visible = True
    txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    txtCampo1.Mask = "##/##/####"
    txtCampo1.Visible = True
    lblcampo2.Caption = "Fecha Final: "
    lblcampo2.Visible = True
    txtCampo2.Mask = "##/##/####"
    txtCampo2.Visible = True
ElseIf Wfile = "ZONA_X_NEGO" Then
 PUB_CODCIA = "00"
 LLENADOS zonas, 35
 frazonas.Visible = True
 opzonas(0).Caption = BUSCA_ETIQUETA(10)
 opzonas(1).Caption = BUSCA_ETIQUETA(11)
 opzonas(2).Caption = BUSCA_ETIQUETA(12)
 Chesup.Visible = True
ElseIf Wfile = "MOROSIDAD" Then
 fracli.Visible = True
 txtDias1(0).Text = "1"
 txtDias1(1).Text = "7"
 txtDias2(0).Text = "8"
 txtDias2(1).Text = "+"
 lblstock.Caption = "Vendedor : "
 lblstock.Visible = True
 pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = ? ORDER BY VEM_CODVEN"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
 PS_REP01(0) = LK_CODCIA
 llave_rep01.Requery
 vendmulti.Clear
 Do Until llave_rep01.EOF
     codi = llave_rep01!vem_codven
     vendmulti.AddItem codi & Trim(llave_rep01!VEM_NOMBRE)
     llave_rep01.MoveNext
 Loop
 vendmulti.Visible = True
ElseIf Wfile = "BALANCE" Then
 fcontab.Visible = True
 pub_cadena = "SELECT * FROM COPARAM WHERE COP_CODCIA = ? "
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
 PS_REP01(0) = LK_CODCIA
 llave_rep01.Requery
 If llave_rep01.EOF Then
   chenivel(0).Enabled = False
   chenivel(1).Enabled = False
   chenivel(2).Enabled = False
   chenivel(3).Enabled = False
   chenivel(4).Enabled = False
   chenivel(5).Enabled = False
 Else
  If Nulo_Valor0(llave_rep01!COP_DIG_NIV1) <> 0 Then chenivel(0).Enabled = True Else chenivel(0).Enabled = False
  If Nulo_Valor0(llave_rep01!COP_DIG_NIV2) <> 0 Then chenivel(1).Enabled = True Else chenivel(1).Enabled = False
  If Nulo_Valor0(llave_rep01!COP_DIG_NIV3) <> 0 Then chenivel(2).Enabled = True Else chenivel(2).Enabled = False
  If Nulo_Valor0(llave_rep01!COP_DIG_NIV4) <> 0 Then chenivel(3).Enabled = True Else chenivel(3).Enabled = False
  If Nulo_Valor0(llave_rep01!COP_DIG_NIV5) <> 0 Then chenivel(4).Enabled = True Else chenivel(4).Enabled = False
  If Nulo_Valor0(llave_rep01!COP_DIG_NIV6) <> 0 Then chenivel(5).Enabled = True Else chenivel(5).Enabled = False
End If
End If
If Wfile = "RESUMEN" Then
    lblcampo1.Caption = "Fecha Inicial : "
    lblcampo1.Visible = True
    txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    txtCampo1.Mask = "##/##/####"
    txtCampo1.Visible = True
    lblcampo2.Caption = "Fecha Final: "
    lblcampo2.Visible = True
    txtCampo2.Mask = "##/##/####"
    txtCampo2.Visible = True
End If
If Wfile = "REPO_CAJA_GEN" Then
 lblmoneda.Visible = True
 cmdMoneda.Visible = True
 cmdMoneda.ListIndex = 0
 cmdMoneda.TabIndex = 0
 lblfecha.Visible = True
 txtfecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtfecha.Mask = "##/##/####"
 txtfecha.Visible = True
End If
If Wfile = "REPO_CAJA" Then
 lblmoneda.Visible = True
 cmdMoneda.Visible = True
 cmdMoneda.ListIndex = 0
 cmdMoneda.TabIndex = 0
 lblfecha.Visible = True
 txtfecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtfecha.Mask = "##/##/####"
 txtfecha.Visible = True
End If

End Sub
Public Sub IMP_STOCK()
'On Error GoTo FINTODO
Dim ws_clave As String
Dim LETRAS(24) As String * 1
Dim WSFECHA As Date
Dim WCODCLIE
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
 If LK_CODUSU = "ADMIN" And Trim(usu!usu_key) = "ADMIN" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  If Trim(usu!usu_key) = "SUPERVISOR" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  usu.MoveNext
Loop

FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Max = llave_rep01.RowCount
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Visible = True
DoEvents
FrmImp2.lblProceso.Visible = True
FrmImp2.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
xl.Worksheets(1).Activate
GoSub LETRAS
xcuenta = 0
xl.Cells(1, 1) = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
xl.Cells(2, 2) = "'" & LK_FECHA_DIA
F1 = 5  'Fila Inicial
WCODCLIE = llave_rep01!art_codclie
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
 WNOMCLIE = Trim(cli_llave!cli_nombre)
End If
FrmImp2.lblProceso.Caption = "Procesando . . .  un Momento ."
DoEvents
F1 = F1 + 1
xl.Cells(F1, 1) = WNOMCLIE
xl.Cells(F1, 1).Font.Bold = True
F1 = F1 + 1
fila = 0
acu_val_ingresos = 0
acu_val_salidas = 0
acu_val_saldos = 0
Do Until llave_rep01.EOF
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   If WCODCLIE <> llave_rep01!art_codclie Then
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
          WNOMCLIE = Trim(cli_llave!cli_nombre)
        End If
        'F1 = F1 + 1
        xl.Cells(F1, 1) = WNOMCLIE
        xl.Cells(F1, 1).Font.Bold = True
        WCODCLIE = llave_rep01!art_codclie
        F1 = F1 + 1
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
      MsgBox "Error en la unidad. " & llave_rep01!ART_NOMBRE
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
   If llave_rep02!far_SIGNO_aRM = 1 Then
     ws_ingresos = ws_ingresos + (llave_rep02!FAR_CANTIDAD / llave_rep03!PRE_EQUIV)
     If llave_rep02!FAR_TIPMOV = 20 Then
        val_ingresos = val_ingresos + llave_rep02!far_precio_neto
     Else
       val_ingresos = val_ingresos + ((llave_rep02!FAR_CANTIDAD / llave_rep03!PRE_EQUIV) * llave_rep02!far_precio_neto)
     End If
   ElseIf llave_rep02!far_SIGNO_aRM = -1 Then
     ws_salidas = ws_salidas + (llave_rep02!FAR_CANTIDAD / llave_rep03!PRE_EQUIV)
     val_salidas = val_salidas + ((llave_rep02!FAR_CANTIDAD / llave_rep02!FAR_equiv) * llave_rep02!far_PRECIO)
   End If
   llave_rep02.MoveNext
   Loop
PASADIA:
   xl.Cells(F1, 1) = Trim(llave_rep01!ART_NOMBRE)
   xl.Cells(F1, 2) = Left(llave_rep03!pre_unidad, 10)
   If chestock.Value = 1 Then
     If ws_ingresos <> 0 Then xl.Cells(F1, 3) = ws_ingresos
     If ws_salidas <> 0 Then xl.Cells(F1, 4) = ws_salidas
     'xl.Cells(f1, 6) = Format(llave_rep01!ART_COSPRO * llave_rep03!PRE_EQUIV, "0.00")
     If val_ingresos <> 0 Then xl.Cells(F1, 7) = val_ingresos
     If val_salidas <> 0 Then xl.Cells(F1, 8) = val_salidas
     'xl.Cells(f1, 9) = (arm_llave!arm_stock / llave_rep03!PRE_EQUIV) * (llave_rep01!ART_COSPRO * llave_rep03!PRE_EQUIV)
     acu_val_saldos = acu_val_saldos + Val(xl.Cells(F1, 9))
     acu_val_ingresos = acu_val_ingresos + val_ingresos
     acu_val_salidas = acu_val_salidas + val_salidas
   End If
   xl.Cells(F1, 5) = Format(arm_llave!arm_stock / llave_rep03!PRE_EQUIV, "0.00")
   F1 = F1 + 1
   llave_rep01.MoveNext
Loop
  If chestock.Value = 1 Then
    xl.Cells(F1, 1) = "Totales "
    xl.Cells(F1, 7) = acu_val_ingresos
    xl.Cells(F1, 7).Font.Bold = True
    xl.Cells(F1, 8) = acu_val_salidas
    xl.Cells(F1, 8).Font.Bold = True
    xl.Cells(F1, 9) = acu_val_saldos
    xl.Cells(F1, 9).Font.Bold = True
  End If
  If chestock.Value = 0 Then
   xl.Range("C4:D5").Delete 4
   xl.Range("D3:G5").Delete 4
  End If
  FrmImp2.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False

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
  FrmImp2.lblProceso.Caption = "Abriendo , Archivo Stock.xls . . . "
  DoEvents
  WPAS = PUB_CLAVE
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  xl.Workbooks.Open CONS_ADMIN & "OFFICE\stock.xls", 0, True, 4, WPAS, WPAS
Return

Exit Sub
CANCELA:
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.Application.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
 
End Sub

Private Sub listVen_Click()
 If txtCampo1.Visible Then
  txtCampo1.SetFocus
 ElseIf pantalla.Enabled Then
  pantalla.SetFocus
 End If

End Sub

Private Sub listven_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If txtCampo1.Visible Then
  txtCampo1.SetFocus
 ElseIf pantalla.Enabled Then
  pantalla.SetFocus
 End If
End If
End Sub

Private Sub opzonas_Click(Index As Integer)
Dim cod As Integer
lblzonas.Caption = Trim(opzonas(Index).Caption) & " :"
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
   If Not IsDate(wsFECHA1) Or Val(Mid(wsFECHA1, 4, 2)) > 12 Then
     MsgBox "Fecha invalida .... ", 48, Pub_Titulo
     txtCampo1.SetFocus
     Exit Sub
   End If
   If Right(txtCampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtCampo2.Text, 8)
   Else
     wsFECHA2 = Trim(txtCampo2.Text)
   End If
   If txtCampo2.Visible Then
    If Not IsDate(wsFECHA2) Or Val(Mid(wsFECHA2, 4, 2)) > 12 Then
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
   If Trim(listVen.Text) = "" Then
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
   
ElseIf Wfile = "VENTA_X_VEND" Or Wfile = "VEN_NEGOCIOS" Or Wfile = "VEN_NEGOCIOS_CLI" Or Wfile = "HER_NEGOCIOS" Or Wfile = "HER_VENDEDOR" Or Wfile = "HER_REGISTRO" Or Wfile = "REGCOMP" Then
   If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
   Else
     wsFECHA1 = Trim(txtCampo1.Text)
   End If
   If Not IsDate(wsFECHA1) Or Val(Mid(wsFECHA1, 4, 2)) > 12 Then
     MsgBox "Fecha invalida .... ", 48, Pub_Titulo
     Azul2 txtCampo1, txtCampo1
     Exit Sub
   End If
   If Right(txtCampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtCampo2.Text, 8)
   Else
     wsFECHA2 = Trim(txtCampo2.Text)
   End If
   If txtCampo2.Visible Then
    If Not IsDate(wsFECHA2) Or Val(Mid(wsFECHA2, 4, 2)) > 12 Then
      MsgBox "Fecha invalida .... ", 48, Pub_Titulo
      Azul2 txtCampo2, txtCampo2
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
    End If
End If
If Wfile = "DEDVENC" Then

ElseIf Wfile = "HSTOCK" Then  'CRYSTAL REPORT
 Call HSTOCK
ElseIf Wfile = "STOCKVA" Or Wfile = "STOCKSEM" Then
 Call STOCKVA
ElseIf Wfile = "ZONA_X_NEGO" Then
 Call ZONA_X_NEGO
ElseIf Wfile = "MOROSIDAD" Then
 Call MOROSIDAD
ElseIf Wfile = "BALANCE" Then
 
ElseIf Wfile = "REPO_CAJA_GEN" Then
 Call REPO_CAJA_GEN
ElseIf Wfile = "REPO_CAJA" Then
 'Call REPO_CAJA
  Call REPO_CAJA_GEN
ElseIf Wfile = "SALDO_CAR" Then
 Call Saldo_Car
ElseIf Wfile = "SALDO_CAR_TE" Then
 Call Saldo_Car_TE
End If

Exit Sub
SALE:
ProgBar.Visible = False
lblProceso.Visible = False
pantalla.Enabled = True
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
If Right(txtCampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtCampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtCampo2.Text)
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

If Trim(listVen.Text) = "" Then
  MsgBox "Seleccione Datos, para Procesar ", 48, Pub_Titulo
  GoTo CANCELA
End If
pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
wcodven = Val(Left(listVen.Text, 4))
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
 If LK_CODUSU = "ADMIN" And Trim(usu!usu_key) = "ADMIN" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  If Trim(usu!usu_key) = "SUPERVISOR" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  usu.MoveNext
Loop
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Max = llave_rep01.RowCount
FrmImp2.ProgBar.Value = 0
FrmImp2.lblProceso.Visible = True
FrmImp2.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
GoSub LETRAS
FrmImp2.ProgBar.Visible = True
DoEvents
xcuenta = 0
xl.Cells(1, 1) = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
xl.Cells(2, 2) = "'" & LK_FECHA_DIA
xl.Cells(3, 1) = "Vendedor"

xl.Cells(3, 2) = Trim(listVen.Text)
F1 = 5  'Fila Inicial

FrmImp2.lblProceso.Caption = "Procesando . . .  un Momento ."
DoEvents
fila = 0
acu_val_ingresos = 0
acu_val_salidas = 0
Do Until llave_rep01.EOF
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   F1 = F1 + 1
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   pu_codclie = llave_rep01!CAR_CODCLIE
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
      MsgBox "Error en Codigo de cliente, NO EXISTE ...", 48, Pub_Titulo
      GoTo CANCELA
   End If
   xl.Cells(F1, 1) = Trim(cli_llave!CLI_CODCLIE)
   xl.Cells(F1, 2) = Trim(cli_llave!cli_nombre)
'   xl.Cells(F1, 2).Font.Bold = True
  ' xl.Cells(F1, 3) = llave_rep01!CAR_CODVEN
   xl.Cells(F1, 3).HorizontalAlignment = xlCenter
   If llave_rep01!CAR_fbg = "F" Then
     xl.Cells(F1, 3) = "FAC."
   ElseIf llave_rep01!CAR_fbg = "B" Then
     xl.Cells(F1, 3) = "BOL."
   ElseIf llave_rep01!CAR_fbg = "G" Then
     xl.Cells(F1, 3) = "GUIA"
   End If
   xl.Cells(F1, 4).HorizontalAlignment = xlCenter
   xl.Cells(F1, 4) = llave_rep01!CAR_numSER
   xl.Cells(F1, 5).HorizontalAlignment = xlCenter
   xl.Cells(F1, 5) = llave_rep01!car_numFAC
   xl.Cells(F1, 5).HorizontalAlignment = xlCenter
   xl.Cells(F1, 6) = "'" & llave_rep01!CAR_FECHA_INGR
   xl.Cells(F1, 7) = "'" & llave_rep01!CAR_FECHA_VCTO
   xl.Cells(F1, 8) = llave_rep01!CAR_IMPORTE
   xl.Cells(F1, 8).HorizontalAlignment = xlRight
   xl.Cells(F1, 9).NumberFormat = "#######.00"
   xl.Cells(F1, 10).NumberFormat = "dd/mm/yyyy"
   xl.Cells(F1, 11) = ""
   xl.Cells(F1, 12) = llave_rep01!CAR_CODCLIE
   xl.Cells(F1, 13) = llave_rep01!CAR_CODCIA
   xl.Cells(F1, 14) = llave_rep01!CAR_SERDOC
   xl.Cells(F1, 15) = llave_rep01!CAR_NUMDOC
   xl.Cells(F1, 16) = llave_rep01!CAR_TIPDOC
   
   llave_rep01.MoveNext
Loop
  FrmImp2.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(2, 2) = "V I S I T A S   A   C L I E N T E S"
  xl.ActiveCell.Range("I6").Activate
  wranF = "A" & 1 & ":J" & F1
  xl.Worksheets(1).Range(wranF).Font.Name = "Draft 17cpi"
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  xl.Workbooks(1).Save
  xl.Application.Visible = True
  DoEvents
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
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
  Dim dd As Excel.Application
  Dim wsfile1
  
  wsfile1 = "PlanVen" & wcodven & ".XLS"
  wsfile = CONS_ADMIN & "OFFICE\" & wsfile1
  On Error GoTo OJO
  Kill wsfile
  On Error GoTo 0
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  FrmImp2.lblProceso.Caption = "Configurando Hoja de Calculo... un Momento ."
  DoEvents
  xl.SheetsInNewWorkbook = 1
  xl.Workbooks.Add
  xl.Worksheets(1).Name = "COBRANZAS"
  xl.Windows(1).Caption = " Vendedor :  " & Trim(listVen.Text)
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
  xl.Application.Visible = False
  With xl.Worksheets(1).PageSetup
    .TopMargin = 28.6
    .HeaderMargin = 28.6
    .PrintTitleRows = "$1:$5"
  End With
  FrmImp2.lblProceso.Caption = "Abriendo , Archivo PLANILLAS.xls . . . "
  DoEvents
  WPAS = PUB_CLAVE
  'Set xl = Nothing
Return

Exit Sub
CANCELA:
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  pantalla.Enabled = True
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
 xl.Application.Visible = True
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
pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
wcodven = Val(Left(listVen.Text, 4))
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
 If LK_CODUSU = "ADMIN" And Trim(usu!usu_key) = "ADMIN" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  If Trim(usu!usu_key) = "SUPERVISOR" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  usu.MoveNext
Loop
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Max = llave_rep01.RowCount
FrmImp2.ProgBar.Value = 0
FrmImp2.lblProceso.Visible = True
FrmImp2.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
FrmImp2.ProgBar.Visible = True
DoEvents
xl.Worksheets(1).Activate
GoSub LETRAS
xcuenta = 0
xl.Cells(1, 1) = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
xl.Cells(2, 2) = "'" & LK_FECHA_DIA
xl.Cells(3, 2) = "Vendedor : " & wcodven
F1 = 5  'Fila Inicial

FrmImp2.lblProceso.Caption = "Procesando . . .  un Momento ."
DoEvents
fila = 0
acu_val_ingresos = 0
acu_val_salidas = 0
Do Until llave_rep01.EOF
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   F1 = F1 + 1
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   pu_codclie = llave_rep01!CAA_CODCLIE
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
      MsgBox "Error en Codigo de cliente, NO EXISTE ...", 48, Pub_Titulo
      GoTo CANCELA
   End If
   xl.Cells(F1, 1) = Trim(cli_llave!CLI_CODCLIE)
   xl.Cells(F1, 2) = Trim(cli_llave!cli_nombre)
   xl.Cells(F1, 2).Font.Bold = True
   xl.Cells(F1, 3).HorizontalAlignment = xlCenter
   If llave_rep01!CAA_FBG = "F" Then
     xl.Cells(F1, 4) = "FAC."
   ElseIf llave_rep01!CAA_FBG = "B" Then
     xl.Cells(F1, 4) = "BOL."
   ElseIf llave_rep01!CAA_FBG = "G" Then
     xl.Cells(F1, 4) = "GUIA"
   End If
   xl.Cells(F1, 4).HorizontalAlignment = xlCenter
   xl.Cells(F1, 5) = "'" & llave_rep01!CAa_numser & " - " & llave_rep01!CAa_numfac
   xl.Cells(F1, 5).HorizontalAlignment = xlCenter
   xl.Cells(F1, 6) = "'" & llave_rep01!CAA_FECHA
   xl.Cells(F1, 7) = "'" & llave_rep01!CAA_FECHA_VCTO
   xl.Cells(F1, 8) = llave_rep01!CAA_SALDO - llave_rep01!CAA_IMPORTE
   xl.Cells(F1, 9).HorizontalAlignment = xlRight
   xl.Cells(F1, 9) = Val(llave_rep01!CAA_IMPORTE)
   xl.Cells(F1, 10) = "'" & llave_rep01!CAA_FECHA_VCTO
   llave_rep01.MoveNext
Loop
  wran1 = "I" & 6
  wran2 = "I" & F1
  wranF = "I" & F1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  FrmImp2.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(2, 2) = "I N F O R M E  D E  C O B R A N Z A"
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
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
  Dim dd As Excel.Application
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblProceso.Caption = "Abriendo , Archivo CONSUPLAN.xls . . . "
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open CONS_ADMIN & "OFFICE\CONSUPLAN.xls", 0, True, 4, WPAS, WPAS
 
Return

Exit Sub
CANCELA:
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  pantalla.Enabled = True
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
 xl.Application.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
 
End Sub

Private Sub txtcampo1_GotFocus()
'Azul txtCampo1, txtCampo1
End Sub

Private Sub txtCampo1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
 Exit Sub
End If

If txtCampo2.Visible Then
 If Not IsDate(txtCampo2) Then
   txtCampo2.Text = Format(txtCampo1.Text, "dd/mm/yyyy")
 End If
 Azul2 txtCampo2, txtCampo2
Else
pantalla.SetFocus
End If
 

End Sub

Private Sub txtcampo2_GotFocus()
'Azul txtCampo2, txtCampo2
End Sub

Private Sub txtcampo2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
 Exit Sub
End If
If pantalla.Enabled Then
   pantalla.SetFocus
End If

End Sub
Public Sub IMP_COMISION()
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
pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
wcodven = Val(Left(listVen.Text, 4))
pub_cadena = "SELECT * FROM CARACU WHERE CAA_CODCIA = ? AND CAA_CODVEN = ? AND CAA_FECHA >= ? AND CAA_FECHA <= ? AND CAA_ESTADO <> 'E'  AND CAA_CONCEPTO <> 'Extorno -' AND CAA_SIGNO_CAR = -1 AND CAA_TIPDOC ='FA' AND (CAA_FBG ='F' OR CAA_FBG ='B' OR CAA_FBG ='G')  ORDER BY CAA_CODCLIE, CAA_FECHA,CAA_NUM_OPER, CAA_SALDO_CAR"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
DoEvents
Dim wsFECHA1, wsFECHA2
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtCampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtCampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtCampo2.Text)
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
FrmImp2.ProgBar.Max = llave_rep01.RowCount
FrmImp2.ProgBar.Value = 0
FrmImp2.lblProceso.Visible = True
FrmImp2.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
FrmImp2.ProgBar.Visible = True
DoEvents
xl.Worksheets(1).Activate
GoSub LETRAS
xcuenta = 0
xl.Cells(1, 1) = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
xl.Cells(3, 1) = "'Comisiones del " & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
xl.Cells(4, 1) = "Vendedor : " & Trim(listVen.Text)

F1 = 6  'Fila Inicial

FrmImp2.lblProceso.Caption = "Procesando . . .  un Momento ."
DoEvents
fila = 0
acu_val_ingresos = 0
acu_val_salidas = 0
Do Until llave_rep01.EOF
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   F1 = F1 + 1
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   pu_codclie = llave_rep01!CAA_CODCLIE
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
      MsgBox "Error en Codigo de cliente, NO EXISTE ...", 48, Pub_Titulo
      GoTo CANCELA
   End If
   xl.Cells(F1, 1) = wcodven
   If llave_rep01!CAA_FBG = "F" Then
     xl.Cells(F1, 2) = "FAC."
   ElseIf llave_rep01!CAA_FBG = "B" Then
     xl.Cells(F1, 2) = "BOL."
   ElseIf llave_rep01!CAA_FBG = "G" Then
     xl.Cells(F1, 2) = "GUIA"
   End If
   xl.Cells(F1, 3) = "'" & llave_rep01!CAa_numser & " - " & llave_rep01!CAa_numfac
   xl.Cells(F1, 4) = Trim(cli_llave!cli_nombre)
   SQ_OPER = 1
   pu_cp = "C"
   pu_codclie = cli_llave!CLI_CODCLIE
   pu_codcia = LK_CODCIA
   PUB_SERDOC = llave_rep01!caa_serdoc
   PUB_NUMDOC = llave_rep01!CAA_NUMDOC
   PUB_TIPDOC = llave_rep01!CAA_TIPDOC
   LEER_CAR_LLAVE
   If car_llave.EOF Then
    MsgBox "Documento Extornado.... ", 48, Pub_Titulo
    GoTo otrito
   End If
   xl.Cells(F1, 5) = "'" & car_llave!CAR_FECHA_INGR
   xl.Cells(F1, 6) = "'" & car_llave!CAR_FECHA_VCTO_ORIG
   xl.Cells(F1, 7) = "'" & llave_rep01!CAA_FECHA
   xl.Cells(F1, 8) = Val(llave_rep01!CAA_IMPORTE) * -1 ' Monto Pagado
   xl.Cells(F1, 9) = Val(car_llave!CAR_IMP_INI) 'deuda original
   xl.Cells(F1, 10) = Val(car_llave!CAR_IMPORTE) 'saldo actual
   xl.Cells(F1, 11) = Val(car_llave!CAR_COMISION) 'comision kardex
   WS_CALCULADA = (llave_rep01!CAA_IMPORTE * -1 * car_llave!CAR_COMISION) / (car_llave!CAR_IMP_INI)
   'WS_DIAS = DateDiff("d", LLAVE_REP01!CAA_FECHA, car_llave!CAR_FECHA_VCTO_ORIG)
   WS_DIAS = DateDiff("d", car_llave!CAR_FECHA_VCTO_ORIG, llave_rep01!CAA_FECHA)
   WS_COMI = POR_COMI(WS_DIAS)
   If WS_COMI = -99 Then
      GoTo CANCELA
   End If
   xl.Cells(F1, 12) = WS_CALCULADA 'comision calculada
   xl.Cells(F1, 13) = WS_COMI '% comision
   xl.Cells(F1, 14) = WS_CALCULADA * WS_COMI 'comision pagar
   xl.Cells(F1, 15) = WS_DIAS
otrito:
   llave_rep01.MoveNext
Loop
  wran1 = "H" & 7
  wran2 = "H" & F1
  wranF = "H" & F1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "I" & 7
  wran2 = "I" & F1
  wranF = "I" & F1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "J" & 7
  wran2 = "J" & F1
  wranF = "J" & F1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "N" & 7
  wran2 = "N" & F1
  wranF = "N" & F1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 
  wran1 = "A" & 7 & ":O" & F1
  xl.Application.Worksheets("HOJA DE COMISIONES x VENDEDOR").Range(wran1).Sort Key1:=xl.Application.Worksheets("HOJA DE COMISIONES x VENDEDOR").Range("O7")
  FrmImp2.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(2, 2) = "INFORME DE COMISIONES x VENDEDOR"
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
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
  Dim dd As Excel.Application
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblProceso.Caption = "Abriendo , Archivo Comisiones.xls . . . "
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open CONS_ADMIN & "OFFICE\Comisiones.xls", 0, True, 4, WPAS, WPAS

Return

Exit Sub
CANCELA:
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
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
 xl.Application.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
 
End Sub


Public Function POR_COMI(WDIAS As Integer) As Currency
Dim WINI As Integer
Dim WFIN As Integer
SQ_OPER = 2
PUB_TIPREG = 444
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
 WINI = Val(tab_mayor!tab_nomcorto)
 WFIN = Val(tab_mayor!TAB_CODART)
 If WFIN = 0 And WINI <> 0 Then
   If WDIAS >= WINI Then
    POR_COMI = Val(tab_mayor!tab_nomlargo)
    Exit Function
   End If
 ElseIf WDIAS >= WINI And WDIAS <= WFIN Then
   POR_COMI = Val(tab_mayor!tab_nomlargo)
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
If Right(txtCampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtCampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtCampo2.Text)
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

pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
'Dim PSCARR As rdoQuery
'Dim CARR As rdoResultset

Screen.MousePointer = 11
FrmImp2.lblProceso.Visible = True
FrmImp2.lblProceso.Caption = "Verificando Datos . . ."
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
FrmImp2.ProgBar.Max = CARR.RowCount
FrmImp2.ProgBar.Value = 0
DoEvents
FrmImp2.lblProceso.Visible = True
DoEvents
FrmImp2.lblProceso.Caption = "Procesando Datos . . ."
Screen.MousePointer = 11
FrmImp2.ProgBar.Visible = True
DoEvents
Do Until CARR.EOF
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   SQ_OPER = 1
   PU_TIPMOV = Nulo_Valor0(CARR!CAR_TIPMOV)
   PU_NUMSER = CARR!CAR_numSER
   PU_NUMFAC = CARR!car_numFAC
   pu_codcia = CARR!CAR_CODCIA
   PU_FBG = Nulo_Valors(CARR!CAR_fbg)
   LEER_FAR_LLAVE
   WS_COMISION = 0
   pu_codcia = LK_CODCIA
   Do Until far_llave.EOF = True
      PUB_KEY = far_llave!far_codart
      LEER_ART_LLAVE
      WS_IMPORTE = 0
      If far_llave!FAR_NUM_PRECIO = "1" Then
         WS_IMPORTE = redondea(art_LLAVE!ART_POR1 * far_llave!FAR_CANTIDAD * far_llave!far_PRECIO / (100 * far_llave!FAR_equiv))
      ElseIf far_llave!FAR_NUM_PRECIO = "2" Then
         WS_IMPORTE = redondea(art_LLAVE!ART_POR2 * far_llave!FAR_CANTIDAD * far_llave!far_PRECIO / (100 * far_llave!FAR_equiv))
      ElseIf far_llave!FAR_NUM_PRECIO = "3" Then
         WS_IMPORTE = redondea(art_LLAVE!ART_POR3 * far_llave!FAR_CANTIDAD * far_llave!far_PRECIO / (100 * far_llave!FAR_equiv))
      ElseIf far_llave!FAR_NUM_PRECIO = "4" Then
         WS_IMPORTE = redondea(art_LLAVE!ART_POR4 * far_llave!FAR_CANTIDAD * far_llave!far_PRECIO / (100 * far_llave!FAR_equiv))
      ElseIf far_llave!FAR_NUM_PRECIO = "5" Then
         WS_IMPORTE = redondea(art_LLAVE!ART_POR5 * far_llave!FAR_CANTIDAD * far_llave!far_PRECIO / (100 * far_llave!FAR_equiv))
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
 FrmImp2.lblProceso.Visible = False
 FrmImp2.ProgBar.Visible = False
 Screen.MousePointer = 0
 FrmImp2.pantalla.Enabled = True
 FrmImp2.cerrar.Enabled = True
 MsgBox "Proceso Terminado Correctamente ", 48, Pub_Titulo

Exit Sub
fin:
 FrmImp2.lblProceso.Visible = False
 FrmImp2.ProgBar.Visible = False
 Screen.MousePointer = 0
 FrmImp2.pantalla.Enabled = True
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
        cont.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!TAB_NUMTAB
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
        cont.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!TAB_NUMTAB
        tab_mayor.MoveNext
    Loop
End Sub

Public Sub HER_NEGOCIOS()
On Error GoTo FINTODO
Dim wRuta As String
Dim WMONTO As Currency
Dim WCODCLIE As Currency
Dim var_ACUTOT As Currency
Dim var_ACUATE As Currency
Dim var_ACUPED As Currency
Dim wnumfac As Currency
Dim ws_clave As String
Dim Wflag As String * 1
Dim wflag2 As String * 1
Dim wsFECHA1, wsFECHA2
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtCampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtCampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtCampo2.Text)
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

pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = "0"
usu.Requery
Do Until usu.EOF
  If LK_CODUSU = "ADMIN" And Trim(usu!usu_key) = "ADMIN" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  If Trim(usu!usu_key) = "SUPERVISOR" Then
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
FrmImp2.lblProceso.Visible = True
FrmImp2.lblProceso.Caption = "Activando Reporte. . . "
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
F1 = 5  'Fila Inicial
Wflag = ""
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Max = tab_mayor.RowCount
FrmImp2.lblProceso.Caption = "Procesando . . . "
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
      WCODCLIE = llave_rep02!CLI_CODCLIE
      GoSub FAR_RECORRE
      llave_rep02.MoveNext
    Loop
    If wnumfac <> -1 Then
     If Trim(Wflag) = "" Then
       GoSub WEXCEL
       FrmImp2.lblProceso.Caption = "Procesando . . . "
       DoEvents
       Wflag = "A"
     End If
     F1 = F1 + 1
     xl.Cells(F1, 1) = F1 - 5
     xl.Cells(F1, 2) = Trim(tab_mayor!tab_nomlargo)
     xl.Cells(F1, 3) = Format(var_ACUPED, "###")
     xl.Cells(F1, 4) = Format(var_ACUATE, "###")
     xl.Cells(F1, 5) = Format(var_ACUTOT, "##,##0.00")
   End If
 End If
 tab_mayor.MoveNext
Loop
 If Wflag <> "A" Then
   FrmImp2.lblProceso.Visible = False
   MsgBox "NO Existe Ventas ...", 48, Pub_Titulo
   GoTo CANCELA
 End If
  xl.Cells(1, 1) = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
  wranF = "B" & F1 + 1
  xl.Range(wranF) = "TOTAL "
  wran1 = "C" & 6
  wran2 = "C" & F1
  wranF = "C" & F1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "D" & 6
  wran2 = "D" & F1
  wranF = "D" & F1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "E" & 6
  wran2 = "E" & F1
  wranF = "E" & F1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  WMONTO = Val(xl.Range(wranF))
  wran1 = "F" & 6
  wran2 = "F" & F1
  wranF = "F" & F1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wranF = "A" & F1 + 1 & ":F" & F1 + 1
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  For fila = 1 To F1 - 5
    wran1 = "E" & fila + 5
    wranF = "F" & fila + 5
    If WMONTO <> 0 Then xl.Range(wranF).Formula = "=(" & wran1 & "* 100) /" & WMONTO
  Next fila
  FrmImp2.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(2, 1) = "LISTA DE ANALISIS DE VENTAS POR TIPO DE NEGOCIO"
  xl.Cells(3, 1) = "'" & LK_FECHA_DIA
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
Exit Sub
FAR_RECORRE:
    PS_REP01(0) = WCODCLIE
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
        var_ACUTOT = var_ACUTOT + (llave_rep01!far_PRECIO * llave_rep01!FAR_CANTIDAD) / llave_rep01!FAR_equiv
       End If
       llave_rep01.MoveNext
    Loop
     
dale_otro:
Return

WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  wRuta = PUB_RUTA_OTRO
  wsfile1 = PUB_RUTA_REPORTE + "hernego.xls"
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblProceso.Caption = "Abriendo , Archivo hernego.xls . . . "
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open wsfile1, 0, True, 4, WPAS, WPAS

Return

Exit Sub
CANCELA:
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
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
Dim WMONTO As Currency
Dim WCODCLIE As Currency
Dim var_ACUTOT As Currency
Dim var_ACUATE As Currency
Dim var_ACUPED As Currency
Dim wnumfac As Currency
Dim ws_clave As String
Dim Wflag As String * 1
Dim wflag2 As String * 1
Dim wsFECHA1, wsFECHA2
Dim xcuenta
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtCampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtCampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtCampo2.Text)
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

pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE

pub_cadena = "SELECT CLI_CODCLIE FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_GRUPO = ? "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

'pub_cadena = "SELECT SUM(CAR_IMPORTE)AS DEUDA, CAR_CODCLIE FROM CARTERA WHERE CAR_CODCIA = '02' AND CAR_CP = 'C' AND CAR_TIPDOC <> 'CH' AND CAR_IMPORTE <> 0  GROUP BY CAR_CODCLIE "
'Set PS_CON03 = CN.CreateQuery("", pub_cadena)
'Set llave_con03 = PS_CON03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


'pub_cadena = "SELECT COUNT(FAR_NUMFAC) AS PEDIDOS, SUM(FAR_CANTIDAD * FAR_PRECIO) AS TOTAL   FROM FACART WHERE FAR_CODCLIE = ? AND FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_ESTADO <> 'E' GROUP BY FAR_FBG, FAR_NUMSER,FAR_NUMFAC"
pub_cadena = "SELECT FAR_CODCLIE, COUNT(FAR_NUMFAC) AS PEDIDOS, SUM(FAR_CANTIDAD * FAR_PRECIO) AS TOTAL  FROM FACART, CLIENTES, TABLAS  WHERE (FAR_CODCLIE = CLI_CODCLIE AND CLI_GRUPO = TAB_NUMTAB AND TAB_TIPREG = 222) AND CLI_GRUPO = ? AND  FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_ESTADO <> 'E' GROUP BY FAR_CODCLIE,FAR_FBG, FAR_NUMSER,FAR_NUMFAC"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT SUM(FAR_CANTIDAD * FAR_PRECIO) AS TOTAL  FROM FACART WHERE FAR_CODCLIE = ? AND FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_ESTADO <> 'E' GROUP BY FAR_FBG, FAR_NUMSER,FAR_NUMFAC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
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
FrmImp2.lblProceso.Visible = True
FrmImp2.lblProceso.Caption = "Activando Reporte. . . "
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
F1 = 5  'Fila Inicial
Wflag = ""
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Max = tab_mayor.RowCount
FrmImp2.lblProceso.Caption = "Procesando . . . "
GoSub WEXCEL
FrmImp2.lblProceso.Caption = "Procesando . . . "
Wflag = "A"
Do Until tab_mayor.EOF
    FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
    PS_REP01(0) = Nulo_Valor0(tab_mayor!TAB_NUMTAB)
    DoEvents
    var_ACUTOT = 0
    var_ACUATE = 0
    var_ACUPED = 0
    llave_rep01.Requery
    If Not llave_rep01.EOF Then
'      xl.Application.Visible = True
       GoSub FAR_RECORRE
       F1 = F1 + 1
       xl.Cells(F1, 1) = F1 - 5
       xl.Cells(F1, 2) = Trim(tab_mayor!tab_nomlargo)
       xl.Cells(F1, 3) = Format(var_ACUPED, "###")
       xl.Cells(F1, 4) = Format(var_ACUATE, "###")
       xl.Cells(F1, 5) = Format(var_ACUTOT, "##,##0.00")
       Wflag = ""
    End If
 tab_mayor.MoveNext
Loop
 If Wflag = "A" Then
   FrmImp2.lblProceso.Visible = False
   MsgBox "NO Existe Ventas ...", 48, Pub_Titulo
   GoTo CANCELA
 End If
  xl.Cells(1, 1) = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
  wranF = "B" & F1 + 1
  xl.Range(wranF) = "TOTAL "
  wran1 = "C" & 6
  wran2 = "C" & F1
  wranF = "C" & F1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "D" & 6
  wran2 = "D" & F1
  wranF = "D" & F1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "E" & 6
  wran2 = "E" & F1
  wranF = "E" & F1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  WMONTO = Val(xl.Range(wranF))
  wran1 = "F" & 6
  wran2 = "F" & F1
  wranF = "F" & F1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wranF = "A" & F1 + 1 & ":F" & F1 + 1
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  For fila = 1 To F1 - 5
    wran1 = "E" & fila + 5
    wranF = "F" & fila + 5
    If WMONTO <> 0 Then xl.Range(wranF).Formula = "=(" & wran1 & "* 100) /" & WMONTO
  Next fila
  FrmImp2.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(2, 1) = "LISTA DE ANALISIS DE VENTAS POR TIPO DE NEGOCIO"
  xl.Cells(3, 1) = "'" & LK_FECHA_DIA
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
Exit Sub
FAR_RECORRE:
'llave_rep01.Requery
    wnumfac = Val(llave_rep01!PEDIDOS)
    wnumfac = -1
    WPEDIDO = 0
    xcuenta = 0
    WPEDIDO = Val(llave_rep01.RowCount)
    Do Until llave_rep01.EOF
      If wnumfac <> llave_rep01!far_codclie Then
        xcuenta = xcuenta + 1
        wnumfac = llave_rep01!far_codclie
        PS_REP02(0) = llave_rep01!far_codclie
        llave_rep02.Requery
        Do Until llave_rep02.EOF
           var_ACUTOT = var_ACUTOT + Val(llave_rep02!total)
           llave_rep02.MoveNext
        Loop
      End If
      llave_rep01.MoveNext
   Loop
   var_ACUATE = xcuenta
   var_ACUPED = WPEDIDO
Return

WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  wRuta = PUB_RUTA_OTRO
  wsfile1 = PUB_RUTA_REPORTE + "hernego.xls"
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblProceso.Caption = "Abriendo , Archivo hernego.xls . . . "
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open wsfile1, 0, True, 4, WPAS, WPAS

Return

Exit Sub
CANCELA:
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
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
lblProceso.Visible = True
pantalla.Enabled = False
cerrar.Enabled = False
If tra_llave!tra_rep1 = "1" Then
  If LK_EMP_PTO = "A" Then
    Reportes.ReportFileName = CONS_ADMIN & "STANDAR\PTOVTA\" & "hstock.rpt"
    wcodcia = "00"
  Else
    Reportes.ReportFileName = PUB_RUTA_REPORTE + "hstock.rpt"
    wcodcia = LK_CODCIA
  End If
Else
   Reportes.ReportFileName = PUB_RUTA_REPORTE + "hstock.rpt"
   wcodcia = LK_CODCIA
End If
Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  " & Trim(tra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
DoEvents
ProgBar.Min = 0
ProgBar.Max = 7
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
    wkSELECT = Str(Val(Right(fami.Text, 6)))
    wfiltra1 = wfiltra1 + wkSELECT + ","
    Modo1 = Modo1 + wkSELECT + ","
  End If
Next fila
If Wche <> 0 Then
  CADENITA = Left(Modo1, Len(Modo1) - 1) & "] "
  wfiltra1 = Left(wfiltra1, Len(wfiltra1) - 1)
Else
  ProgBar.Visible = False
  lblProceso.Visible = False
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
wformula2 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
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
lblProceso.Visible = False
pantalla.Enabled = True
cerrar.Enabled = True
ProgBar.Visible = False
Exit Sub
procancela:
MsgBox Err.Description, 48, Pub_Titulo
Unload FrmImp2
Exit Sub
Cancel:
ProgBar.Visible = False
lblProceso.Visible = False
pantalla.Enabled = True
cerrar.Enabled = True

End Sub
Public Sub HER_VENDEDOR()
Dim CADENITA, wformula, wformula1, wformula2, wformula3, wformula4
Dim Modo, Modo1
Dim Wche, wkSELECT, DIA, MES, ano, DIA1, MES1, ANO1
Dim wfecha, wfiltra1
Dim wsFECHA1 As String
Dim wsFECHA2 As String
pub_cadena = ""
lblProceso.Visible = True
pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtCampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtCampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtCampo2.Text)
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
    Modo1 = Modo1 + Str(Val(Left(vendmulti.Text, 3))) + ","
  End If
Next fila
If Wche <> 0 Then
 pub_cadena = pub_cadena + Left(Modo1, Len(Modo1) - 1) & "] AND "
Else
pub_cadena = ""
End If

DIA = Day(wsFECHA1)
MES = Month(wsFECHA1)
ano = Year(wsFECHA1)
DIA1 = Day(wsFECHA2)
MES1 = Month(wsFECHA2)
ANO1 = Year(wsFECHA2)
pub_cadena = pub_cadena + "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 10 AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_FECHA} >= Date ( " & ano & "," & MES & "," & DIA & ") AND {FACART.FAR_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  " & Trim(tra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
DoEvents
ProgBar.Min = 0
ProgBar.Max = 7
ProgBar.Value = 0
ProgBar.Visible = True
ProgBar.Value = ProgBar.Value + 1
wsFECHA1 = txtCampo1.Text
wsFECHA2 = txtCampo2.Text
wfecha = "DEL " & wsFECHA1 & " AL " & wsFECHA2
ProgBar.Value = ProgBar.Value + 1
Reportes.Formulas(0) = ""
Reportes.Formulas(1) = ""
Reportes.Formulas(2) = ""
Reportes.ReportFileName = PUB_RUTA_REPORTE + "hvend.rpt"
ProgBar.Value = ProgBar.Value + 1
DoEvents
wformula1 = "TITULO=  'VENTAS ACUMULADAS x VENDEDOR '"
wformula2 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
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
lblProceso.Visible = False
pantalla.Enabled = True
cerrar.Enabled = True
ProgBar.Visible = False
Exit Sub
procancela:
MsgBox Err.Description, 48, Pub_Titulo
Unload FrmImp2
Exit Sub
Cancel:
ProgBar.Visible = False
lblProceso.Visible = False
pantalla.Enabled = True
cerrar.Enabled = True
Exit Sub
SALE:
 If Err.Number = 20504 Then
   MsgBox "No esta el informe : " & Reportes.ReportFileName, 48, Pub_Titulo
 Else
   MsgBox "Fechas Invalidas .. intente nuevamente .. ", 48, Pub_Titulo
  
 End If
 Azul2 txtCampo1, txtCampo1
 lblProceso.Visible = False
 pantalla.Enabled = True
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
 pantalla.SetFocus
End If
End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If pantalla.Enabled Then pantalla.SetFocus
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

Dim LETRAS(48) As String * 2
Dim Wche As Integer
Dim wRuta As String
Dim WMONTO As Currency
Dim WCODCLIE As Currency
Dim var_ACUTOT As Currency
Dim var_ACUATE As Currency
Dim var_ACUPED As Currency
Dim wnumfac As Currency
Dim ws_clave As String
Dim Wflag As String * 1
Dim wsFECHA1, wsFECHA2
Dim Acodven() As Currency
Dim Modo1 As String
Dim wcodven As Currency
Dim xcuenta As Integer
pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtCampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtCampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtCampo2.Text)
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

pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE

pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = '" & LK_CODCIA & "'"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM FACART WHERE FAR_CODVEN = ? AND FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_ESTADO <> 'E' ORDER BY FAR_CODCIA, FAR_FECHA, FAR_NUMSER, FAR_NUMFAC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


' el PS_REP1(0) ESTA MAS ABAJO
PS_REP02(1) = LK_CODCIA
PS_REP02(2) = 10
PS_REP02(3) = wsFECHA1
PS_REP02(4) = wsFECHA2
DoEvents
FrmImp2.lblProceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep01.Requery
If llave_rep01.EOF Then
  GoTo CANCELA
End If
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Max = llave_rep01.RowCount
F1 = 5  'Fila Inicial
Wflag = ""
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
  F1 = F1 + 1
  xl.Cells(F1, 1) = Trim(Format(llave_rep01!vem_codven, "000")) & " " & Trim(llave_rep01!VEM_NOMBRE)
  PS_REP02(0) = llave_rep01!vem_codven
  llave_rep02.Requery
  If llave_rep02.EOF Then
     GoTo otroVEN
  End If
  C1 = 1
  tab_mayor.MoveFirst
  Do Until tab_mayor.EOF
    C1 = C1 + 1
    If Trim(Wflag) = "" Then
     xl.Cells(5, C1) = Left(tab_mayor!tab_nomlargo, 8)
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
    xl.Cells(F1, C1) = Format(xcuenta, "###")
    tab_mayor.MoveNext
  Loop
  Wflag = "A"
otroVEN:
 llave_rep01.MoveNext
Loop
  GoSub LETRAS
  xcuenta = C1 + 1
  For fila = 6 To F1
    wranF = Trim(LETRAS(xcuenta)) & fila
    wran1 = Trim(LETRAS(2)) & fila
    wran2 = Trim(LETRAS(C1)) & fila
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  Next
  For fila = 2 To C1 + 1
    wranF = Trim(LETRAS(fila)) & F1 + 1
    wran1 = Trim(LETRAS(fila)) & 6
    wran2 = Trim(LETRAS(fila)) & F1
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  Next
  wranF = "A" & F1 + 1
  xl.Range(wranF) = "Totales :"
  wranF = Trim(LETRAS(C1 + 1)) & 5
  xl.Range(wranF) = "Total"
  wranF = "A" & F1 + 1 & ":" & Trim(LETRAS(C1 + 1)) & F1 + 1
  xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  If Chesup.Value = 1 Then
    FrmImp2.lblProceso.Caption = "Suprimiendo  0 ..."
    fila = 1
    Do Until fila >= C1 + 1
     fila = fila + 1
     wranF = Trim(LETRAS(fila)) & F1 + 1
     If Val(xl.Range(wranF).Value) = 0 Then
         xl.Range(wranF).Delete 4
         fila = fila - 1
         C1 = C1 - 1
     End If
     Loop
  End If
  FrmImp2.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
  xl.Cells(2, 1) = Trim(tra_llave!tra_descripcion)
  xl.Cells(3, 1) = "'" & LK_FECHA_DIA
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
Exit Sub

CANCELA:
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblProceso.Caption = "Abriendo , Archivo Comisiones.xls . . . "
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open PUB_RUTA_REPORTE + "VENXNEGO.xls", 0, True, 4, WPAS, WPAS
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
 MsgBox "Reintente Nuevamente ..", 48, Pub_Titulo
 GoTo CANCELA
End Sub

Public Sub VENTA_X_VEND()
'On Error GoTo FINTODO
Dim LETRAS(48) As String * 2
Dim Wche As Integer
Dim wfami As Integer
Dim WCODCLIE As Currency
Dim wcodven As Currency
Dim wnumfac As Currency
Dim ws_clave As String
Dim Wflag As String * 1
Dim wsFECHA1, wsFECHA2
Dim Modo1 As String
Dim xcuenta As Integer
Dim CADENA As String
Dim TOT_VEN As Integer
Dim CANTIDAD_VEN, wcantidad As Currency
Dim wtot_soles As Currency
pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtCampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtCampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtCampo2.Text)
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
    Modo1 = Modo1 + Str(Val(Left(vendmulti.Text, 3))) + " OR FAR_CODVEN = "
    xcuenta = xcuenta + 1
    WCOLUNNA(xcuenta) = Val(Left(vendmulti.Text, 3))
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
 pub_cadena = "SELECT * FROM FACART WHERE FAR_CODART = ? AND FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_ESTADO <> 'E' AND (" & Modo1 & ") ORDER BY FAR_CODCIA, FAR_CODVEN , FAR_FECHA, FAR_NUMSER, FAR_NUMFAC"
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

pantalla.Enabled = False
cerrar.Enabled = False
GoSub WEXCEL
DoEvents
FrmImp2.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
For fila = 1 To TOT_VEN
  xl.Cells(5, fila + 3) = "V-" & Format(WCOLUNNA(fila), "00")
Next fila
ws_clave = PUB_CLAVE

Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = '" & LK_CODCIA & "' AND (" & CADENA & ") AND ART_CALIDAD = 1 ORDER BY ART_CODCIA, ART_FAMILIA, ART_ALTERNO"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT PRE_UNIDAD,PRE_EQUIV FROM PRECIOS WHERE PRE_CODCIA = '" & LK_CODCIA & "' and PRE_CODART = ? AND PRE_FLAG_UNIDAD = 'A' ORDER BY PRE_CODART"
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


' el PS_REP1(0) ESTA MAS ABAJO
PS_REP02(1) = LK_CODCIA
PS_REP02(2) = 10
PS_REP02(3) = wsFECHA1
PS_REP02(4) = wsFECHA2
DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.lblProceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep01.Requery
If llave_rep01.EOF Then
  GoTo CANCELA
End If
Wflag = ""
FrmImp2.lblProceso.Caption = "Procesando... un Momento ."
DoEvents
FrmImp2.ProgBar.Max = llave_rep01.RowCount
F1 = 5  'Fila Inicial
SQ_OPER = 1
PUB_TIPREG = 122
PUB_CODCIA = LK_CODCIA
PUB_NUMTAB = llave_rep01!art_familia
LEER_TAB_LLAVE
If tab_llave.EOF = False Then F1 = F1 + 1: xl.Cells(F1, 1) = ">" & Trim(tab_llave!tab_nomlargo)
wfami = llave_rep01!art_familia
Do Until llave_rep01.EOF
  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  If wfami <> llave_rep01!art_familia Then
    PUB_NUMTAB = llave_rep01!art_familia
    LEER_TAB_LLAVE
    If tab_llave.EOF = False Then F1 = F1 + 1: xl.Cells(F1, 1) = ">" & Trim(tab_llave!tab_nomlargo)
    wfami = llave_rep01!art_familia
  End If
  wfami = llave_rep01!art_familia
  F1 = F1 + 1
  wtot_soles = 0
  xl.Cells(F1, 1) = Trim(llave_rep01!ART_ALTERNO)
  xl.Cells(F1, 2) = Trim(llave_rep01!ART_NOMBRE)
  PS_REP03(0) = llave_rep01!ART_KEY
  llave_rep03.Requery
  If llave_rep03.EOF Then GoTo CANCELA
  xl.Cells(F1, 3) = Trim(llave_rep03!pre_unidad)
  PS_REP02(0) = llave_rep01!ART_KEY
  llave_rep02.Requery
  If llave_rep02.EOF Then
     GoTo otroVEN
  End If
  Wflag = "A"
  wcodven = llave_rep02!FAR_codven
  xcuenta = 0
  wcantidad = 0
   Do Until llave_rep02.EOF
     If wcodven = llave_rep02!FAR_codven Then
        wcantidad = wcantidad + llave_rep02!FAR_CANTIDAD / llave_rep03!PRE_EQUIV
        If llave_rep02!FAR_equiv <> 0 And llave_rep02!FAR_DESCTO = 0 Then
          wtot_soles = wtot_soles + llave_rep02!far_PRECIO * (llave_rep02!FAR_CANTIDAD / llave_rep02!FAR_equiv)
        End If
        xcuenta = 0
     Else
        CANTIDAD_VEN = wcantidad
        GoSub PONE_CANTIDAD
        wcodven = llave_rep02!FAR_codven
        wcantidad = llave_rep02!FAR_CANTIDAD / llave_rep03!PRE_EQUIV
        If llave_rep02!FAR_equiv <> 0 And llave_rep02!FAR_DESCTO = 0 Then
          wtot_soles = wtot_soles + llave_rep02!far_PRECIO * (llave_rep02!FAR_CANTIDAD / llave_rep02!FAR_equiv)
        End If
        xcuenta = 1
     End If
     wcodven = llave_rep02!FAR_codven
     llave_rep02.MoveNext
   Loop
     CANTIDAD_VEN = wcantidad
     GoSub PONE_CANTIDAD
 
otroVEN:
 llave_rep01.MoveNext
Loop
  GoSub LETRAS
'    xl.Application.Visible = True
 If Wflag <> "A" Then
  MsgBox " No hay Información para el filtro", 48, Pub_Titulo
  GoTo salta
 End If
  xcuenta = TOT_VEN + 4
  For fila = 6 To F1
    wranF = Trim(LETRAS(xcuenta)) & fila
    wran1 = Trim(LETRAS(4)) & fila
    wran2 = Trim(LETRAS(TOT_VEN + 3)) & fila
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
    If Val(xl.Range(wranF).Value) = 0 Then xl.Range(wranF).Value = ""
  Next
  For fila = 4 To TOT_VEN + 4 + 1
    wranF = Trim(LETRAS(fila)) & F1 + 1
    wran1 = Trim(LETRAS(fila)) & 6
    wran2 = Trim(LETRAS(fila)) & F1
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  Next
  wranF = "A" & F1 + 1
  xl.Range(wranF) = "Totales :"
  wranF = Trim(LETRAS(TOT_VEN + 4)) & 5
  xl.Range(wranF) = "Total"
  wranF = Trim(LETRAS(TOT_VEN + 5)) & 5
  xl.Range(wranF) = "SOLES"
  wranF = "A" & F1 + 1 & ":" & Trim(LETRAS(TOT_VEN + 5)) & F1 + 1
  xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  If Chesup.Visible And Chesup.Value = 1 Then
    FrmImp2.lblProceso.Caption = "Suprimiendo  0 ..."
    fila = 3
    Do Until fila >= TOT_VEN + 4
     fila = fila + 1
     wranF = Trim(LETRAS(fila)) & F1 + 1
     If Val(xl.Range(wranF).Value) = 0 Then
         xl.Range(wranF).Delete 4
         fila = fila - 1
         TOT_VEN = TOT_VEN - 1
     End If
     Loop
  End If
salta:
  FrmImp2.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
  xl.Cells(2, 1) = Trim(tra_llave!tra_descripcion)
  xl.Cells(3, 1) = "'" & "DEL " & wsFECHA1 & " AL " & wsFECHA2
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
Exit Sub
PONE_CANTIDAD:
For fila = 1 To TOT_VEN
  If WCOLUNNA(fila) = wcodven Then
    xl.Cells(F1, fila + 3) = Format(CANTIDAD_VEN, "####")
    Exit For
  End If
Next fila
xl.Cells(F1, TOT_VEN + 5) = Format(wtot_soles, "0.00")


Return

CANCELA:
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblProceso.Caption = "Abriendo , Archivo Comisiones.xls . . . "
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open PUB_RUTA_REPORTE + "VENTAXVEN.xls", 0, True, 4, WPAS, WPAS
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
 xl.Application.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
 
End Sub
Public Sub VEN_NEGOCIOS_CLI()
On Error GoTo FINTODO
Dim LETRAS(48) As String * 2
Dim Wche As Integer
Dim wRuta As String
Dim WMONTO As Currency
Dim WCODCLIE As Currency
Dim var_ACUTOT As Currency
Dim var_ACUATE As Currency
Dim var_ACUPED As Currency
Dim WCODCLI As Currency
Dim ws_clave As String
Dim Wflag As String * 1
Dim wsFECHA1, wsFECHA2
Dim Acodven() As Currency
Dim Modo1 As String
Dim wcodven As Currency
Dim xcuenta As Integer
pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtCampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtCampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtCampo2.Text)
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
    Modo1 = Modo1 + Str(Val(Left(vendmulti.Text, 3))) + " OR VEM_CODVEN = "
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

pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = "0"
usu.Requery
Do Until usu.EOF
  If LK_CODUSU = "ADMIN" And Trim(usu!usu_key) = "ADMIN" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  If Trim(usu!usu_key) = "SUPERVISOR" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  usu.MoveNext
Loop

Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM FACART WHERE FAR_CODVEN = ? AND FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_ESTADO <> 'E' ORDER BY FAR_CODCIA, FAR_CODCLIE " ', FAR_FECHA, FAR_NUMSER, FAR_NUMFAC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

' el PS_REP1(0) ESTA MAS ABAJO
PS_REP02(1) = LK_CODCIA
PS_REP02(2) = 10
PS_REP02(3) = wsFECHA1
PS_REP02(4) = wsFECHA2
DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.lblProceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep01.Requery
If llave_rep01.EOF Then
  GoTo CANCELA
End If
FrmImp2.lblProceso.Caption = "Procesando... un Momento ."
DoEvents
FrmImp2.ProgBar.Max = llave_rep01.RowCount
F1 = 5  'Fila Inicial
Wflag = ""
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
  F1 = F1 + 1
  xl.Cells(F1, 1) = Trim(llave_rep01!VEM_NOMBRE)
  PS_REP02(0) = llave_rep01!vem_codven
  llave_rep02.Requery
  If llave_rep02.EOF Then
     GoTo otroVEN
  End If
  C1 = 1
  tab_mayor.MoveFirst
  Do Until tab_mayor.EOF
    C1 = C1 + 1
    If Trim(Wflag) = "" Then xl.Cells(5, C1) = Left(tab_mayor!tab_nomlargo, 8)
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
    xl.Cells(F1, C1) = Format(xcuenta, "###")
    tab_mayor.MoveNext
  Loop
  Wflag = "A"
otroVEN:
 llave_rep01.MoveNext
Loop
  GoSub LETRAS
  xcuenta = C1 + 1
  For fila = 6 To F1
    wranF = Trim(LETRAS(xcuenta)) & fila
    wran1 = Trim(LETRAS(2)) & fila
    wran2 = Trim(LETRAS(C1)) & fila
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  Next
  For fila = 2 To C1 + 1
    wranF = Trim(LETRAS(fila)) & F1 + 1
    wran1 = Trim(LETRAS(fila)) & 6
    wran2 = Trim(LETRAS(fila)) & F1
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  Next
  wranF = "A" & F1 + 1
  xl.Range(wranF) = "Totales :"
  wranF = Trim(LETRAS(C1 + 1)) & 5
  xl.Range(wranF) = "Total"
  wranF = "A" & F1 + 1 & ":" & Trim(LETRAS(C1 + 4)) & F1 + 1
  xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  If Chesup.Value = 1 Then
    FrmImp2.lblProceso.Caption = "Suprimiendo  0 ..."
    fila = 1
    Do Until fila >= C1 + 1
     fila = fila + 1
     wranF = Trim(LETRAS(fila)) & F1 + 1
     If Val(xl.Range(wranF).Value) = 0 Then
         xl.Range(wranF).Delete 4
         fila = fila - 1
         C1 = C1 - 1
     End If
     Loop
  End If
  FrmImp2.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
  xl.Cells(2, 1) = Trim(tra_llave!tra_descripcion)
  xl.Cells(3, 1) = "'" & "DEL " & wsFECHA1 & " AL " & wsFECHA2
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
Exit Sub

CANCELA:
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblProceso.Caption = "Abriendo , Archivo Comisiones.xls . . . "
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open PUB_RUTA_REPORTE + "VENXNEGO.xls", 0, True, 4, WPAS, WPAS
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
 MsgBox "Verificar , Reintente  Nuevamente ..", 48, Pub_Titulo
 xl.DisplayAlerts = False
 xl.Application.Visible = True
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
lblProceso.Visible = True
pantalla.Enabled = False
cerrar.Enabled = False

Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  " & Trim(tra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
DoEvents
ProgBar.Min = 0
ProgBar.Max = 7
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
    wkSELECT = Str(Val(Right(fami.Text, 6)))
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
     Reportes.ReportFileName = CONS_ADMIN & "STANDAR\PTOVTA\" & "STOCKSEM.rpt"
   Else
     Reportes.ReportFileName = PUB_RUTA_REPORTE + "STOCKSEM.rpt"
   End If
ElseIf Wfile = "STOCKVA" Then
 Reportes.ReportFileName = PUB_RUTA_REPORTE + "STOCKVA.rpt"
End If
ProgBar.Value = ProgBar.Value + 1
DoEvents
'wformula1 = "FECHA=  '" & wFecha & "'"
'wformula1 = "TITULO=  'LISTADO DE STOCK PARA TOMA DE INVENTARIO '"
wformula1 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
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
lblProceso.Visible = False
pantalla.Enabled = True
cerrar.Enabled = True
ProgBar.Visible = False
Exit Sub
SALE:
procancela:
MsgBox Err.Description, 48, Pub_Titulo
Unload FrmImp2
Exit Sub
Cancel:
ProgBar.Visible = False
lblProceso.Visible = False
pantalla.Enabled = True
cerrar.Enabled = True

End Sub

Public Sub ZONA_X_NEGO()
'On Error GoTo FINTODO
Dim LETRAS(48) As String * 2
Dim Wche As Integer
Dim wcodtipo As Currency
Dim WCODCLIE As Currency
Dim ws_clave As String
Dim Wflag As String * 1
Dim Modo1 As String
Dim xcuenta As Integer
Dim CADENA As String
Dim TOT_COLU As Integer
Dim CANTIDAD_CLI, wcantidad As Currency
Dim wtot As Currency
Dim cod As Integer
Dim wsolos As String * 1
pantalla.Enabled = False
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
 WCOLUNNA(xcuenta) = tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!TAB_NUMTAB
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
    Modo1 = Modo1 + Str(Val(Right(zonas.Text, 5))) + " OR TAB_NUMTAB = "
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

pantalla.Enabled = False
cerrar.Enabled = False
GoSub WEXCEL
DoEvents
FrmImp2.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
For fila = 1 To TOT_COLU
  xl.Cells(5, fila + 1) = Left(WCOLUNNA(fila), 8)
Next fila
ws_clave = "0"
usu.Requery
Do Until usu.EOF
  If LK_CODUSU = "ADMIN" And Trim(usu!usu_key) = "ADMIN" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  If Trim(usu!usu_key) = "SUPERVISOR" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  usu.MoveNext
Loop

DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.lblProceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep01.Requery
If llave_rep01.EOF Then
  GoTo CANCELA
End If
wtot = 0
Wflag = ""
FrmImp2.lblProceso.Caption = "Procesando... un Momento ."
DoEvents
FrmImp2.ProgBar.Max = llave_rep01.RowCount
F1 = 5  'Fila Inicial
pub_cadena = ""
Do Until llave_rep01.EOF
  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  F1 = F1 + 1
  xl.Cells(F1, 1) = Trim(llave_rep01!tab_nomlargo)
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
     WCODCLIE = llave_rep02!CLI_CODCLIE
     llave_rep02.MoveNext
     Wflag = "A"
   Loop
   CANTIDAD_CLI = wcantidad
   GoSub PONE_CANTIDAD
  
OTRO:
 llave_rep01.MoveNext
Loop
  GoSub LETRAS
 If Wflag <> "A" Then
  MsgBox " No hay Información para el filtro", 48, Pub_Titulo
  GoTo salta
 End If
 If wtot <> 0 Then MsgBox "Existe " & wtot & " Cliente(s) sin Tipo de Negocio, su(s) Codigo(s) son : " & pub_cadena & " . se recomienda relacionarlo.", 48, Pub_Titulo
  xcuenta = TOT_COLU + 2
  For fila = 6 To F1
    wranF = Trim(LETRAS(xcuenta)) & fila
    wran1 = Trim(LETRAS(2)) & fila
    wran2 = Trim(LETRAS(TOT_COLU + 1)) & fila
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
    If Val(xl.Range(wranF).Value) = 0 Then xl.Range(wranF).Value = ""
  Next
  For fila = 2 To TOT_COLU + 2
    wranF = Trim(LETRAS(fila)) & F1 + 1
    wran1 = Trim(LETRAS(fila)) & 6
    wran2 = Trim(LETRAS(fila)) & F1
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  Next
  wranF = "A" & F1 + 1
  xl.Range(wranF) = "Totales :"
  wranF = Trim(LETRAS(TOT_COLU + 2)) & 5
  xl.Range(wranF) = "Total"
  wranF = "A5"
  xl.Range(wranF) = Left(Trim(lblzonas.Caption), Len(lblzonas.Caption) - 2)

  wranF = "A" & F1 + 1 & ":" & Trim(LETRAS(TOT_COLU + 2)) & F1 + 1
  xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  If Chesup.Visible And Chesup.Value = 1 Then
    FrmImp2.lblProceso.Caption = "Suprimiendo  0 ..."
    fila = 1
    Do Until fila >= TOT_COLU + 2
     fila = fila + 1
     wranF = Trim(LETRAS(fila)) & F1 + 1
     If Val(xl.Range(wranF).Value) = 0 Then
         xl.Range(wranF).Delete 4
         fila = fila - 1
         TOT_COLU = TOT_COLU - 1
     End If
     Loop
  End If
salta:
  FrmImp2.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
  xl.Cells(2, 1) = Trim(tra_llave!tra_descripcion)
  xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
Exit Sub

PONE_CANTIDAD:
wsolos = ""
For fila = 1 To TOT_COLU
  If Trim(Right(WCOLUNNA(fila), 6)) = Trim(Str(wcodtipo)) Then
    xl.Cells(F1, fila + 1) = CANTIDAD_CLI
    wsolos = "A"
    Exit For
  End If
Next fila
If wsolos <> "A" Then wtot = wtot + 1: pub_cadena = pub_cadena + Str(WCODCLIE) + ", "
Return

CANCELA:
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblProceso.Caption = "Abriendo , Archivo Comisiones.xls . . . "
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open PUB_RUTA_REPORTE + "ZONAXNEGO.xls", 0, True, 4, WPAS, WPAS
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
 xl.Application.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
 
End Sub

Private Sub zonas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If pantalla.Enabled Then pantalla.SetFocus
End If
End Sub
Public Sub MOROSIDAD()
Dim CADENITA, wformula, wformula1, wformula2, wformula3, wformula4
Dim wformula5, wformula6, wformula7, wformula8
Dim Modo, Modo1
Dim Wche, wkSELECT
Dim wfecha, wfiltra1
Dim DIA, MES, ano
Dim valor As Integer
valor = 0
lblProceso.Visible = True
pantalla.Enabled = False
cerrar.Enabled = False

If Val(txtDias1(0).Text) > Val(txtDias1(1).Text) Then
 MsgBox "Rango Invalido ...", 48, Pub_Titulo
 Azul txtDias1(0), txtDias1(0)
 GoTo CANCELA
ElseIf Val(txtDias1(0).Text) < 0 Or Val(txtDias1(1).Text) < 0 Then
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
 If Val(txtDias2(0).Text) > Val(txtDias2(1).Text) Then
  MsgBox "Rango Invalido ...", 48, Pub_Titulo
  Azul txtDias2(0), txtDias2(0)
  GoTo CANCELA
 ElseIf Val(txtDias2(0).Text) < 0 Or Val(txtDias2(1).Text) < 0 Then
  MsgBox "Rango  No Procede ...", 48, Pub_Titulo
  Azul txtDias2(0), txtDias2(0)
  GoTo CANCELA
 End If
 valor = Val(txtDias2(1).Text)
Else
 If Val(txtDias2(0).Text) < 0 Then
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
    Modo1 = Modo1 + Str(Val(Left(vendmulti.Text, 3))) + ","
  End If
Next fila
If Wche <> 0 Then
 pub_cadena = pub_cadena + Left(Modo1, Len(Modo1) - 1) & "] "
Else
pub_cadena = ""
End If

Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  " & Trim(tra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
DoEvents
ProgBar.Min = 0
ProgBar.Max = 7
ProgBar.Value = 0
ProgBar.Visible = True
ProgBar.Value = ProgBar.Value + 1
wfecha = Format(LK_FECHA_DIA, "dd/mm/yyyy")
DIA = Day(wfecha)
MES = Month(wfecha)
ano = Year(wfecha)
If pub_cadena = "" Then
 pub_cadena = "{CARTERA.CAR_IMPORTE} >= 0.001 AND {CARTERA.CAR_TIPDOC} = 'FA' AND {CARTERA.CAR_NUMFAC} >= 0.001 AND {CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "'"
Else
 pub_cadena = pub_cadena + " AND {CARTERA.CAR_IMPORTE} >= 0.001 AND {CARTERA.CAR_TIPDOC} = 'FA' AND {CARTERA.CAR_NUMFAC} >= 0.001 AND {CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "'"
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
Reportes.ReportFileName = PUB_RUTA_REPORTE + "MOROSIDAD.rpt"
ProgBar.Value = ProgBar.Value + 1
DoEvents
wformula1 = "TITULO=  'MOROSIDAD X VENDEDORES'"
wformula2 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
wformula3 = "DIA=  '" & wfecha & "'"
wformula4 = "DIAS1.1=  " & txtDias1(0).Text
wformula5 = "DIAS1.2=  " & txtDias1(1).Text
wformula6 = "DIAS2.1=  " & txtDias2(0).Text
wformula7 = "DIAS2.2=  " & valor

ProgBar.Value = ProgBar.Value + 1
ProgBar.Value = ProgBar.Value + 1
wformula8 = "DIA_FECHA = Date ( " & ano & "," & MES & "," & DIA & ")"
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
lblProceso.Visible = False
pantalla.Enabled = True
cerrar.Enabled = True
ProgBar.Visible = False
Exit Sub
CANCELA:
ProgBar.Visible = False
lblProceso.Visible = False
pantalla.Enabled = True
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
If Right(txtfecha.Text, 2) = "__" Then
     wsFECHA1 = Left(txtfecha.Text, 8)
Else
     wsFECHA1 = Trim(txtfecha.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If


FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Visible = True
DoEvents
FrmImp2.lblProceso.Visible = True
FrmImp2.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL

xl.Worksheets(1).Activate
WS_SALDO = 0
xcuenta = 0
F1 = 5

WS_MONEDA = Left(cmdMoneda.Text, 1)

If WS_MONEDA = "S" Then
   xl.Cells(4, 2) = "MONEDA:" & " NUEVOS SOLES"
ElseIf WS_MONEDA = "D" Then
   xl.Cells(4, 2) = "MONEDA:" & " DOLLARES"
End If


PUB_FECHA = wsFECHA1
SQ_OPER = 1
pu_codcia = LK_CODCIA
LEER_ALL_LLAVE
If all_llave.EOF Then GoTo VAMOS
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Max = all_llave.RowCount
FrmImp2.ProgBar.Value = 0
F1 = F1 + 1
ws_conta = 0
If WS_MONEDA = "S" Then
     WS_SALDO = all_llave!ALL_IMPORTE
Else
     WS_SALDO = all_llave!ALL_IMPORTE_DOLL
End If
'MsgBox all_llave!all_codtra
xl.Cells(F1, 2) = "Saldo Anterior:"
xl.Cells(F1, 5) = WS_SALDO
all_llave.MoveNext
Do Until all_llave.EOF
   FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
   If all_llave!ALL_SIGNO_CAJA = 0 Then GoTo OTRO
   If all_llave!all_codCIA <> LK_CODCIA Then GoTo OTRO
   If LK_EMP = "HER" And all_llave!ALL_codtra = 2727 And all_llave!ALL_SIGNO_CAR = 1 Then GoTo OTRO
   If all_llave!all_flag_ext = "E" Or all_llave!all_flag_ext = "X" Then GoTo OTRO
   If all_llave!ALL_SIGNO_CAR = 0 And all_llave!all_TIPMOV = 0 Then
      WS_IMPORTE = all_llave!ALL_IMPORTE
   Else
      WS_IMPORTE = all_llave!all_Importe_AMORT
   End If
  
   If Trim(all_llave!ALL_moneda_CCM) <> " " And Val(all_llave!ALL_CODBAN) <> 0 Then
      ww_moneda = all_llave!ALL_moneda_CCM
   ElseIf Trim(all_llave!ALL_moneda_CLI) <> " " And Val(all_llave!ALL_CODCLIE) <> 0 Then
      ww_moneda = all_llave!ALL_moneda_CLI
   ElseIf Trim(all_llave!ALL_moneda_CAJA) <> " " Then
      ww_moneda = all_llave!ALL_moneda_CAJA
   End If
   If ww_moneda <> WS_MONEDA Then GoTo OTRO
   WS_NOMCLI = ""
   WS_NOMBAN = ""
   ws_mensaje = ""
   WS_FBG = ""
   WS_LARGO = ""
   ws_nomche = ""
   ws_codclie = Val(all_llave!ALL_CODCLIE)
   If all_llave!ALL_NUMFAC <> 0 Then
       If all_llave!ALL_FBG = "F" Then
          WS_FBG = "Fact. " & all_llave!ALL_NUMSER & "-" & all_llave!ALL_NUMFAC
       ElseIf all_llave!ALL_FBG = "B" Then
          WS_FBG = "Bolet." & all_llave!ALL_NUMSER & "-" & all_llave!ALL_NUMFAC
       Else
          WS_FBG = "Guia. " & all_llave!ALL_NUMSER & "-" & all_llave!ALL_NUMFAC
       End If
   End If
   If all_llave!ALL_CHENUM <> 0 And all_llave!ALL_SIGNO_CCM = -1 Then
       ws_nomche = "O/Pago: " & all_llave!ALL_CHENUM
   End If
      
   If ws_codclie <> 0 Then
         pu_cp = all_llave!ALL_cp
         pu_codcia = LK_CODCIA
         pu_codclie = all_llave!ALL_CODCLIE
         SQ_OPER = 1
         LEER_CLI_LLAVE
         WS_NOMCLI = Left(cli_llave!cli_nombre, 18) & ":"
   End If
   If Val(all_llave!ALL_CODBAN) <> 0 Then
         pu_codcia = LK_CODCIA
         PUB_CODBAN = all_llave!ALL_CODBAN
         SQ_OPER = 1
         LEER_CCM_LLAVE
         WS_NOMBAN = ccm_llave!CCM_nombre & ":"
   End If
   If all_llave!ALL_SIGNO_CAR <> 0 Then
       SQ_OPER = 1
       pu_cp = all_llave!ALL_cp
       pu_codclie = ws_codclie
       pu_codcia = LK_CODCIA
       PUB_SERDOC = Nulo_Valor0(all_llave!ALL_serdoc)
       PUB_NUMDOC = all_llave!ALL_NUMDOC
       PUB_TIPDOC = all_llave!ALL_tipdoc
       LEER_CAR_LLAVE
       WS_NOMCLI = Left(WS_NOMCLI, 30) & car_llave!CAR_fbg & "-." & car_llave!CAR_numSER & "-" & car_llave!car_numFAC
   End If
   WS_LARGO = Trim(WS_NOMCLI) & Trim(WS_FBG) & Trim(ws_mensaje) & Trim(WS_NOMBAN) & Trim(ws_nomche)
   If WS_LARGO = "" Then WS_LARGO = all_llave!ALL_CONCEPTO
      
   F1 = F1 + 1
   If all_llave!ALL_SIGNO_CAJA = 1 Then
      WS_SALDO = WS_SALDO + WS_IMPORTE
      WS_SALDO_ING = WS_SALDO_ING + WS_IMPORTE
      xl.Cells(F1, 3) = WS_IMPORTE
   Else
      WS_SALDO = WS_SALDO - WS_IMPORTE
      WS_SALDO_SAL = WS_SALDO_SAL + WS_IMPORTE
      xl.Cells(F1, 4) = WS_IMPORTE
   End If
   ws_conta = ws_conta + 1
   xl.Cells(F1, 1) = ws_conta
   xl.Cells(F1, 2) = all_llave!ALL_autocon 'WS_LARGO
   xl.Cells(F1, 5) = WS_SALDO
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
      tab_llave!tab_nomlargo = PUB_NUM_OPER
      tab_llave!tab_nomcorto = Format(LK_FECHA_DIA, "dd/mm/yyyy")
      tab_llave!TAB_CONTABLE2 = WS_SALDO
      tab_llave.Update
   End If
   
VAMOS:
wran1 = "C5"
wran2 = "C" & F1
xl.Visible = True
wranF = "C" & F1 + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
wran1 = "D5"
wran2 = "D" & F1
wranF = "D" & F1 + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
wranF = "C" & 5 & ":C" & F1
xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 3
wranF = "D" & 5 & ":D" & F1
xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 3
wranF = "E" & 5 & ":E" & F1
xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 3

FrmImp2.ProgBar.Value = todos_cli
Screen.MousePointer = 0
DoEvents
FrmImp2.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
DoEvents
xcuenta = 1
xl.Cells(2, 2) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
xl.Cells(2, 5) = "'" & Format(wsFECHA1, "dd mmm yyyy")
xl.DisplayAlerts = False
xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
xl.Worksheets("Hoja1").Protect PUB_CLAVE
xl.Application.Visible = True
DoEvents
FrmImp2.lblProceso.Visible = False
FrmImp2.ProgBar.Visible = False
Set xl = Nothing
Screen.MousePointer = 0
FrmImp2.pantalla.Enabled = True
FrmImp2.pantalla.Caption = "Por &Pantalla"
FrmImp2.lblProceso.Visible = False

Exit Sub


WEXCEL:
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FrmImp2.lblProceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open CONS_ADMIN & "OFFICE\CAJA.xls", 0, True, 4, WPAS, WPAS
Return

Exit Sub

'LLENA_VALOR:
'For I = QJ To WDIF
'  If I >= 3 Then
'    xl.Cells(F1 + 7, I) = Format(LOC_VALOR, "0.00")
'  End If
'Next I
'Return

Exit Sub
CANCELA:
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
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


Public Sub Saldo_Car()
Dim ws_file As String
FrmImp2.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
FrmImp2.lblProceso.Visible = True
FrmImp2.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
FILAX = 2
If Val(xl.Cells(FILAX, 5)) = 0 Then
  MsgBox "no hay datos para cargar.. todo empieza desde la fila 2 ", 48, Pub_Titulo
  GoTo CANCELA
End If
caa_histo.Requery

PUB_CP = InputBox("ingrese C=Clientes ,  P=Proveedores")
Do Until Val(xl.Cells(FILAX, 1)) = 0
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = PUB_CP
   pu_codclie = xl.Cells(FILAX, 5)
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
      MsgBox "Error en Codigo de cliente, NO EXISTE ...FILA:" & FILAX
      GoTo CANCELA
   End If
   SQ_OPER = 1
   pu_cp = PUB_CP
   PUB_CODVEN = xl.Cells(FILAX, 1)
   LEER_VEN_LLAVE
   If ven_llave.EOF Then
      MsgBox "Error en Codigo de Vendedor, NO EXISTE ...FILA:" & FILAX
      GoTo CANCELA
   End If
   
   If xl.Cells(FILAX, 2) = "F" Or xl.Cells(FILAX, 2) = "B" Or xl.Cells(FILAX, 2) = "G" Then
   Else
      MsgBox "Solo Válido F,B,G : en fila " & FILAX
      GoTo CANCELA
   End If
   
   If IsNumeric(xl.Cells(FILAX, 3)) = False Then
      MsgBox "Serie Invalido ...FILA:" & FILAX
      GoTo CANCELA
   End If
   If IsNumeric(xl.Cells(FILAX, 4)) = False Then
      MsgBox "N.Documento Invalido ...FILA:" & FILAX
      GoTo CANCELA
   End If
   
   If IsDate(xl.Cells(FILAX, 6)) = False Then
      MsgBox "Fecha Invalida ...FILA:" & FILAX
      GoTo CANCELA
   End If
   If IsDate(xl.Cells(FILAX, 7)) = False Then
      MsgBox "Fecha Invalida ...FILA:" & FILAX
      GoTo CANCELA
   End If
   If IsDate(xl.Cells(FILAX, 8)) = False Then
      MsgBox "Fecha Invalida ...FILA:" & FILAX
      GoTo CANCELA
   End If
   If IsNumeric(xl.Cells(FILAX, 9)) = False Then
      MsgBox "Numero Invalido ...FILA:" & FILAX
      GoTo CANCELA
   End If
   If IsNumeric(xl.Cells(FILAX, 10)) = False Then
      MsgBox "Numero Invalido ...FILA:" & FILAX
      GoTo CANCELA
   End If
   FILAX = FILAX + 1
Loop

FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Max = FILAX
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Visible = True
DoEvents
PUB_NUM_OPER_XXX = 5000
FILAX = 2
Do Until Val(xl.Cells(FILAX, 1)) = 0

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
   PUB_NUMDOC = car_menor!CAR_NUMDOC + 1
End If

PUB_CODCLIE = xl.Cells(FILAX, 5)
PUB_TIPDOC = "FA"

PUB_CODVEN = xl.Cells(FILAX, 1)

car_llave.AddNew
car_llave!CAR_CODCLIE = PUB_CODCLIE
car_llave!CAR_CODCIA = LK_CODCIA
car_llave!car_numguia = 0
car_llave!CAR_TIPDOC = PUB_TIPDOC
car_llave!CAR_cp = PUB_CP
car_llave!CAR_SERDOC = PUB_SERDOC
car_llave!CAR_NUMDOC = PUB_NUMDOC
car_llave!CAR_FECHA_INGR = xl.Cells(FILAX, 6)
car_llave!CAR_FECHA_VCTO = xl.Cells(FILAX, 8)
car_llave!CAR_FECHA_VCTO_ORIG = xl.Cells(FILAX, 7)
car_llave!CAR_SITUACION = 0
car_llave!CAR_COMISION = 0
car_llave!CAR_NAT_JUR = ""
car_llave!CAR_NUM_REN = 0
car_llave!car_concepto = "Saldo Inicial "
car_llave!car_nombre_banco = " "
car_llave!CAR_NUM_CHEQUE = 0
car_llave!car_SIGNO_CAJA = 0
car_llave!car_numguia = 0
PUB_IMPORTE_AMORT = xl.Cells(FILAX, 10)
car_llave!CAR_IMP_INI = xl.Cells(FILAX, 9)
car_llave!CAR_IMPORTE = PUB_IMPORTE_AMORT
car_llave!car_codtra = 0
car_llave!car_PRECIO = 0
car_llave!CAR_signo_car = 1
car_llave!CAR_numSER = xl.Cells(FILAX, 3)
car_llave!car_numFAC = xl.Cells(FILAX, 4)
car_llave!CAR_TIPMOV = 10
car_llave!CAR_fbg = Left(xl.Cells(FILAX, 2), 1)
car_llave!CAR_CODVEN = xl.Cells(FILAX, 1)
car_llave!CAR_numser_C = 0
car_llave!CAR_numfac_C = 0
car_llave!CAR_codban = 0
car_llave.Update


   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = PUB_CP
   pu_codclie = xl.Cells(FILAX, 5)
   LEER_CLI_LLAVE


cli_llave.Edit
cli_llave!cli_SALDO = Nulo_Valor0(cli_llave!cli_SALDO) + PUB_IMPORTE_AMORT
cli_llave.Update

caa_histo.AddNew
caa_histo!CAA_CODCLIE = PUB_CODCLIE
caa_histo!CAA_CODCIA = LK_CODCIA
caa_histo!CAA_TIPDOC = PUB_TIPDOC
caa_histo!CAA_CP = PUB_CP
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
caa_histo!CAA_NUM_OPER = PUB_NUM_OPER_XXX
caa_histo!caa_serdoc = PUB_SERDOC
caa_histo!CAA_NUMDOC = PUB_NUMDOC
caa_histo!CAA_FECHA = LK_FECHA_DIA
caa_histo!CAA_FECHA_VCTO = xl.Cells(FILAX, 8)
caa_histo!caa_situacion = 0
caa_histo!caa_concepto = "Saldo Inicial"
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
caa_histo!CAa_numser = xl.Cells(FILAX, 3)
caa_histo!CAa_numfac = xl.Cells(FILAX, 4)
caa_histo!CAa_numser_C = 0
caa_histo!CAA_NUMFAC_C = 0
caa_histo!CAa_numGUIA = 0
caa_histo!CAA_FBG = Left(xl.Cells(FILAX, 2), 1)
caa_histo!CAA_CODVEN = xl.Cells(FILAX, 1)
caa_histo.Update
FILAX = FILAX + 1
Loop
MsgBox "PROCESO TERMINADO SATISFACTORIAMENTE..", 48, Pub_Titulo
GoTo CANCELA

WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  If xl Is Nothing Then
    DoEvents
    Set xl = CreateObject("Excel.Application")
    DoEvents
  End If
  lblProceso.Caption = "Abriendo , Archivo SALDO_CAR.xls . . . "
  DoEvents
  xl.Workbooks.Open CONS_ADMIN & "OFFICE\SALDO_CAR.xls", 0, True, 4, WPAS
  Return

Exit Sub
CANCELA:
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  xl.Application.Visible = True
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
 xl.Application.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImp2
 
End Sub

Public Sub Saldo_Car_TE()
Dim ws_file As String
pub_cadena = "SELECT CAA_NUM_OPER FROM CARACU WHERE CAA_CP = 'C'  AND CAA_CODCIA = ? ORDER BY CAA_NUM_OPER DESC"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT CAR_CODCLIE FROM CARTERA WHERE CAR_CP = 'C'  AND CAR_CODCIA = ? AND CAR_CODCLIE = ?  ORDER BY CAR_CODCLIE"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
 
FrmImp2.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
FrmImp2.lblProceso.Visible = True
FrmImp2.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
FILAX = 2
If Val(xl.Cells(FILAX, 3)) = 0 Then
  MsgBox "no hay datos para cargar.. todo empieza desde la fila 2 ", 48, Pub_Titulo
  GoTo CANCELA
End If
caa_histo.Requery

PUB_CP = "C"
SQ_OPER = 1
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = xl.Cells(FILAX, 1)
llave_rep02.Requery
If Not llave_rep02.EOF Then
  SQ_OPER = 1
  pu_codcia = LK_CODCIA
  pu_cp = PUB_CP
  pu_codclie = xl.Cells(FILAX, 1)
  LEER_CLI_LLAVE
  If cli_llave.EOF Then
      MsgBox "Error en Codigo de cliente, NO EXISTE ...FILA:" & FILAX
      GoTo CANCELA
  End If
  pub_mensaje = cli_llave!CLI_CODCLIE + " : " + cli_llave!cli_nombre + "YA TIENE SALDO EN CTA CTE. ¿Desea Continuar... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
     GoTo CANCELA
  End If
End If
Do Until Val(xl.Cells(FILAX, 1)) = 0
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = PUB_CP
   pu_codclie = xl.Cells(FILAX, 1)
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
      MsgBox "Error en Codigo de cliente, NO EXISTE ...FILA:" & FILAX
      GoTo CANCELA
   End If
   
   If IsNumeric(xl.Cells(FILAX, 2)) = False Then
      MsgBox "N.Documento Invalido ...FILA:" & FILAX
      GoTo CANCELA
   End If
   
   If IsDate(xl.Cells(FILAX, 4)) = False Then
      MsgBox "Fecha Invalida ...FILA:" & FILAX
      GoTo CANCELA
   End If
   If IsNumeric(xl.Cells(FILAX, 3)) = False Then
      MsgBox "Numero Invalido ...FILA:" & FILAX
      GoTo CANCELA
   End If
   
   FILAX = FILAX + 1
Loop

FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Max = FILAX
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Visible = True
DoEvents
PS_REP01(0) = LK_CODCIA
llave_rep01.Requery
PUB_NUM_OPER_XXX = llave_rep01!CAA_NUM_OPER + 1
FILAX = 2
Do Until Val(xl.Cells(FILAX, 1)) = 0
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
   PUB_NUMDOC = car_menor!CAR_NUMDOC + 1
End If

PUB_CODCLIE = xl.Cells(FILAX, 1)
PUB_TIPDOC = "TE"

PUB_CODVEN = 1
PUB_FECHA_VCTO = xl.Cells(FILAX, 4)
car_llave.AddNew
car_llave!CAR_CODCLIE = PUB_CODCLIE
car_llave!CAR_CODCIA = LK_CODCIA
car_llave!car_numguia = 0
car_llave!CAR_TIPDOC = PUB_TIPDOC
car_llave!CAR_cp = PUB_CP
car_llave!CAR_SERDOC = PUB_SERDOC
car_llave!CAR_NUMDOC = PUB_NUMDOC
car_llave!CAR_FECHA_INGR = LK_FECHA_DIA
car_llave!CAR_FECHA_VCTO = PUB_FECHA_VCTO
car_llave!CAR_FECHA_VCTO_ORIG = PUB_FECHA_VCTO
car_llave!CAR_SITUACION = 0
car_llave!CAR_COMISION = 0
car_llave!CAR_NUM_REN = 0
car_llave!car_concepto = "Saldo Inicial "
car_llave!car_nombre_banco = " "
car_llave!CAR_NUM_CHEQUE = 0
car_llave!car_SIGNO_CAJA = 0
car_llave!car_numguia = 0
PUB_IMPORTE_AMORT = xl.Cells(FILAX, 3)
car_llave!CAR_IMP_INI = xl.Cells(FILAX, 3)
car_llave!CAR_IMPORTE = PUB_IMPORTE_AMORT
car_llave!car_codtra = 0
car_llave!car_PRECIO = 0
car_llave!CAR_signo_car = 1
car_llave!CAR_numSER = 0
car_llave!car_numFAC = 0
car_llave!CAR_TIPMOV = 0
car_llave!CAR_fbg = ""
car_llave!CAR_CODVEN = 0
car_llave!CAR_numser_C = 0
car_llave!CAR_numfac_C = xl.Cells(FILAX, 2)
car_llave!CAR_codban = 0
car_llave.Update


   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = PUB_CP
   pu_codclie = xl.Cells(FILAX, 1)
   LEER_CLI_LLAVE


cli_llave.Edit
cli_llave!cli_SALDO = Nulo_Valor0(cli_llave!cli_SALDO) + PUB_IMPORTE_AMORT
cli_llave.Update

caa_histo.AddNew
caa_histo!CAA_CODCLIE = PUB_CODCLIE
caa_histo!CAA_CODCIA = LK_CODCIA
caa_histo!CAA_TIPDOC = PUB_TIPDOC
caa_histo!CAA_CP = PUB_CP
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
caa_histo!CAA_NUM_OPER = PUB_NUM_OPER_XXX
caa_histo!caa_serdoc = PUB_SERDOC
caa_histo!CAA_NUMDOC = PUB_NUMDOC
caa_histo!CAA_FECHA = LK_FECHA_DIA
caa_histo!CAA_FECHA_VCTO = PUB_FECHA_VCTO
caa_histo!caa_situacion = 0
caa_histo!caa_concepto = "Saldo Inicial"
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
caa_histo!CAa_numser_C = 0
caa_histo!CAA_NUMFAC_C = xl.Cells(FILAX, 2)
caa_histo!CAa_numser = 0
caa_histo!CAa_numfac = 0
caa_histo!CAa_NOMBRE = cli_llave!cli_nombre
caa_histo!CAa_numGUIA = 0
caa_histo!CAA_FBG = ""
caa_histo!CAA_CODVEN = 0
caa_histo.Update
FILAX = FILAX + 1
Loop
MsgBox "PROCESO TERMINADO SATISFACTORIAMENTE..", 48, Pub_Titulo
GoTo CANCELA

WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  If xl Is Nothing Then
    DoEvents
    Set xl = CreateObject("Excel.Application")
    DoEvents
  End If
  lblProceso.Caption = "Abriendo , Archivo SALDO_CAR.xls . . . "
  DoEvents
  xl.Workbooks.Open "C:\CARGAR.xls", 0, True, 4, WPAS
  Return

Exit Sub
CANCELA:
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  xl.Application.Visible = True
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
 xl.Application.Visible = True
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
Dim WMONTO As Currency
Dim WINGRESOS As Currency
Dim TOT_INGRESOS As Currency
Dim WSALDO_CAJA As Currency


If Right(txtfecha.Text, 2) = "__" Then
     wsFECHA1 = Left(txtfecha.Text, 8)
Else
     wsFECHA1 = Trim(txtfecha.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If


FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Visible = True
DoEvents
FrmImp2.lblProceso.Visible = True
FrmImp2.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL

xl.Worksheets(1).Activate
WS_SALDO = 0
xcuenta = 0
F1 = 5

WS_MONEDA = Left(cmdMoneda.Text, 1)

If WS_MONEDA = "S" Then
   xl.Cells(4, 2) = "MONEDA:" & " NUEVOS SOLES"
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
F1 = F1 + 1
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
PUB_CODCIA = "00"
SQ_OPER = 1
LEER_PAR_LLAVE
WS_SALDO = par_llave!PAR_SALDO_CAJA_ayer
WSALDO_CAJA = WSALDO_CAJA + WS_SALDO
xl.Cells(F1, 2) = "Saldo Anterior:"
xl.Cells(F1, 5) = WS_SALDO
'all_llave.MoveNext
pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ? AND ALL_FBG = ?  AND ALL_TIPMOV = 10 AND ALL_SIGNO_CAJA = 1  AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_DIA, ALL_FBG, ALL_NUMSER, ALL_NUMFAC"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT PAR_NOMBRE FROM PARGEN WHERE PAR_CODCIA = ? "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)

xcuenta = 1
F1 = 6
F1 = F1 + 1
xl.Cells(F1, 1) = "MAS  : INGRESOS"
F1 = F1 + 1
xl.Cells(F1, 1) = 1
xl.Cells(F1, 2) = "VENTAS AL CONTADO:"
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
    F1 = F1 + 1
    xl.Cells(F1, 2) = Trim(llave_rep02!PAR_NOMBRE)
    ' SOLO PARA FACTURAS
    PS_REP01(2) = "F"
    llave_rep01.Requery
    WMONTO = 0
    If Not llave_rep01.EOF Then
       docus = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!ALL_NUMFAC
      Do Until llave_rep01.EOF
        WMONTO = WMONTO + llave_rep01!all_Importe_AMORT
        docus2 = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!ALL_NUMFAC
        llave_rep01.MoveNext
      Loop
      F1 = F1 + 1
      ' xl.Cells(f1, 3) = "F/. " & docus & " AL " & docus2
      xl.Cells(F1, 3) = "FACTURAS"
      xl.Cells(F1, 4) = WMONTO
      WINGRESOS = WINGRESOS + WMONTO
    End If
    'SOLO PARA BOLETAS
    PS_REP01(2) = "B"
    llave_rep01.Requery
    WMONTO = 0
    If Not llave_rep01.EOF Then
       docus = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!ALL_NUMFAC
      Do Until llave_rep01.EOF
        WMONTO = WMONTO + llave_rep01!all_Importe_AMORT
        docus2 = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!ALL_NUMFAC
        llave_rep01.MoveNext
      Loop
      F1 = F1 + 1
      ' xl.Cells(f1, 3) = "B/. " & docus & " AL " & docus2
      xl.Cells(F1, 3) = "BOLETAS"
      xl.Cells(F1, 4) = WMONTO
      WINGRESOS = WINGRESOS + WMONTO
    End If
    ' SOLO PARA GUIAS
    PS_REP01(2) = "G"
    llave_rep01.Requery
    WMONTO = 0
    If Not llave_rep01.EOF Then
       docus = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!ALL_NUMFAC
      Do Until llave_rep01.EOF
        WMONTO = WMONTO + llave_rep01!all_Importe_AMORT
        docus2 = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!ALL_NUMFAC
        llave_rep01.MoveNext
      Loop
      F1 = F1 + 1
      'xl.Cells(f1, 3) = "G/. " & docus & " AL " & docus2
      xl.Cells(F1, 3) = "GUIAS"
      xl.Cells(F1, 4) = WMONTO
      WINGRESOS = WINGRESOS + WMONTO
    End If
    ' SOLO PARA VENTAS ADMINISTRADORES P
    PS_REP01(2) = "P"
    llave_rep01.Requery
    WMONTO = 0
    If Not llave_rep01.EOF Then
       docus = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!ALL_NUMFAC
      Do Until llave_rep01.EOF
        WMONTO = WMONTO + llave_rep01!all_Importe_AMORT
        docus2 = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!ALL_NUMFAC
        llave_rep01.MoveNext
      Loop
      F1 = F1 + 1
      xl.Cells(F1, 3) = "P/. " & docus & " AL " & docus2
      xl.Cells(F1, 3) = "ADMINISTRACION"
      xl.Cells(F1, 4) = WMONTO
      WINGRESOS = WINGRESOS + WMONTO
    End If
      F1 = F1 + 1
      xl.Cells(F1, 4) = WINGRESOS
      TOT_INGRESOS = TOT_INGRESOS + WINGRESOS
OTRA_CIA:
   xcuenta = xcuenta + 2
 Next fila
 F1 = F1 + 1
 xl.Cells(F1, 5) = TOT_INGRESOS
 WSALDO_CAJA = WSALDO_CAJA + TOT_INGRESOS
 
' SOLO COBRANZA DE FA,LE,AD
 F1 = F1 + 1
 pub_cadena = "SELECT ALL_IMPORTE_AMORT, ALL_CODCLIE, ALL_CODCIA FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ?  AND ALL_TIPDOC = ? AND ALL_SIGNO_CAJA = 1 AND ALL_SIGNO_CAR = -1 AND ALL_IMPORTE_AMORT <> 0  AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122) ORDER BY  ALL_FECHA_DIA, ALL_NUMOPER"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

 F1 = F1 + 1
 xl.Cells(F1, 1) = 2
 xl.Cells(F1, 2) = "COBRANZA DE FACTURAS"
 WINGRESOS = 0
 TOT_INGRESOS = 0
 
 SQ_OPER = 1
 pu_cp = "C"
 xcuenta = 1
 For fila = 1 To 30
   ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
   If Trim(ws_codcia) = "00" Then GoTo OTRA_CIA2
   If Trim(ws_codcia) = "" Then Exit For
    F1 = F1 + 1
    PS_REP02(0) = ws_codcia
    llave_rep02.Requery
    xl.Cells(F1, 2) = Trim(llave_rep02!PAR_NOMBRE)
    WINGRESOS = 0
    PS_REP01(0) = ws_codcia
    PS_REP01(1) = wsFECHA1
    PS_REP01(2) = "FA"
    llave_rep01.Requery
    F1 = F1 + 1
    xl.Cells(F1, 2) = ".- FACTURAS "
    WMONTO = 0
    If Not llave_rep01.EOF Then
      Do Until llave_rep01.EOF
        WMONTO = WMONTO + llave_rep01!all_Importe_AMORT
        pu_codclie = llave_rep01!ALL_CODCLIE
        pu_codcia = llave_rep01!all_codCIA
        LEER_CLI_LLAVE
        F1 = F1 + 1
        xl.Cells(F1, 3) = Trim(cli_llave!cli_nombre)
        xl.Cells(F1, 4) = Val(llave_rep01!all_Importe_AMORT)
        llave_rep01.MoveNext
      Loop
      WINGRESOS = WINGRESOS + WMONTO
    End If
    F1 = F1 + 1
    xl.Cells(F1, 4) = WINGRESOS
    TOT_INGRESOS = TOT_INGRESOS + WINGRESOS
    
    WINGRESOS = 0
    PS_REP01(0) = ws_codcia
    PS_REP01(1) = wsFECHA1
    PS_REP01(2) = "AD"
    llave_rep01.Requery
    F1 = F1 + 1
    xl.Cells(F1, 2) = ".- COBRANZA DE ADMINISTRACION "
    WMONTO = 0
    If Not llave_rep01.EOF Then
      Do Until llave_rep01.EOF
        WMONTO = WMONTO + llave_rep01!all_Importe_AMORT
        pu_codclie = llave_rep01!ALL_CODCLIE
        pu_codcia = llave_rep01!all_codCIA
        LEER_CLI_LLAVE
        F1 = F1 + 1
        xl.Cells(F1, 3) = Trim(cli_llave!cli_nombre)
        xl.Cells(F1, 4) = Val(llave_rep01!all_Importe_AMORT)
        llave_rep01.MoveNext
      Loop
      WINGRESOS = WINGRESOS + WMONTO
    End If
    F1 = F1 + 1
    xl.Cells(F1, 4) = WINGRESOS
    
    TOT_INGRESOS = TOT_INGRESOS + WINGRESOS
OTRA_CIA2:
   xcuenta = xcuenta + 2
 Next fila
F1 = F1 + 1
xl.Cells(F1, 5) = TOT_INGRESOS
WSALDO_CAJA = WSALDO_CAJA + TOT_INGRESOS

' SOLO COBRANZA DE ORDINARIAS Y JUDICIALES
 F1 = F1 + 1
 
 pub_cadena = "SELECT ALL_TIPDOC, ALL_SITUACION, ALL_IMPORTE_AMORT, ALL_CODCIA FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ?  AND (ALL_TIPDOC = ? OR ALL_TIPDOC = ? OR ALL_TIPDOC = ? OR ALL_TIPDOC = ? ) AND  ALL_SIGNO_CAJA = 1 AND ALL_SIGNO_CAR = -1  AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_DIA, ALL_NUMOPER"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

 F1 = F1 + 1
 xl.Cells(F1, 1) = 3
 xl.Cells(F1, 2) = "COBRANZA ORDINARIA"
 
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
    F1 = F1 + 1
    xl.Cells(F1, 2) = Trim(llave_rep02!PAR_NOMBRE)
    llave_rep01.Requery
    WMONTO = 0
    If Not llave_rep01.EOF Then
      Do Until llave_rep01.EOF
      If llave_rep01!ALL_tipdoc = "LE" And llave_rep01!ALL_SITUACION <> "P" Then
      Else
        WMONTO = WMONTO + llave_rep01!all_Importe_AMORT
       End If
        llave_rep01.MoveNext
      Loop
      WINGRESOS = WINGRESOS + WMONTO
    End If
    F1 = F1 + 1
    xl.Cells(F1, 4) = Val(WMONTO)
    TOT_INGRESOS = TOT_INGRESOS + WINGRESOS
OTRA_CIA3:
   xcuenta = xcuenta + 2
 Next fila
F1 = F1 + 1
xl.Cells(F1, 5) = TOT_INGRESOS
WSALDO_CAJA = WSALDO_CAJA + TOT_INGRESOS

F1 = F1 + 1

'COBRANZA JUDICIAL
 pub_cadena = "SELECT ALL_TIPDOC, ALL_SITUACION, ALL_IMPORTE_AMORT, ALL_CODCIA FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ?  AND (ALL_TIPDOC = ? OR ALL_TIPDOC = ? OR ALL_TIPDOC = ? ) AND  (ALL_SITUACION = ? OR ALL_SITUACION = ? ) AND ALL_SIGNO_CAJA = 1 AND ALL_SIGNO_CAR = -1  AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_DIA, ALL_NUMOPER"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

F1 = F1 + 1
 xl.Cells(F1, 1) = 3
 xl.Cells(F1, 2) = "COBRANZA JUDICIAL"
 
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
    F1 = F1 + 1
    xl.Cells(F1, 2) = Trim(llave_rep02!PAR_NOMBRE)
    llave_rep01.Requery
    WMONTO = 0
    If Not llave_rep01.EOF Then
      Do Until llave_rep01.EOF
        WMONTO = WMONTO + llave_rep01!all_Importe_AMORT
        llave_rep01.MoveNext
      Loop
      WINGRESOS = WINGRESOS + WMONTO
    End If
    F1 = F1 + 1
    xl.Cells(F1, 4) = Val(WMONTO)
    TOT_INGRESOS = TOT_INGRESOS + WINGRESOS
OTRA_CIAJ:
   xcuenta = xcuenta + 2
 Next fila
F1 = F1 + 1
xl.Cells(F1, 5) = TOT_INGRESOS
WSALDO_CAJA = WSALDO_CAJA + TOT_INGRESOS
 
' SOLO INGRESOS VARIOS
F1 = F1 + 1
 
 pub_cadena = "SELECT ALL_IMPORTE FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ?  AND ALL_CODTRA = 5350 AND ALL_SIGNO_CAJA = 1  AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_DIA, ALL_NUMOPER"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

 F1 = F1 + 1
 xl.Cells(F1, 1) = 4
 xl.Cells(F1, 2) = "INGRESOS VARIOS"
 
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
    F1 = F1 + 1
    xl.Cells(F1, 2) = Trim(llave_rep02!PAR_NOMBRE)
    llave_rep01.Requery
    WMONTO = 0
    If Not llave_rep01.EOF Then
      Do Until llave_rep01.EOF
        WMONTO = WMONTO + llave_rep01!ALL_IMPORTE
        llave_rep01.MoveNext
      Loop
      WINGRESOS = WINGRESOS + WMONTO
    End If
    F1 = F1 + 1
    xl.Cells(F1, 4) = Val(WMONTO)
    TOT_INGRESOS = TOT_INGRESOS + WINGRESOS
OTRA_CIA4:
   xcuenta = xcuenta + 2
 Next fila
F1 = F1 + 1
xl.Cells(F1, 5) = TOT_INGRESOS
WSALDO_CAJA = WSALDO_CAJA + TOT_INGRESOS
' SOLO depositos a bancos
F1 = F1 + 1
 
 pub_cadena = "SELECT ALL_IMPORTE FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ?  AND (ALL_CODTRA = 5310 ) AND ALL_SIGNO_CAJA = -1  AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_DIA, ALL_NUMOPER"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

 F1 = F1 + 1
 xl.Cells(F1, 1) = 5
 xl.Cells(F1, 2) = "MENOS : EGRESOS "
 F1 = F1 + 1
 xl.Cells(F1, 2) = "DEPOSITOS "
 
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
    F1 = F1 + 1
    xl.Cells(F1, 2) = Trim(llave_rep02!PAR_NOMBRE)
    llave_rep01.Requery
    WMONTO = 0
    If Not llave_rep01.EOF Then
      Do Until llave_rep01.EOF
        WMONTO = WMONTO + llave_rep01!ALL_IMPORTE
        llave_rep01.MoveNext
      Loop
      WINGRESOS = WINGRESOS + WMONTO
    End If
    F1 = F1 + 1
    xl.Cells(F1, 4) = Val(WMONTO)
    TOT_INGRESOS = TOT_INGRESOS + WINGRESOS
OTRA_CIA5:
   xcuenta = xcuenta + 2
 Next fila
F1 = F1 + 1
xl.Cells(F1, 5) = TOT_INGRESOS
WSALDO_CAJA = WSALDO_CAJA - TOT_INGRESOS '(DEPOSITOS)
F1 = F1 + 1
F1 = F1 + 1
xl.Cells(F1, 3) = "SALDO DE CAJA = "
xl.Cells(F1, 5) = WSALDO_CAJA
If LK_EMP = "DPV" Then
  PUB_CODCIA = "00"
  SQ_OPER = 1
  LEER_PAR_LLAVE
  par_llave.Edit
  par_llave!par_saldo_caja_hoy = WSALDO_CAJA
  par_llave.Update
  PUB_CODCIA = LK_CODCIA
  SQ_OPER = 1
  LEER_PAR_LLAVE
End If


DoEvents
FrmImp2.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
DoEvents
xcuenta = 1
xl.Cells(2, 2) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
xl.Cells(2, 5) = "'" & Format(wsFECHA1, "dd mmm yyyy")
xl.DisplayAlerts = False
xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
xl.Worksheets("Hoja1").Protect PUB_CLAVE
xl.Application.Visible = True
DoEvents
FrmImp2.lblProceso.Visible = False
FrmImp2.ProgBar.Visible = False
Set xl = Nothing
Screen.MousePointer = 0
FrmImp2.pantalla.Enabled = True
FrmImp2.pantalla.Caption = "Por &Pantalla"
FrmImp2.lblProceso.Visible = False

Exit Sub


WEXCEL:
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FrmImp2.lblProceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open CONS_ADMIN & "OFFICE\CAJA.xls", 0, True, 4, WPAS, WPAS
Return

Exit Sub

'LLENA_VALOR:
'For I = QJ To WDIF
'  If I >= 3 Then
'    xl.Cells(F1 + 7, I) = Format(LOC_VALOR, "0.00")
'  End If
'Next I
'Return

Exit Sub
CANCELA:
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
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
Dim WMONTO As Currency
Dim WCODCLIE As Currency
Dim valor_venta As Currency
Dim descto As Currency
Dim valor_igv As Currency
Dim valor_precio As Currency
Dim s_valor_venta As Currency
Dim s_descto As Currency
Dim s_valor_igv As Currency
Dim s_valor_precio As Currency

Dim t_valor_venta As Currency
Dim t_descto As Currency
Dim t_valor_igv As Currency
Dim t_valor_precio As Currency

Dim wnumfac As Currency
Dim Wflag As String * 1
Dim wsFECHA1, wsFECHA2
Dim xcuenta As Integer
Dim wq_fecha, wq_codclie, wq_codven, wq_docu, wq_nombre, wq_bruto, wq_gastos, wq_desto, wq_flete, wq_fbg, wq_serie
Dim wq_tot_descto, wq_impto, wq_total, wq_estado, wq_condi
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
var_secuencia = 0
pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtCampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtCampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtCampo2.Text)
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
If CDate(wsFECHA1) <> LK_FECHA_COP1 Then cheasiento.Value = 0
If CDate(wsFECHA2) <> LK_FECHA_COP2 Then cheasiento.Value = 0
GoSub WEXCEL
pub_cadena = ""
xcuenta = 0

pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
WCONTROL = 0
F1 = 5  'Fila Inicial
t_valor_venta = 0
t_descto = 0
t_valor_igv = 0
t_valor_precio = 0

'NCREDITO: ' empieza

'WCONTROL = WCONTROL + 1
WMONTO = 0
WCODCLIE = 0
valor_venta = 0
descto = 0
valor_igv = 0
valor_precio = 0
wnumfac = 0
Wflag = ""
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
s_valor_venta = 0
s_descto = 0
s_valor_igv = 0
s_valor_precio = 0

WCONTROL = WCONTROL + 1
If WCONTROL = 1 Then
  pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND (FAR_TIPMOV = 20 OR FAR_TIPMOV = 99)  AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND FAR_CP = 'P' AND FAR_ESTADO <> 'E' ORDER BY FAR_FBG,FAR_FECHA_COMPRA, FAR_NUMOPER"
  WS_SIGNO = 1
ElseIf WCONTROL = 2 Then
  pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = 97 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND FAR_FBG = 'C' AND FAR_ESTADO <> 'E' AND FAR_CP = 'P'  ORDER BY FAR_FECHA_COMPRA, FAR_NUMOPER"
  WS_SIGNO = -1
ElseIf WCONTROL = 3 Then
  pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = 98 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND FAR_FBG = 'A' and FAR_ESTADO <> 'E' AND FAR_CP = 'P' ORDER BY FAR_FECHA_COMPRA, FAR_NUMOPER"
  WS_SIGNO = 1
End If
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


' el PS_REP1(0) ESTA MAS ABAJO
PS_REP02(0) = LK_CODCIA
'PS_REP02(1) = 10
PS_REP02(1) = wsFECHA1
PS_REP02(2) = wsFECHA2
DoEvents
FrmImp2.lblProceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep02.Requery
If llave_rep02.EOF Then
  GoTo OTRO_DOCUMENTO
End If
FrmImp2.lblProceso.Caption = "Procesando . . . "
DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Max = llave_rep02.RowCount

Wflag = ""
SQ_OPER = 1
pu_cp = "P"
pu_codcia = LK_CODCIA
wnumfac = llave_rep02!far_numfac
wfecha = llave_rep02!far_fbg
wserie = llave_rep02!far_numser
wwwserie = llave_rep02!far_numser
wq_fecha = "01/01/1900"
xcuenta = 0
Wflag = "A"
wflag_numfac = "A"
AWQ_DESCTOS = 0
AWQ_GASTOS = 0
AWQ_FLETES = 0
www_fbg = Trim(llave_rep02!far_fbg)

Do Until llave_rep02.EOF
'  If llave_rep02!FAR_numfac = 401 Then Stop
  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  If wfecha <> llave_rep02!far_fbg Then '
     If wflag_numfac = "A" Then
       GoSub IMPRI_FAC
     End If
     wflag_numfac = ""
     F1 = F1 + 1
     GoSub TOTAL_DIA
     t_valor_venta = t_valor_venta + s_valor_venta
     t_descto = t_descto + s_descto
     t_valor_igv = t_valor_igv + s_valor_igv
     t_valor_precio = t_valor_precio + s_valor_precio
     wnumfac = llave_rep02!far_numfac
     wfecha = llave_rep02!far_fbg
     s_valor_venta = 0
     s_descto = 0
     s_valor_igv = 0
     s_valor_precio = 0
  End If
  
  If wnumfac <> llave_rep02!far_numfac Then
     GoSub IMPRI_FAC
     wflag_numfac = ""
     wnumfac = llave_rep02!far_numfac
  ElseIf Val(wwwserie) <> Val(llave_rep02!far_numser) And CDate(wq_fecha) = CDate(llave_rep02!FAR_fecha_compra) Then
    If Trim(www_fbg) <> Trim(llave_rep02!far_fbg) And wflag_numfac <> "" Then
      GoSub IMPRI_FAC
      wflag_numfac = ""
      wnumfac = llave_rep02!far_numfac
    End If
  ElseIf Val(wwwserie) = Val(llave_rep02!far_numser) And CDate(wq_fecha) = CDate(llave_rep02!FAR_fecha_compra) Then
    If Trim(www_fbg) <> Trim(llave_rep02!far_fbg) And wflag_numfac <> "" Then
      GoSub IMPRI_FAC
      wflag_numfac = ""
      wnumfac = llave_rep02!far_numfac
    End If
  End If
'  xl.Application.Visible = True
  wq_fecha = Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy")
  wq_codclie = llave_rep02!far_codclie
  wq_codven = llave_rep02!FAR_codven
  www_fbg = Trim(llave_rep02!far_fbg)
  wq_fbg = "'" & llave_rep02!far_COD_SUNAT ' Trim(llave_rep02!far_fbg)
  wq_docu = "'" & llave_rep02!far_numfac
  
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
     PUB_CAL_INI = llave_rep02!FAR_fecha_compra
     PUB_CAL_FIN = llave_rep02!FAR_fecha_compra
     PUB_CODCIA = LK_CODCIA
     LEER_CAL_LLAVE
     If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 Then
       AWQ_TIPO_CAMBIO = 0
     Else
       AWQ_TIPO_CAMBIO = cal_llave!cal_tipo_cambio
     End If
     If AWQ_TIPO_CAMBIO <= 0 Then
         MsgBox "Definir Tipo de Cambios para el Periodo Actual. Dia : " & llave_rep02!FAR_fecha_compra & " (en el Calendario del Sistema)", 48, Pub_Titulo
         xl.DisplayAlerts = False
         xl.Cells(F1 + 1, 1) = "Falta Tipo de Cambio.... "
         GoTo CANCELA
         Exit Sub
     End If
     If FLAG_TC = 8888 Then AWQ_TIPO_CAMBIO = 1 / AWQ_TIPO_CAMBIO
  Else
    AWQ_TIPO_CAMBIO = 1
  End If
  
  wwwserie = llave_rep02!far_numser
  wq_serie = "'" & llave_rep02!far_numser_c
  wserie = llave_rep02!far_numser_c
  wq_docu = "'" & llave_rep02!far_numfac_c
  www_tipmov = llave_rep02!FAR_TIPMOV
  wq_nombre = ""
  
  'wq_impto = (Format(Val(llave_rep02!far_impto), "0.00") * WS_SIGNO) * AWQ_TIPO_CAMBIO
  'wq_total = (Val(llave_rep02!far_bruto) + Val(llave_rep02!far_impto) - Val(llave_rep02!FAR_TOT_DESCTO) + Val(llave_rep02!FAR_GASTOS)) * WS_SIGNO
  'wq_total = wq_bruto + wq_gastos + wq_impto '(wq_total * AWQ_TIPO_CAMBIO)
  wq_desto = (Val(llave_rep02!FAR_TOT_DESCTO) * WS_SIGNO) * AWQ_TIPO_CAMBIO
  wq_gastos = (Val(llave_rep02!FAR_GASTOS) * WS_SIGNO) * AWQ_TIPO_CAMBIO
  wq_flete = (Val(Nulo_Valor0(llave_rep02!FAR_TOT_FLETE)) * WS_SIGNO) * AWQ_TIPO_CAMBIO
  
  wq_bruto = Format((Val(llave_rep02!far_bruto) - Val(llave_rep02!FAR_TOT_DESCTO)) * WS_SIGNO, "0.00")
  wq_impto = Format(Val(llave_rep02!far_impto) * WS_SIGNO, "0.00")
  wq_total = Format((Val(wq_bruto) + Val(wq_impto)) * AWQ_TIPO_CAMBIO, "0.00")   ' (Val(llave_rep02!far_bruto) + Val(llave_rep02!far_impto) - Val(llave_rep02!FAR_TOT_DESCTO) + Val(llave_rep02!FAR_GASTOS)) * WS_SIGNO * WS_TC
  wq_bruto = Format(wq_total / ((LK_IGV / 100) + 1), "0.00")
  wq_impto = wq_total - wq_bruto + wq_gastos
  
  
  'wq_tot_descto =
  
  
  If wq_bruto = 0 Then wq_total = 0
  wq_estado = llave_rep02!far_estado
  If wq_estado <> "E" Then
    If Left(UCase(llave_rep02!far_subtra), 1) <> "A" Then
     AWQ_COSTO_VENTA = AWQ_COSTO_VENTA + ((llave_rep02!FAR_COSPRO * llave_rep02!FAR_CANTIDAD)) * WS_SIGNO
    End If
  End If
  wq_condi = llave_rep02!far_NUMGUIA
  var_secuencia = Nulo_Valor0(llave_rep02!far_NUM_LOTE)
  wflag_numfac = "A"
 Wflag = "A"
 llave_rep02.MoveNext
Loop
If wflag_numfac = "A" Then
    GoSub IMPRI_FAC
    wflag_numfac = ""
End If
If Wflag = "A" Then
    F1 = F1 + 1
    GoSub TOTAL_DIA
    t_valor_venta = t_valor_venta + s_valor_venta
    t_descto = t_descto + s_descto
    t_valor_igv = t_valor_igv + s_valor_igv
    t_valor_precio = t_valor_precio + s_valor_precio
End If
'  If cheasiento.Value = 1 Then
'    FrmImp2.lblProceso.Caption = "Procesando Pase de Contabilidad . . . "
'    DoEvents
'    GoSub PASE_CONTAB
'  End If
  xcuenta = C1 + 1
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
    FrmImp2.lblProceso.Caption = "Procesando Pase de Contabilidad . . . "
    DoEvents
    GoSub PASE_CONTAB
    cop_llave.Edit
    cop_llave!cop_FLAG_REGC = "A"
    cop_llave.Update
   End If


  F1 = F1 + 2
  xl.Cells(F1, 1) = "Total Genral = "
  xl.Cells(F1, 9) = t_valor_venta
  xl.Cells(F1, 10) = t_descto
  xl.Cells(F1, 11) = t_valor_igv
  xl.Cells(F1, 12) = t_valor_precio


  FrmImp2.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
  xl.Cells(2, 1) = Trim(tra_llave!tra_descripcion)
  xl.Cells(3, 1) = "'" & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
Exit Sub

IMPRI_FAC:
     F1 = F1 + 1
     pu_codclie = wq_codclie
     LEER_CLI_LLAVE
     If Not cli_llave.EOF Then
         wq_codclie = cli_llave!CLI_CODCLIE 'xl.Cells(F1, 2)
         wq_nombre = Trim(cli_llave!cli_nombre)
         wq_codven = Trim(cli_llave!cli_ruc_esposo)
         wcodcta = Trim(cli_llave!CLI_CUENTA_CONTAB)
         wcodcta2 = Trim(cli_llave!CLI_CUENTA_CONTAB2)
     End If
     xl.Cells(F1, 1) = "'" & wq_fecha
     xl.Cells(F1, 2) = wq_codclie
     xl.Cells(F1, 3) = wq_nombre
     xl.Cells(F1, 4) = wq_ruc
     xl.Cells(F1, 5) = wq_condi
     xl.Cells(F1, 6) = wq_fbg
     xl.Cells(F1, 7) = wq_serie
     xl.Cells(F1, 8) = wq_docu
     xl.Cells(F1, 16) = www_tipmov
     xl.Cells(F1, 17) = var_secuencia
     If Trim(www_fbg) = "" Then
       xl.Cells(F1, 18) = "L"
     Else
       xl.Cells(F1, 18) = www_fbg
     End If
     If wq_estado = "E" Then
         xl.Cells(F1, 3) = "[ANULADO] " & wq_nombre
     Else
         xl.Cells(F1, 3) = wq_nombre
     End If
     If wq_estado <> "E" Then
      If Left(cli_llave!CLI_CUENTA_CONTAB, 2) <> "12" Then
       xl.Cells(F1, 9) = wq_bruto
       xl.Cells(F1, 10) = Val(wq_desto)
       xl.Cells(F1, 11) = Val(wq_impto)
       xl.Cells(F1, 12) = Val(wq_total)
       xl.Cells(F1, 15) = Trim(wcodcta)
       xl.Cells(F1, 14) = Trim(wcodcta2)

       'xl.Cells(F1, 14) = Trim(cli_llave!CLI_CUENTA_CONTAB)
       AWQ_BRUTO_ACT_FIJO = AWQ_BRUTO_ACT_FIJO + wq_bruto
       AWQ_NETO_ACT_FIJO = AWQ_NETO_ACT_FIJO + Val(wq_total)
       AWQ_CTA_ACT_FIJO = Trim(cli_llave!CLI_CUENTA_CONTAB)
       s_valor_igv = s_valor_igv + Val(wq_impto)
       AWQ_IMPTO = AWQ_IMPTO + Val(wq_impto)
       ' ACUMULA OTROS ********
       s_valor_venta = s_valor_venta + wq_bruto
       s_descto = s_descto + Val(wq_desto)
       s_valor_precio = s_valor_precio + Val(wq_total)
      Else
       s_valor_venta = s_valor_venta + wq_bruto
       s_descto = s_descto + Val(wq_desto)
       s_valor_igv = s_valor_igv + Val(wq_impto)
       s_valor_precio = s_valor_precio + Val(wq_total)
       xl.Cells(F1, 8) = wq_bruto
       xl.Cells(F1, 9) = Val(wq_desto)
       xl.Cells(F1, 10) = Val(wq_impto)
       xl.Cells(F1, 11) = Val(wq_total)
       xl.Cells(F1, 14) = Trim(cli_llave!CLI_CUENTA_CONTAB)
       AWQ_IMPTO = AWQ_IMPTO + Val(wq_impto)
       AWQ_NETO = AWQ_NETO + Val(wq_total)
       AWQ_BRUTO = AWQ_BRUTO + wq_bruto
       AWQ_DESCTOS = wq_desto
       AWQ_GASTOS = wq_gastos
       AWQ_FLETES = wq_flete
       If wq_condi = "CRED" Then
         xl.Cells(F1, 15) = -1
         AWQ_NETO_CRED = AWQ_NETO_CRED + Val(wq_total)
       Else
         xl.Cells(F1, 15) = 0
         AWQ_NETO_CONT = AWQ_NETO_CONT + Val(wq_total)
       End If
     End If
     End If
Return

CANCELA:
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  lblProceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open PUB_RUTA_REPORTE + "REGCOMPRA.xls", 0, True, 4, WPAS, WPAS
Return

TOTAL_DIA:
  If wfecha = " " Then
    xl.Cells(F1, 1) = "Total Facturas x Mercaderia    = "
  ElseIf wfecha = "K" Then
    xl.Cells(F1, 1) = "Total Facturas Varias    = "
  ElseIf wfecha = "C" Then
    xl.Cells(F1, 1) = "Total N.Creditos  = "
  ElseIf wfecha = "A" Then
    xl.Cells(F1, 1) = "Total N.Debito    = "
  End If
  xl.Cells(F1, 3) = ""
  xl.Cells(F1, 4) = ""
  xl.Cells(F1, 8) = ""
  xl.Cells(F1, 9) = s_valor_venta
  xl.Cells(F1, 10) = s_descto
  xl.Cells(F1, 11) = s_valor_igv
  xl.Cells(F1, 12) = s_valor_precio
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
wsq_fecha = Format(LK_FECHA_COP1, "yyyy/mm/dd")
wsq_fecha2 = Format(LK_FECHA_COP2, "yyyy/mm/dd")

pub_cadena = "DELETE COMOV  WHERE COV_FLAG_AUTOMATICA = '4'  AND COV_CODUSU = '" & LK_CODCIA & "' AND (COV_FECHA_VOUCHER >=  ' " & wsq_fecha & "' AND COV_FECHA_VOUCHER <=  ' " & wsq_fecha2 & "') "
CN.Execute pub_cadena, rdExecDirect

FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Max = 7
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1

If LK_EMP_PTO = "A" Then
PSCOV_VOUCHER(0) = "00"
Else
PSCOV_VOUCHER(0) = LK_CODCIA
End If
PSCOV_VOUCHER(1) = LK_FECHA_COP1
PSCOV_VOUCHER(2) = LK_FECHA_COP2
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
wran1 = "A" & 6 & ":R" & F1
'xl.Application.Visible = True
xl.Application.Worksheets("Hoja1").Range(wran1).Sort Key1:=xl.Application.Worksheets("Hoja1").Range("R6"), Order1:=xlDescending, Key1:=xl.Application.Worksheets("Hoja1").Range("R6")
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
 ws_glosa = "Registro de Compra - " & Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
Else
  ws_glosa = "Registro de Compra"
End If

wdh = "H"
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
wfg1 = "X"
wwwf = ""
flag_pasaigv = ""
For fila = 6 To F1
  If Trim(xl.Cells(fila, 15)) = "" Then
    GoTo otrito
  Else
    If wfg1 = "X" Then
      wfg1 = ""
      wcta = Trim(xl.Cells(fila, 15))
'      xl.Application.Visible = True
      wcta2 = Trim(xl.Cells(fila, 14))
    End If
  End If
  If Trim(xl.Cells(fila, 16)) = "" Then GoTo otrito
  If Trim(xl.Cells(fila, 15)) = "" Then GoTo otrito
  If xl.Cells(fila, 16) = 98 Or xl.Cells(fila, 16) = 97 Then
      If wwwf = "A" Then
          wdh = "H"
          GoSub GRABA
          wdh = "D"
          wcta = wcta2
          wcta_clientes = wcta_bruto
          GoSub GRABA
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
              GoSub GRABA
         End If
         flag_pasaigv = "A"
       End If
        ws_nro_voucher = ws_nro_voucher + 1
        If Val(xl.Cells(fila, 16)) = 97 Then
            ws_glosa = "Nota de Credito/Prov: " + Trim(xl.Cells(fila, 7)) + " - " + xl.Cells(fila, 8)
        ElseIf Val(xl.Cells(fila, 16)) = 98 Then
            ws_glosa = "Nota de Debito/Prov: " + Trim(xl.Cells(fila, 7)) + " - " + xl.Cells(fila, 8)
        End If
        SQ_OPER = 1
        PUB_SECUENCIA = Val(xl.Cells(fila, 17))
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
          wcta_clientes = Abs(Val(xl.Cells(fila, 12)))
          GoSub GRABA
        End If
        If Trim(cnt_llave!CNT_CTA2) <> "" Then
          wcta = cnt_llave!CNT_CTA2  'cuenta de IGV fijo
          wdh = cnt_llave!CNT_DH2
          wcta_clientes = Abs(Val(xl.Cells(fila, 11)))
          GoSub GRABA
        End If
        If Trim(cnt_llave!CNT_CTA3) <> "" Then
          wcta = cnt_llave!CNT_CTA3  'cuenta de bruto fijo
          wdh = cnt_llave!CNT_DH3
          wcta_clientes = Abs(Val(xl.Cells(fila, 9)))
          GoSub GRABA
        End If
        ws_nro_voucher = ws_nro_voucher + 1
        wwwf = ""
        GoTo otrito
  End If
  If wcta <> Trim(xl.Cells(fila, 15)) Then
     wdh = "H"
     GoSub GRABA
     wdh = "D"
     wcta = wcta2
     wcta_clientes = wcta_bruto
     GoSub GRABA
     wcta_clientes = 0
     wcta_bruto = 0
     wcta = Trim(xl.Cells(fila, 15))
     wcta2 = Trim(xl.Cells(fila, 14))
     wcta_clientes = wcta_clientes + Val(xl.Cells(fila, 12))
     wcta_igv = wcta_igv + Val(xl.Cells(fila, 11))
     wcta_bruto = wcta_bruto + Val(xl.Cells(fila, 9))
     If Val(wtipo_mov) = 20 And Left(Trim(wcta2), 2) <> "33" Then
       wcta_total_bruto = wcta_total_bruto + Val(xl.Cells(fila, 9))
     End If
     wwwf = "A"
  Else
     wcta_clientes = wcta_clientes + Val(xl.Cells(fila, 12))
     wcta_igv = wcta_igv + Val(xl.Cells(fila, 11))
     wcta_bruto = wcta_bruto + Val(xl.Cells(fila, 9))
     wtipo_mov = Trim(xl.Cells(fila, 16))
     'If Val(wtipo_mov) <> 98 And Val(wtipo_mov) <> 97 And Val(wtipo_mov) = 20 And Left(Trim(wcta2), 2) <> "33" Then
     wcta_total_bruto = wcta_total_bruto + Val(xl.Cells(fila, 9))
     'Else
     '  ws_nro_voucher = ws_nro_voucher + 1
     'End If
     wwwf = "A"
  End If
otrito:
Next fila

If wwwf = "A" Then
   wdh = "H"
   GoSub GRABA
   wdh = "D"
   wcta = wcta2
   wcta_clientes = wcta_bruto
   GoSub GRABA
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
       GoSub GRABA
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

GRABA:
'    If Trim(wcta) = "40101" Then Stop
     ws_nro_sec = ws_nro_sec + 1
     cov_voucher.AddNew
     If LK_EMP_PTO = "A" Then
       cov_voucher!COV_CODCIA = "00"
     Else
       cov_voucher!COV_CODCIA = LK_CODCIA
     End If
     cov_voucher!COV_FECHA_VOUCHER = LK_FECHA_COP2
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

Public Sub REG_VENTA()
On Error GoTo FINTODO
Dim AWQ_NETO_ACT_FIJO As Currency
Dim AWQ_CTA_ACT_FIJO As String
Dim AWQ_BRUTO_ACT_FIJO As Currency
Dim WCONTROL As Integer
Dim WQ
Dim wfecha
Dim ws_clave
Dim LETRAS(24) As String * 1
Dim wRuta As String
Dim WMONTO As Currency
Dim WCODCLIE As Currency
Dim valor_venta As Currency
Dim descto As Currency
Dim valor_igv As Currency
Dim valor_precio As Currency
Dim s_valor_venta As Currency
Dim s_descto As Currency
Dim s_valor_igv As Currency
Dim s_valor_precio As Currency

Dim t_valor_venta As Currency
Dim t_descto As Currency
Dim t_valor_igv As Currency
Dim t_valor_precio As Currency

Dim wnumfac As Currency
Dim Wflag As String * 1
Dim wsFECHA1, wsFECHA2
Dim xcuenta As Integer
Dim wq_fecha, wq_codclie, wq_codven, wq_docu, wq_nombre, wq_bruto, wq_gastos, wq_desto, wq_flete, wq_fbg, wq_serie
Dim wq_tot_descto, wq_impto, wq_total, wq_estado, wq_condi
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
Dim WS_TC As Currency
pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtCampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtCampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtCampo2.Text)
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
If CDate(wsFECHA1) <> LK_FECHA_COP1 Then cheasiento.Value = 0
If CDate(wsFECHA2) <> LK_FECHA_COP2 Then cheasiento.Value = 0
GoSub WEXCEL
pub_cadena = ""
xcuenta = 0

pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp2.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
WCONTROL = 0
F1 = 5  'Fila Inicial
t_valor_venta = 0
t_descto = 0
t_valor_igv = 0
t_valor_precio = 0

'NCREDITO: ' empieza

'WCONTROL = WCONTROL + 1
WMONTO = 0
WCODCLIE = 0
valor_venta = 0
descto = 0
valor_igv = 0
valor_precio = 0
wnumfac = 0
Wflag = ""
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

NCREDITO:
s_valor_venta = 0
s_descto = 0
s_valor_igv = 0
s_valor_precio = 0

WCONTROL = WCONTROL + 1
If WCONTROL = 1 Then
  pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = 10 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND (FAR_FBG = 'F' OR FAR_FBG = 'B') ORDER BY FAR_TIPMOV, FAR_FBG DESC ,FAR_NUMSER,FAR_NUMFAC"
  WS_SIGNO = 1
ElseIf WCONTROL = 2 Then
  pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = 97 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND FAR_FBG <> ' ' AND FAR_FBG <> 'G' and FAR_FBG <> 'X' AND FAR_CP = 'C'  ORDER BY FAR_TIPMOV, FAR_FBG DESC , FAR_FECHA_COMPRA,  FAR_NUMSER,FAR_NUMFAC"
  WS_SIGNO = -1
ElseIf WCONTROL = 3 Then
  pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = 98 AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ? AND FAR_FBG <> ' ' AND FAR_FBG <> 'G' and FAR_FBG <> 'X' AND FAR_CP = 'C' ORDER BY FAR_TIPMOV, FAR_FBG DESC , FAR_FECHA_COMPRA,  FAR_NUMSER,FAR_NUMFAC"
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
FrmImp2.lblProceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep02.Requery
If llave_rep02.EOF Then
  GoTo OTRO_DOCUMENTO
End If
FrmImp2.lblProceso.Caption = "Procesando . . . "
DoEvents
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Max = llave_rep02.RowCount

Wflag = ""
SQ_OPER = 1
pu_cp = "C"
pu_codcia = LK_CODCIA
wnumfac = llave_rep02!far_numfac
wfecha = llave_rep02!far_fbg 'llave_rep02!far_fecha
wserie = llave_rep02!far_numser
wq_fecha = "01/01/1900"
xcuenta = 0
Wflag = "A"
wflag_numfac = "A"
AWQ_DESCTOS = 0
AWQ_GASTOS = 0
AWQ_FLETES = 0

Do Until llave_rep02.EOF
'  If llave_rep02!FAR_numfac = 401 Then Stop
  FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
  If wfecha <> llave_rep02!far_fbg Then '
     If wflag_numfac = "A" Then
       GoSub IMPRI_FAC
     End If
     wflag_numfac = ""
     F1 = F1 + 1
     GoSub TOTAL_DIA
     t_valor_venta = t_valor_venta + s_valor_venta
     t_descto = t_descto + s_descto
     t_valor_igv = t_valor_igv + s_valor_igv
     t_valor_precio = t_valor_precio + s_valor_precio
     wnumfac = llave_rep02!far_numfac
     wfecha = llave_rep02!far_fbg
     s_valor_venta = 0
     s_descto = 0
     s_valor_igv = 0
     s_valor_precio = 0
  End If
  
  If wnumfac <> llave_rep02!far_numfac Then
     GoSub IMPRI_FAC
     wflag_numfac = ""
     wnumfac = llave_rep02!far_numfac
  ElseIf Val(wserie) <> Val(llave_rep02!far_numser) And CDate(wq_fecha) = CDate(llave_rep02!FAR_fecha_compra) Then
    'If Trim(wq_fbg) <> Trim(llave_rep02!far_fbg) And wflag_numfac <> "" Then
    If wflag_numfac <> "" Then
      GoSub IMPRI_FAC
      wflag_numfac = ""
      wnumfac = llave_rep02!far_numfac
    End If
  ElseIf Val(wserie) = Val(llave_rep02!far_numser) And CDate(wq_fecha) = CDate(llave_rep02!FAR_fecha_compra) Then
    If Trim(wq_fbg) <> Trim(llave_rep02!far_fbg) And wflag_numfac <> "" Then
      GoSub IMPRI_FAC
      wflag_numfac = ""
      wnumfac = llave_rep02!far_numfac
    End If
  End If
'  xl.Application.Visible = True
  wq_fecha = Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy")
  wq_codclie = llave_rep02!far_codclie
  wq_codven = llave_rep02!FAR_codven
  wq_fbg = Trim(llave_rep02!far_fbg)
  wq_serie = "'" & llave_rep02!far_numser
  wserie = llave_rep02!far_numser
  wq_docu = "'" & llave_rep02!far_numfac
  wq_nombre = ""
  WS_TC = 1
  If llave_rep02!FAR_MONEDA = "D" Then
    WS_TC = JALAR(llave_rep02!FAR_fecha_compra)
    If WS_TC <= 0 Then
        MsgBox "Falta Ingresar el Tipo de Cambio del día : " & Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy"), 48, Pub_Titulo
        GoTo CANCELA
    End If
  End If
  
  wq_bruto = Format((Val(llave_rep02!far_bruto) - Val(llave_rep02!FAR_TOT_DESCTO)) * WS_SIGNO, "0.00")
  wq_impto = Format(Val(llave_rep02!far_impto) * WS_SIGNO, "0.00")
  wq_total = Format((Val(wq_bruto) + Val(wq_impto)) * WS_TC, "0.00")  ' (Val(llave_rep02!far_bruto) + Val(llave_rep02!far_impto) - Val(llave_rep02!FAR_TOT_DESCTO) + Val(llave_rep02!FAR_GASTOS)) * WS_SIGNO * WS_TC
  wq_bruto = Format(wq_total / ((LK_IGV / 100) + 1), "0.00")
  wq_impto = wq_total - wq_bruto
  wq_desto = Val(llave_rep02!FAR_TOT_DESCTO) * WS_SIGNO * WS_TC
  wq_gastos = Val(llave_rep02!FAR_GASTOS) * WS_SIGNO * WS_TC
  wq_flete = Val(Nulo_Valor0(llave_rep02!FAR_TOT_FLETE)) * WS_SIGNO * WS_TC
  'wq_tot_descto =
  
  If wq_bruto = 0 Then wq_total = 0
  wq_estado = llave_rep02!far_estado
  If wq_estado <> "E" Then
    If Left(UCase(llave_rep02!far_subtra), 1) <> "A" Then
     AWQ_COSTO_VENTA = AWQ_COSTO_VENTA + ((llave_rep02!FAR_COSPRO * llave_rep02!FAR_CANTIDAD)) * WS_SIGNO
    End If
  End If
  If llave_rep02!far_signo_car <> 0 And llave_rep02!FAR_DIAS <> 0 Then
      If wq_estado <> "E" Then
         'AWQ_NETO_CONT = wq_total
       End If
     wq_condi = "CRED"
  Else
      If wq_estado <> "E" Then
         'AWQ_NETO_CRED = AWQ_NETO_CRED + wq_total
      End If
     wq_condi = "CONT"
  End If
  wflag_numfac = "A"
 Wflag = "A"
 llave_rep02.MoveNext
Loop
If wflag_numfac = "A" Then
    GoSub IMPRI_FAC
    wflag_numfac = ""
End If
If Wflag = "A" Then
    F1 = F1 + 1
    GoSub TOTAL_DIA
    t_valor_venta = t_valor_venta + s_valor_venta
    t_descto = t_descto + s_descto
    t_valor_igv = t_valor_igv + s_valor_igv
    t_valor_precio = t_valor_precio + s_valor_precio
End If
'  If cheasiento.Value = 1 Then
'    FrmImp2.lblProceso.Caption = "Procesando Pase de Contabilidad . . . "
'    DoEvents
'    GoSub PASE_CONTAB
'  End If
  xcuenta = C1 + 1
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
    FrmImp2.lblProceso.Caption = "Procesando Pase de Contabilidad . . . "
    DoEvents
    GoSub PASE_CONTAB
    cop_llave.Edit
    cop_llave!cop_FLAG_REGV = "A"
    cop_llave.Update
   End If


  F1 = F1 + 2
  xl.Cells(F1, 1) = "Total Genral = "
  xl.Worksheets(1).Rows(F1).RowHeight = 20
  xl.Cells(F1, 8) = t_valor_venta
  xl.Cells(F1, 9) = t_descto
  xl.Cells(F1, 10) = t_valor_igv
  xl.Cells(F1, 11) = t_valor_precio


  FrmImp2.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
  xl.Cells(2, 1) = Trim(tra_llave!tra_descripcion)
  xl.Cells(3, 1) = "'" & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp2.pantalla.Enabled = True
  FrmImp2.cerrar.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
Exit Sub

IMPRI_FAC:
       F1 = F1 + 1
       pu_codclie = wq_codclie
       LEER_CLI_LLAVE
       If Not cli_llave.EOF Then
         wq_codclie = cli_llave!CLI_CODCLIE 'xl.Cells(F1, 2)
         wq_nombre = Trim(cli_llave!cli_nombre)
         wq_ruc = Trim(cli_llave!cli_ruc_esposo)
       End If
     xl.Cells(F1, 1) = "'" & wq_fecha
 '    xl.Cells(F1, 2) = wq_codclie
     xl.Cells(F1, 3) = wq_ruc
     xl.Cells(F1, 2) = wq_nombre
     xl.Cells(F1, 4) = wq_condi
     xl.Cells(F1, 5) = wq_fbg
     xl.Cells(F1, 6) = wq_serie
     xl.Cells(F1, 7) = wq_docu
     If wq_estado = "E" Then
         xl.Cells(F1, 2) = "[ANULADO] " & wq_nombre
     Else
         xl.Cells(F1, 2) = wq_nombre
     End If
     If wq_estado <> "E" Then
      If Left(cli_llave!CLI_CUENTA_CONTAB, 2) <> "12" Then
       xl.Cells(F1, 8) = wq_bruto
       xl.Cells(F1, 9) = Val(wq_desto)
       xl.Cells(F1, 10) = Val(wq_impto)
       xl.Cells(F1, 11) = Val(wq_total)
       'xl.Cells(F1, 14) = Trim(cli_llave!CLI_CUENTA_CONTAB)
       AWQ_BRUTO_ACT_FIJO = AWQ_BRUTO_ACT_FIJO + wq_bruto
       AWQ_NETO_ACT_FIJO = AWQ_NETO_ACT_FIJO + Val(wq_total)
       AWQ_CTA_ACT_FIJO = Trim(cli_llave!CLI_CUENTA_CONTAB)
       s_valor_igv = s_valor_igv + Val(wq_impto)
       AWQ_IMPTO = AWQ_IMPTO + Val(wq_impto)
       ' ACUMULA OTROS ********
       s_valor_venta = s_valor_venta + wq_bruto
       s_descto = s_descto + Val(wq_desto)
       s_valor_precio = s_valor_precio + Val(wq_total)
      Else
       s_valor_venta = s_valor_venta + wq_bruto
       s_descto = s_descto + Val(wq_desto)
       s_valor_igv = s_valor_igv + Val(wq_impto)
       s_valor_precio = s_valor_precio + Val(wq_total)
       xl.Cells(F1, 8) = wq_bruto
       xl.Cells(F1, 9) = Val(wq_desto)
       xl.Cells(F1, 10) = Val(wq_impto)
       xl.Cells(F1, 11) = Val(wq_total)
       xl.Cells(F1, 14) = Trim(cli_llave!CLI_CUENTA_CONTAB)
       AWQ_IMPTO = AWQ_IMPTO + Val(wq_impto)
       AWQ_NETO = AWQ_NETO + Val(wq_total)
       AWQ_BRUTO = AWQ_BRUTO + wq_bruto
       AWQ_DESCTOS = wq_desto
       AWQ_GASTOS = wq_gastos
       AWQ_FLETES = wq_flete
       If wq_condi = "CRED" Then
         xl.Cells(F1, 15) = -1
         AWQ_NETO_CRED = AWQ_NETO_CRED + Val(wq_total)
       Else
         xl.Cells(F1, 15) = 0
         AWQ_NETO_CONT = AWQ_NETO_CONT + Val(wq_total)
       End If
     End If
     End If
Return

CANCELA:
  FrmImp2.pantalla.Enabled = True
  FrmImp2.pantalla.Caption = "Por &Pantalla"
  FrmImp2.lblProceso.Visible = False
  FrmImp2.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  lblProceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open PUB_RUTA_REPORTE + "REGVENTA.xls", 0, True, 4, WPAS, WPAS
Return

TOTAL_DIA:
  If wfecha = "F" Then
    xl.Cells(F1, 1) = "Total Facturas    = "
  ElseIf wfecha = "B" Then
    xl.Cells(F1, 1) = "Total Boletas     = "
  ElseIf wfecha = "N" Then
    xl.Cells(F1, 1) = "Total N.Creditos  = "
  ElseIf wfecha = "D" Then
    xl.Cells(F1, 1) = "Total N.Debito    = "
  End If
  xl.Worksheets(1).Rows(F1).RowHeight = 20
  xl.Cells(F1, 2) = ""
  xl.Cells(F1, 3) = ""
  xl.Cells(F1, 7) = ""
  xl.Cells(F1, 8) = s_valor_venta
  xl.Cells(F1, 9) = s_descto
  xl.Cells(F1, 10) = s_valor_igv
  xl.Cells(F1, 11) = s_valor_precio
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
 ws_glosa = "Registro de Venta - " & Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
End If
wsq_fecha = Format(LK_FECHA_COP1, "yyyy/mm/dd")
wsq_fecha2 = Format(LK_FECHA_COP2, "yyyy/mm/dd")
pub_cadena = "DELETE COMOV  WHERE  COV_FLAG_AUTOMATICA = '3' AND COV_CODUSU = '" & LK_CODCIA & "' AND (COV_FECHA_VOUCHER >=  ' " & wsq_fecha & "' AND COV_FECHA_VOUCHER <=  ' " & wsq_fecha2 & "')"
CN.Execute pub_cadena, rdExecDirect

FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.Max = 10
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
PSCOV_VOUCHER(0) = wscodcia
PSCOV_VOUCHER(1) = LK_FECHA_COP1
PSCOV_VOUCHER(2) = LK_FECHA_COP2
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
wran1 = "A" & 6 & ":Q" & F1
xl.Application.Worksheets("Hoja1").Range(wran1).Sort Key1:=xl.Application.Worksheets("Hoja1").Range("N6")
fila = 6
' xl.Application.Visible = True
wcta = Trim(xl.Cells(fila, 14))
wcta_clientes = 0
wdh = "D"
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
For fila = 6 To F1
  If Val(xl.Cells(fila, 15)) = 0 Then GoTo otrito
  'If Trim(xl.Cells(fila, 5)) = "N" Then
  '   GoTo OTRITO
  'End If
  If Val(xl.Cells(fila, 14)) = 0 Then GoTo otrito
   If wcta <> Trim(xl.Cells(fila, 14)) Then
     GoSub GRABA
     wcta_clientes = 0
     wcta = Trim(xl.Cells(fila, 14))
     wcta_clientes = wcta_clientes + Val(xl.Cells(fila, 11))
  Else
     wcta_clientes = wcta_clientes + Val(xl.Cells(fila, 11))
  End If
otrito:
Next fila
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
GoSub GRABA
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
  GoSub GRABA
 End If
  wcta = AWQ_CTA_ACT_FIJO
  wdh = "D"
  wcta_clientes = AWQ_NETO_ACT_FIJO
  GoSub GRABA
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
 GoSub GRABA
End If
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
If Trim(cnt_llave!CNT_CTA2) <> "" Then
 wcta = cnt_llave!CNT_CTA2 'Impuesto
 wdh = cnt_llave!CNT_DH2
 wcta_clientes = AWQ_IMPTO
 GoSub GRABA
End If

If AWQ_NETO_CONT <> 0 Then
 SQ_OPER = 1
 PUB_SECUENCIA = 26
 PUB_CODTRA = 2401
 PUB_CODCIA = wscodcia
 LEER_CNT_LLAVE
 If cnt_llave.EOF Then
  MsgBox "Error de Dato de Transaccion , Consulte a su Proveedor.", 48, Pub_Titulo
  End
 End If
 If Trim(cnt_llave!CNT_CTA1) <> "" Then
  wcta = cnt_llave!CNT_CTA1  'Ventas al contado
  wdh = "D" ' ES HABER PERO AL DEBE
  wcta_clientes = AWQ_NETO_CONT
  GoSub GRABA
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
 GoSub GRABA
End If
FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1
If Trim(cnt_llave!CNT_CTA4) <> "" Then
 wcta = cnt_llave!CNT_CTA4  'Ventas a costo
 wdh = cnt_llave!CNT_DH4
 wcta_clientes = AWQ_COSTO_VENTA
 GoSub GRABA
End If

FrmImp2.ProgBar.Value = FrmImp2.ProgBar.Value + 1

Return

GRABA:
     If wcta_clientes = 0 Then Return
     ws_nro_sec = ws_nro_sec + 1
     cov_voucher.AddNew
     cov_voucher!COV_CODCIA = wscodcia
     cov_voucher!COV_FECHA_VOUCHER = LK_FECHA_COP2
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


Public Function JALAR(wfecha As Date) As Currency
PUB_CAL_INI = wfecha
PUB_CAL_FIN = wfecha
pu_codcia = LK_CODCIA
PUB_CODCIA = LK_CODCIA
SQ_OPER = 1
LEER_CAL_LLAVE
If cal_llave.EOF Then
  JALAR = -1
  Exit Function
End If
If IsNull(cal_llave!cal_tipo_cambio) Then
  JALAR = -1
  Exit Function
End If
JALAR = cal_llave!cal_tipo_cambio

End Function

