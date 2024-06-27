VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F6E4F630-E903-11D5-8BB9-0080AD40A177}#1.18#0"; "OSControlsUser.ocx"
Begin VB.Form frmCotizacion 
   Caption         =   "Cotización a Clientes"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   ControlBox      =   0   'False
   Icon            =   "frmCotizacion.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Reportes 
      Left            =   795
      Top             =   7065
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Im&primir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5955
      TabIndex        =   48
      Top             =   6645
      Width           =   1140
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
      Height          =   330
      Left            =   9375
      TabIndex        =   47
      Top             =   6270
      Width           =   1140
   End
   Begin VB.CommandButton cmdConsulta 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7092
      TabIndex        =   46
      Top             =   6270
      Width           =   1140
   End
   Begin VB.CheckBox chkCambiar 
      Caption         =   "Cambiar"
      Height          =   240
      Left            =   10140
      TabIndex        =   45
      Top             =   6690
      Width           =   1215
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   3945
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   4605
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   3
      Left            =   3990
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   4245
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2400
      Left            =   2205
      TabIndex        =   43
      Top             =   3345
      Visible         =   0   'False
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   4233
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   40
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtFindRegistro 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBE1D3&
      Height          =   285
      Left            =   3030
      TabIndex        =   39
      Top             =   2565
      Visible         =   0   'False
      Width           =   1125
   End
   Begin MSMask.MaskEdBox txt2 
      Height          =   270
      Left            =   3045
      TabIndex        =   41
      Top             =   2955
      Visible         =   0   'False
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777152
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dddddd"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtNumSer 
      Height          =   285
      Left            =   8370
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   6660
      Width           =   435
   End
   Begin VB.TextBox txtNumFac 
      Height          =   285
      Left            =   8865
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   6660
      Width           =   1020
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "&Anular"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8229
      TabIndex        =   32
      Top             =   6270
      Width           =   1140
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerra&r"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10515
      TabIndex        =   31
      Top             =   6270
      Width           =   915
   End
   Begin VB.CommandButton cmdIngresar 
      Caption         =   "&Ingresar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5955
      TabIndex        =   30
      Top             =   6270
      Width           =   1140
   End
   Begin VB.Frame frmSubtotal 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   45
      TabIndex        =   21
      Tag             =   "cls"
      Top             =   6240
      Width           =   5865
      Begin ComctlLib.StatusBar stb_resumen 
         Height          =   300
         Left            =   45
         TabIndex        =   22
         Top             =   90
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   529
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   5
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   2293
               MinWidth        =   2293
               Text            =   "SubTotal"
               TextSave        =   "SubTotal"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   2293
               MinWidth        =   2293
               Text            =   "Descto.(%)"
               TextSave        =   "Descto.(%)"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   2293
               MinWidth        =   2293
               Text            =   "Impto."
               TextSave        =   "Impto."
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Object.Width           =   776
               MinWidth        =   776
               TextSave        =   ""
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   2293
               MinWidth        =   2293
               Text            =   "Total"
               TextSave        =   "Total"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ComctlLib.StatusBar stbSubtotales 
         Height          =   300
         Left            =   45
         TabIndex        =   23
         Tag             =   "cls"
         Top             =   405
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   529
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   5
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               Object.Width           =   2293
               MinWidth        =   2293
               TextSave        =   ""
               Key             =   ""
               Object.Tag             =   ""
               Object.ToolTipText     =   "SubTotal"
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               Object.Width           =   2293
               MinWidth        =   2293
               TextSave        =   ""
               Key             =   ""
               Object.Tag             =   ""
               Object.ToolTipText     =   "Descto.(%)"
            EndProperty
            BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               Object.Width           =   2293
               MinWidth        =   2293
               TextSave        =   ""
               Key             =   ""
               Object.Tag             =   ""
               Object.ToolTipText     =   "Impto."
            EndProperty
            BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Object.Width           =   776
               MinWidth        =   776
               TextSave        =   ""
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               Object.Width           =   2293
               MinWidth        =   2293
               TextSave        =   ""
               Key             =   ""
               Object.Tag             =   ""
               Object.ToolTipText     =   "Total"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin OSControlsUser.OSFindItem txtcli 
      Height          =   285
      Left            =   1215
      TabIndex        =   20
      Tag             =   "cls"
      Top             =   120
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   503
      Enabled         =   0   'False
      Locked          =   0   'False
   End
   Begin VB.Frame f1 
      Enabled         =   0   'False
      Height          =   1545
      Left            =   135
      TabIndex        =   0
      Top             =   -60
      Width           =   11775
      Begin VB.TextBox txtContacto 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   36
         Tag             =   "cls"
         Top             =   675
         Width           =   5595
      End
      Begin VB.TextBox txtOferta 
         Height          =   705
         Left            =   6840
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Tag             =   "cls"
         Top             =   735
         Width           =   4515
      End
      Begin VB.TextBox txtNroCot 
         Height          =   285
         Left            =   5415
         MaxLength       =   20
         TabIndex        =   25
         Tag             =   "cls"
         Top             =   1050
         Width           =   1245
      End
      Begin OSControlsUser.ctlMaskEdBox txtFecha 
         Height          =   300
         Left            =   3120
         TabIndex        =   24
         Tag             =   "cls"
         Top             =   1050
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         Format          =   "ddddd"
      End
      Begin VB.TextBox txtruc 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   8235
         Locked          =   -1  'True
         TabIndex        =   7
         Tag             =   "cls"
         Top             =   195
         Width           =   1515
      End
      Begin VB.ComboBox moneda 
         Height          =   315
         ItemData        =   "frmCotizacion.frx":000C
         Left            =   1065
         List            =   "frmCotizacion.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "cls"
         Top             =   1050
         Width           =   975
      End
      Begin VB.ComboBox i_fbg 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCotizacion.frx":002F
         Left            =   5805
         List            =   "frmCotizacion.frx":0039
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2385
         Width           =   1110
      End
      Begin VB.ComboBox i_destino 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCotizacion.frx":0043
         Left            =   240
         List            =   "frmCotizacion.frx":004D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2370
         Width           =   5490
      End
      Begin VB.ComboBox i_condi 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCotizacion.frx":0066
         Left            =   7905
         List            =   "frmCotizacion.frx":0068
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2430
         Width           =   2730
      End
      Begin VB.TextBox i_dias 
         Enabled         =   0   'False
         Height          =   300
         Left            =   7080
         TabIndex        =   2
         Top             =   2400
         Width           =   495
      End
      Begin VB.ComboBox cmdtipo 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCotizacion.frx":006A
         Left            =   8445
         List            =   "frmCotizacion.frx":0074
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2040
         Width           =   2745
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         Caption         =   "Contacto :"
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
         Index           =   7
         Left            =   210
         TabIndex        =   37
         Tag             =   "9999"
         Top             =   705
         Width           =   765
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         Caption         =   "Oferta : <para pasar a otra fila presione Ctrl +Enter>"
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
         Index           =   5
         Left            =   6885
         TabIndex        =   28
         Tag             =   "9999"
         Top             =   510
         Width           =   3900
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         Caption         =   "Nº Solicitud : "
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
         Index           =   3
         Left            =   4440
         TabIndex        =   27
         Tag             =   "9999"
         Top             =   1095
         Width           =   960
      End
      Begin VB.Label lcodart 
         Caption         =   "Fecha : "
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
         Height          =   255
         Index           =   0
         Left            =   2445
         TabIndex        =   26
         Tag             =   "9999"
         Top             =   1095
         Width           =   735
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
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
         Index           =   4
         Left            =   375
         TabIndex        =   18
         Tag             =   "9999"
         Top             =   195
         Width           =   600
      End
      Begin VB.Label lcodart 
         Caption         =   "R.U.C. ó D.N.I. :"
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
         Index           =   1
         Left            =   7050
         TabIndex        =   17
         Tag             =   "9999"
         Top             =   195
         Width           =   1140
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         Caption         =   "Moneda : "
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
         Index           =   2
         Left            =   255
         TabIndex        =   16
         Tag             =   "9999"
         Top             =   1095
         Width           =   720
      End
      Begin VB.Label lblcli 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2250
         TabIndex        =   15
         Tag             =   "cls"
         Top             =   195
         Width           =   4410
      End
      Begin VB.Label Label1 
         Caption         =   "Vendedor:"
         Height          =   255
         Index           =   0
         Left            =   5175
         TabIndex        =   14
         Top             =   2025
         Width           =   915
      End
      Begin VB.Label lcodart 
         Caption         =   "Fact./Bolet."
         Height          =   255
         Index           =   8
         Left            =   5805
         TabIndex        =   13
         Tag             =   "9999"
         Top             =   2430
         Width           =   1005
      End
      Begin VB.Label lblven 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5175
         TabIndex        =   12
         Top             =   2010
         Width           =   2955
      End
      Begin VB.Label lcodart 
         Caption         =   "   Destino Almacen :"
         Height          =   255
         Index           =   9
         Left            =   150
         TabIndex        =   11
         Tag             =   "9999"
         Top             =   2460
         Width           =   1470
      End
      Begin VB.Label lcodart 
         Caption         =   "Condición Venta"
         Height          =   255
         Index           =   10
         Left            =   7935
         TabIndex        =   10
         Tag             =   "9999"
         Top             =   2430
         Width           =   1995
      End
      Begin VB.Label lcodart 
         Caption         =   "Dias Cred."
         Height          =   255
         Index           =   11
         Left            =   7110
         TabIndex        =   9
         Tag             =   "9999"
         Top             =   2400
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Venta"
         Height          =   255
         Index           =   1
         Left            =   8475
         TabIndex        =   8
         Top             =   2055
         Width           =   2010
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4350
      Left            =   120
      TabIndex        =   38
      Top             =   1680
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   7673
      _Version        =   393216
      Enabled         =   0   'False
      AllowUserResizing=   1
   End
   Begin OSControlsUser.OSFindItem Txt_key 
      Height          =   285
      Left            =   4665
      TabIndex        =   19
      Top             =   1155
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   503
      Enabled         =   0   'False
      Locked          =   0   'False
   End
   Begin VB.Label lblItems 
      AutoSize        =   -1  'True
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   255
      TabIndex        =   49
      Top             =   1485
      Width           =   45
   End
   Begin VB.Label lcodart 
      AutoSize        =   -1  'True
      Caption         =   "Nº : "
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
      Index           =   6
      Left            =   7935
      TabIndex        =   35
      Tag             =   "9999"
      Top             =   6705
      Width           =   330
   End
End
Attribute VB_Name = "frmCotizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PSFAR_TRANS As rdoQuery
Dim FAR_TRANS As rdoResultset
Dim PSPedido_LLAVE As rdoQuery
Dim Pedido_llave As rdoResultset
Dim cbmTop As Long
Dim cbmLeft As Long
Dim iRowGrd As Long

Private loListItem As MSComctlLib.ListItem
Dim liIndexRegAct As Integer
Dim lvKey As Variant
Dim FlagTypeFind As Integer
Dim FlagLvw As Integer
Dim FLAG As Integer
Dim FlagNumC As Integer
Private ColumnNext() As Integer

Dim ColorCellDefault As Variant

Private Sub cbo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim sText As String
Dim sKey As String
Dim iPos As Integer
On Error GoTo Handler
    
    If KeyCode = 13 Then
        If cbo(Index).ListIndex = -1 Then Exit Sub
        If grd.COL = 4 Then
            sText = Trim(cbo(Index).Text)
            GoTo SEGUIR
        End If
        If Len(Trim(cbo(Index).Text)) = 0 Then GoTo SEGUIR
        iPos = InStr(1, cbo(Index).Text, "      ")
        If iPos = 0 Then iPos = Len(cbo(Index).Text)
        sText = Trim(Mid(cbo(Index).Text, 1, iPos))
SEGUIR:
        
        grd.Text = sText
        If grd.COL = 4 Then
            CalculaSubTotales1 Val(grd.Row), Val(grd.COL)
            CalculaTotales1 Val(grd.Row), Val(grd.COL)
        End If
        cbo(Index).Visible = False
        grd.SetFocus
        ValRowGrd grd.Row
        If Index = 3 Then
            sKey = Trim(Right(cbo(Index).Text, 20))
            grd.TextMatrix(grd.Row, 15) = sKey
            LOADPRECIOS
        End If
        If ColumnNext(grd.COL) = -1 And Trim(grd.TextMatrix(grd.Row, 13)) = "X" And (grd.Row = grd.Rows - 1) Then GoTo OtraFila
        If grd.Cols - 1 > grd.COL Then
            grd.COL = grd.COL + ColumnNext(grd.COL)
        ElseIf grd.Cols = grd.COL + 1 Then
            If grd.Row = grd.Rows - 1 Then
OtraFila:
                grd.Rows = grd.Rows + 1
            End If
            grd.COL = 1
            grd.Row = grd.Row + 1
        End If
    End If
    lblItems.Caption = "Nº de Items " & grd.Rows - 2
    If grd.Rows >= 27 Then
        'MsgBox "Se llego al tope de items.. maximo 25 items", vbInformation, Pub_Titulo
        grd.Rows = grd.Rows - 1
        lblItems.Caption = "Nº de Items " & grd.Rows - 1
        Exit Sub
    End If
    Exit Sub
Handler:
End Sub

'Private Sub cmbprecios_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        grd.TextMatrix(iRowGrd, 4) = cmbprecios.Text
'        Call CalculaSubTotales1(iRowGrd, 4)
'        Call CalculaTotales1(iRowGrd, 4)
'        cmbprecios.Visible = False
'    End If
'End Sub

Private Sub cbo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cbo(Index).Visible = False
        grd.SetFocus
    End If
End Sub

Private Sub cbo_LostFocus(Index As Integer)
    cbo_KeyUp Index, 27, 0
End Sub

Private Sub chkCambiar_Click()
    If chkCambiar.Value = 1 Then
        txtNumFac.Locked = False
        txtNumFac.SetFocus
        Azul txtNumFac, txtNumFac
    Else
        txtNumFac.Locked = True
    End If
End Sub

Private Sub cmdAnular_Click()
Dim SQL As String
Dim vbresp As Integer

    vbresp = MsgBox("Esta seguro que desea anular el Pedido????..", vbYesNo, Pub_Titulo)
    If vbresp = vbYes Then
        SQL = "UPDATE PEDIDOS SET PED_ESTADO = 'E' WHERE PED_TIPMOV = 300 AND PED_CODCIA='" & LK_CODCIA & "' AND PED_NUMSER='" & txtNumSer.Text & "' AND PED_NUMFAC = " & txtNumFac.Text
        CN.Execute SQL
        MsgBox "Se realizó la Transacción", vbOKOnly, Pub_Titulo
        cmdcancelar_Click
    End If
End Sub

Private Sub cmdcancelar_Click()
    ClearForm Me
    FormatGrid
    txtcli_Cancel
    txtNumFac.Locked = True
    chkCambiar.Value = 0
    cmdIngresar.Enabled = True
    cmdConsulta.Enabled = False
    cmdAnular.Enabled = False
    If cmdIngresar.Caption = "&Grabar" Then
       cmdIngresar.Caption = "&Ingresar"
       f1.Enabled = False
       grd.Enabled = False
       cmdIngresar.SetFocus
    End If
    If cmdConsulta.Caption = "&Grabar" Then
       cmdConsulta.Caption = "&Modificar"
    End If
    txt.Visible = False
    txtFindRegistro.Visible = False
    cbo(3).Visible = False
    cbo(4).Visible = False
    stbSubtotales.Panels(1).Text = "0.000"
    stbSubtotales.Panels(3).Text = "0.000"
    stbSubtotales.Panels(4).Text = ""
    stbSubtotales.Panels(5).Text = "0.000"
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdconsulta_Click()
Dim iRowsFacar As Integer
Dim iCount As Integer
Dim iRow1 As Integer
On Error GoTo Handler
    SQ_OPER = 1
    PUB_TIPMOV = 300
    pu_codcia = LK_CODCIA
    PUB_PEDSER = txtNumSer.Text
    PUB_PEDFAC = txtNumFac.Text
    LEER_PED_LLAVE
    If Not ped_llave.EOF Then
        pub_cadena = "SELECT * FROM CONTROLL"
        CN.Execute "Begin Transaction", rdExecDirect
        Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
        
        iRowsFacar = grd.Rows - 1
        Do While Not ped_llave.EOF
            iCount = iCount + 1
            If Trim(grd.TextMatrix(iCount, 1)) = "" Then
                GoTo NextReg1
            End If
            If iCount > grd.Rows - 1 Then
                ped_llave.Delete
                iCount = iCount - 1
                GoTo NextReg1
            End If
            If ped_llave("ped_numsec") <> Val(grd.TextMatrix(iCount, 14)) Then
                ped_llave.Delete
                iCount = iCount - 1
                GoTo NextReg1
            End If
            ped_llave.Edit
            AsignaValores iCount, 0
            ped_llave.Update
NextReg1:
            ped_llave.MoveNext
        Loop
        
        For iRow1 = iCount + 1 To iRowsFacar
            If Trim(grd.TextMatrix(iRow1, 1)) <> "" Then
                iCount = iCount + 1
                ped_llave.AddNew
                AsignaValores iCount, 1
                ped_llave.Update
            End If
        Next iRow1
        CN.Execute "Commit Transaction", rdExecDirect
        con_llave.Close
        MsgBox "Los datos se grabaron correctamente", vbInformation, Pub_Titulo
        cmdConsulta.Caption = "&Modificar"
        cmdcancelar_Click
    Else
        MsgBox "Error Cotizaciones no existe doccumento", vbCritical, Pub_Titulo
    End If
    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical, Pub_Titulo
     con_llave.Close
     CN.Execute "Rollback Transaction", rdExecDirect
End Sub

Private Sub CmdImprimir_Click()
Dim RUCEmpresa As String
On Error GoTo Handler
    Reportes.Connect = PUB_ODBC
    If LK_CODCIA = "01" Then 'pereda
        Reportes.ReportFileName = PUB_RUTA_OTRO & "COTIZACLIpere.RPT"
    ElseIf LK_CODCIA = "02" Then ' chif
        Reportes.ReportFileName = PUB_RUTA_OTRO & "COTIZACLI.RPT"
    End If
    
    Reportes.Formulas(0) = "CIA='RUC " & Trim(par_llave("par_nombre_corto")) & "'" ' " & Trim(par_llave("par_nombre")) & "
    Reportes.WindowTitle = "Cotización de Clientes"
    Reportes.Destination = crptToWindow
    Reportes.SelectionFormula = " {PEDIDOS.PED_NUMFAC} = " & txtNumFac.Text & " AND {PEDIDOS.PED_NUMSER} = '" & txtNumSer.Text & "' AND {PEDIDOS.PED_CODCIA} = '" & LK_CODCIA & "' AND {PEDIDOS.PED_TIPMOV} = 300 AND {PEDIDOS.PED_ESTADO} <> 'E'"
    Reportes.Action = 1
    Exit Sub
Handler:
    MsgBox Err.Description
End Sub

Private Sub cmdIngresar_Click()
Dim iSecuencia As Long
Dim RowsGrd As Integer
On Error GoTo Handler

    If cmdIngresar.Caption = "&Ingresar" Then
        cmdIngresar.Caption = "&Grabar"
        NumCotizacion
        txtFecha.Text = LK_FECHA_DIA
        f1.Enabled = True
        grd.Enabled = True
        txtcli.Enabled = True
        txtcli.SetFocus
    Else
        If ConsisGrl <> 0 Then Exit Sub
        If FlagNumC = 0 Then NumCotizacion
        pub_cadena = "SELECT * FROM CONTROLL"
        CN.Execute "Begin Transaction", rdExecDirect
        Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
        
        RowsGrd = grd.Rows - 1
        For iSecuencia = 1 To RowsGrd
            If Trim(grd.TextMatrix(iSecuencia, 13)) = "X" Then
                ped_llave.AddNew
                AsignaValores iSecuencia, 1
                ped_llave.Update
            End If
        Next iSecuencia
        CN.Execute "Commit Transaction", rdExecDirect
        con_llave.Close
        f1.Enabled = False
        grd.Enabled = False
        txtcli.Enabled = False
        cmdcancelar_Click
    End If
    FlagNumC = 0
    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical, Pub_Titulo
    
     con_llave.Close
     CN.Execute "Rollback Transaction", rdExecDirect
    
End Sub
Private Sub AsignaValores(ByVal fila As Integer, ByVal tipo As Integer)
If tipo = 1 Then
    ped_llave("Ped_CODCIA") = LK_CODCIA
    ped_llave("Ped_NUMSER") = txtNumSer.Text
    ped_llave("Ped_NUMFAC") = Val(txtNumFac.Text)
    ped_llave("Ped_NUMSEC") = fila
    ped_llave("Ped_ESTADO") = "N" '
    ped_llave("Ped_TIPMOV") = 300 '
    'ped_llave ("Ped_FORMA")
    'ped_llave ("Ped_TIEMPO")
    ped_llave("Ped_FBG") = "C" '
    'ped_llave ("Ped_TRANSP")
    'ped_llave ("Ped_CONDI")
    'ped_llave ("Ped_DIAS")
    'ped_llave ("Ped_CODVEN")
    'ped_llave ("Ped_DIRCLI")
    ped_llave("Ped_SITUACION") = " " '
    'ped_llave ("Ped_TIPVTA")
    'ped_llave ("Ped_NUM_UNIDAD")
    'ped_llave ("Ped_DIA_VISITA")
    'ped_llave ("Ped_EQUIV_ACT")
    ped_llave("PED_FECHA_PRO") = LK_FECHA_DIA
End If
    ped_llave("Ped_FECHA") = txtFecha.Text
    ped_llave("Ped_CANTIDAD") = dDouble(grd.TextMatrix(fila, 2)) * dDouble(grd.TextMatrix(fila, 11))
    ped_llave("Ped_PRECIO") = dDouble(grd.TextMatrix(fila, 4))
    ped_llave("Ped_CODUSU") = LK_CODUSU
    ped_llave("Ped_IGV") = PUB_IMPTO
    ped_llave("Ped_BRUTO") = PUB_IMPORTE
    ped_llave("Ped_CODART") = Val(grd.TextMatrix(fila, 12))
    ped_llave("Ped_UNIDAD") = Trim(grd.TextMatrix(fila, 3))
    ped_llave("Ped_EQUIV") = Val(grd.TextMatrix(fila, 11))
    ped_llave("Ped_CODCLIE") = txtcli.TEXTO
    ped_llave("Ped_HORA") = Time
    ped_llave("Ped_DESCTO") = dDouble(grd.TextMatrix(fila, 8))
    ped_llave("Ped_MONEDA") = PUB_DS
    ped_llave("Ped_CONTACTO") = Left(txtContacto.Text, 50)
    ped_llave("Ped_NOMCLIE") = Left(lblcli.Caption, 50)
    ped_llave("Ped_RUCCLIE") = Left(txtruc.Text, 15)
    ped_llave("Ped_OFERTA") = Left(txtOferta.Text, 100)
    ped_llave("Ped_SUBTOTAL") = dDouble(grd.TextMatrix(fila, 10))
    ped_llave("Ped_NUMPRE") = dDouble(grd.TextMatrix(fila, 15))
    ped_llave("Ped_DESCTO_PRE") = dDouble(grd.TextMatrix(fila, 9))
    ped_llave("PED_NUMDOC") = Left(txtNroCot.Text, 15)
    ped_llave("PED_MARCA") = Left(Trim(grd.TextMatrix(fila, 6)), 15)
End Sub

Private Sub Form_Load()
    FormatObjects 0
    FormatGrid
'    carga_venta
'    LlenadoCbo cmdtipo, 65
    txtNumSer.Text = 0
    NumCotizacion
    txtFecha.Text = LK_FECHA_DIA
 '   ColHeaders
End Sub

'Private Sub i_condi_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        moneda.SetFocus
'        RES = SendMessageLong(moneda.hwnd, &H14F, True, 0)
'    End If
'End Sub
'
'Private Sub i_condi_LostFocus()
'    PUB_CODTRA = 2401
'    PUB_SECUENCIA = Val(Trim(Left(i_condi.Text, 2)))
'    SQ_OPER = 1
'    LEER_SUT_LLAVE
'    If SUT_LLAVE.EOF Then Exit Sub
'    pub_signo_car = SUT_LLAVE!SUT_SIGNO_CAR
'End Sub

'Private Sub i_dias_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
''        If UCase(cmdIngreso.Caption) <> "&GRABAR" Then
'           If i_destino.Enabled = True Then
'             i_destino.SetFocus
'             RES = SendMessageLong(i_destino.hwnd, &H14F, True, 0)
'           End If
''        End If
'    End If
'End Sub

Private Sub moneda_Click()
    If Left(moneda.Text, 1) = "S" Then
        stbSubtotales.Panels(4).Text = "S/."
        PUB_DS = "S"
    Else
        stbSubtotales.Panels(4).Text = "US$."
        PUB_DS = "D"
    End If
End Sub

Private Sub moneda_GotFocus()
    If moneda.ListCount = 1 Then moneda_KeyPress 13
End Sub

Private Sub moneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If moneda.ListIndex = -1 Then
            moneda.SetFocus
            Exit Sub
        End If
        txtFecha.SetFocus
    End If
End Sub

'===========================================================================
'============Para Presentacion de Datos de Cliente=============
Private Sub txtcli_Cancel()
    txtcli.TEXTO = ""
    lblcli.Caption = ""
    txtContacto.Text = ""
End Sub

Private Sub txtcli_GetRegistros(ByVal oKeyFind As Variant)
Dim sSql As String
On Error GoTo ErroHandle
    sSql = "SELECT 'Razon Social de la Empresa'=cli_nombre, 'Codigo'=cli_codcliE, 'RUC'=CLI_RUC_ESPOSO, 'Direccion'=cli_casa_direc FROM CLIENTES WHERE CLI_CP='C' AND CLI_NOMBRE LIKE '" & oKeyFind & "%' AND CLI_CODCIA = '" & LK_CODCIA & "' ORDER BY CLI_NOMBRE"
    txtcli.TypeFind = NameField
    txtcli.SetRecordset = OpenSQLForwardOnly(sSql)
    Exit Sub
ErroHandle:
 MsgBox Err.Description
End Sub

Private Sub txtcli_GotFocus()
    txtcli.ZOrder 0
End Sub

Private Sub txtcli_ShowData(ByVal oKey As Variant)
    SQ_OPER = 1
    pu_cp = "C"
    pu_codclie = oKey
    pu_codcia = LK_CODCIA
    LEER_CLI_LLAVE
    If Not cli_llave.EOF Then
        lblcli.Caption = cli_llave("CLI_NOMBRE")
        txtruc.Text = cli_llave("CLI_RUC_ESPOSO")
        txtContacto.Text = cli_llave("CLI_NOMBRE_empresa")
        moneda.ListIndex = 0
        txtFecha.SetFocus
        'RES = SendMessageLong(moneda.hwnd, &H14F, True, 0)
    End If
End Sub
'===========================================================================
'============Para Presentacion de Datos de Empleado=============
Private Sub Txt_key_Cancel()
    Txt_key.TEXTO = ""
    lblven.Caption = ""
End Sub

Private Sub Txt_key_GetRegistros(ByVal oKeyFind As Variant)
Dim sSql As String
On Error GoTo ErroHandle
    sSql = "SELECT  VEM_NOMBRE AS 'Apellidos y Nombres', vem_codven as CODIGO FROM VEMAEST WHERE VEM_CODCIA = '" & LK_CODCIA & "' ORDER BY VEM_NOMBRE"
    Txt_key.TypeFind = NameField
    Txt_key.SetRecordset = OpenSQLForwardOnly(sSql)
    Exit Sub
ErroHandle:
 MsgBox Err.Description
End Sub

Private Sub txt_key_GotFocus()
    Txt_key.ZOrder 0
End Sub

Private Sub Txt_key_ShowData(ByVal oKey As Variant)
    SQ_OPER = 1
    PUB_CODVEN = oKey
    pu_codcia = LK_CODCIA
    LEER_VEN_LLAVE
    If Not ven_llave.EOF Then
        lblven.Caption = ven_llave("VEM_NOMBRE")
        cmdtipo.SetFocus
        RES = SendMessageLong(cmdtipo.hwnd, &H14F, True, 0)
    End If
End Sub

Private Sub FormatGrid()
    ReDim ColumnNext(1 To 14)
    With grd
        .Clear
        .Rows = 2
        .RowHeightMin = 280
        .Visible = True
        .Width = 11770
        .Height = 4530
        .Cols = 17
        .FormatString = "Descripción|Código|Cantidad|Unidad|Precio|Division|Familia||||SubTotal"
        .ColWidth(0) = 4500 'descripcion
        .ColWidth(1) = 1200 'codigo
        .ColWidth(2) = 900 'cantidad
        .ColWidth(3) = 1000 'unidad
        .ColWidth(4) = 900 'precio
        .ColWidth(5) = 900 'division
        .ColWidth(6) = 900 'familia
        .ColWidth(7) = 0 ' dispo
        .ColWidth(8) = 0  '% descuento
        .ColWidth(9) = 0 ' descuento
        .ColWidth(10) = 900 'subtotal
        .ColWidth(11) = 0 'equivalencia
        .ColWidth(12) = 0 'cod_art
        .ColWidth(13) = 0 'graba(X) o no( )
        .ColWidth(14) = 0 'secuencia DEL DETALLE
        .ColWidth(15) = 0 'secuencia DEL PRECIO
        .ColWidth(16) = 0 'prec reposicion segun unidad
        
        ColumnNext(1) = 1
        ColumnNext(2) = 1
        ColumnNext(3) = 1
        ColumnNext(4) = -1
        ColumnNext(5) = 0
        ColumnNext(6) = 0
        ColumnNext(7) = 0
        ColumnNext(10) = 0
    End With
End Sub
Private Sub FormatObjects(ByVal lFlag As Integer)
Dim gSpaceS As Long
Dim gSpaceI As Long
Dim lHeigth As Long
    If FlagLvw = 0 Then Exit Sub
    If lFlag = 1 Then
        gSpaceS = txtFindRegistro.Top
        'gSpaceI = grd.Height - (txtFindRegistro.Height + txtFindRegistro.Top)
        If gSpaceS < 5000 Then 'gSpaceS = gSpaceI Or gSpaceI > gSpaceS Then
            lvw.Top = txtFindRegistro.Height + txtFindRegistro.Top
        Else 'If gSpaceS > gSpaceI Then
            lvw.Top = txtFindRegistro.Top - lvw.Height
        End If
        lvw.Width = 11300
        lvw.Left = 500
        'lvw.ColumnHeaders.Item(1).Width = txtFindRegistro.Width
        lvw.Visible = True
        lvw.Height = 2100
        lvw.ZOrder 0
        txtFindRegistro.Visible = True
    ElseIf lFlag = 2 Then
        'lvw.SelectedItem.Text = ""
        'lvw.SelectedItem.ListSubItems(1).Text = ""
        txtFindRegistro.Visible = False
        lvw.Visible = False
        lvw.ColumnHeaders.Clear
    End If
End Sub
Private Sub CalculaTotales1(ByVal RowA As Long, ByVal ColA As Long)
Dim iRow As Long
Dim iRows As Integer
    On Error GoTo Handler
    PUB_SUBTOTAL = dDouble(grd.TextMatrix(RowA, 2)) * dDouble(grd.TextMatrix(RowA, 4))
    grd.TextMatrix(RowA, 10) = Format(PUB_SUBTOTAL, "##0.000")
    PUB_IMPORTE = 0
    PUB_IMPORTE_AMORT = 0
    PUB_SUBTOTAL = 0
    PUB_IMPTO = 0
    iRows = grd.Rows
    For iRow = 1 To iRows - 1
        PUB_IMPORTE = dDouble(grd.TextMatrix(iRow, 10)) + PUB_IMPORTE
        PUB_SUBTOTAL = PUB_IMPORTE / (1 + LK_IGV / 100)
        PUB_IMPTO = PUB_IMPORTE - PUB_SUBTOTAL
    Next iRow
    stbSubtotales.Panels(1).Text = Format(PUB_SUBTOTAL, "##0.000")
    stbSubtotales.Panels(3).Text = Format(PUB_IMPTO, "##0.000")
    stbSubtotales.Panels(5).Text = Format(PUB_IMPORTE, "##0.000")
    Exit Sub
Handler:
End Sub
Private Sub CalculaSubTotales1(ByVal RowA As Long, ByVal ColA As Long)
On Error GoTo Handler
    PUB_SUBTOTAL = dDouble(grd.TextMatrix(RowA, 2)) * dDouble(grd.TextMatrix(RowA, 4))
    grd.TextMatrix(RowA, 10) = Format(PUB_SUBTOTAL, "##0.000")
    Exit Sub
Handler:
End Sub
'Private Sub UbicaCombo()
'    cmbprecios.Visible = True
'    cbmTop = grd.Top + (iRowGrd * 360) '180
'    cbmLeft = grd.Left + 7100
'    cmbprecios.Left = cbmLeft
'    cmbprecios.Top = cbmTop
'End Sub
Private Sub NumCotizacion()

    archi = "SELECT MAX(PED_NUMFAC) AS NUMFAC FROM PEDIDOS WHERE PED_TIPMOV=300 AND Ped_FBG = 'C' AND PED_CODCIA = '" & LK_CODCIA & "'"
    Set PSX = CN.CreateQuery("", archi)
    Set X = PSX.OpenResultset(rdOpenKeyset)
    X.Requery
    If X.EOF Then
        txtNumFac.Text = 1
    Else
        txtNumFac.Text = Nulo_Valor0(X("NUMFAC")) + 1
    End If
End Sub
'===============================================================
Private Sub grd_EnterCell()
    ColorCellDefault = grd.CellBackColor
    If Trim(grd.TextMatrix(grd.Row, 12)) = "" Then Exit Sub
    If grd.COL = 3 Then
        LOADUNIDADES
    ElseIf grd.COL = 4 Then
        LOADPRECIOS
    End If
End Sub
Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    'ColorCell grd.Row, grd.Col
    If KeyCode = 13 Then
        FLAG = 1 'Editar Celda
        If grd.COL = 2 Then
            txt = ""
            txt.Visible = False
            txt.Top = grd.Top + grd.CellTop - 15
            txt.Left = grd.Left + grd.CellLeft - 15
            txt.Width = grd.CellWidth + 30
            'txt.Height = grd.CellHeight
            txt.Text = Trim(grd.Text)
            txt.Visible = True
            If Len(grd.Text) > 0 Then Azul txt, txt
            txt.SetFocus
'        ElseIf ColumnControl(grd.COL) = eMaskEdBox Then
'            txt2.Text = "__/__/____"
'            txt2.Visible = False
'            txt2.Top = grd.Top + grd.CellTop - 15
'            txt2.Left = grd.Left + grd.CellLeft - 15
'            txt2.Width = grd.CellWidth + 30
'            'txt.Height = grd.CellHeight
'            txt2.Text = IIf(IsDate(grd.Text), grd.Text, "__/__/____")
'            txt2.Visible = True
'            txt2.SetFocus
        ElseIf grd.COL = 3 Or grd.COL = 4 Then
            If cbo(grd.COL).ListCount = 0 Then Exit Sub
            cbo(grd.COL).Visible = False
            cbo(grd.COL).Top = grd.Top + grd.CellTop - 15
            cbo(grd.COL).Left = grd.Left + grd.CellLeft - 15
            cbo(grd.COL).Width = grd.CellWidth + 30
            cbo(grd.COL).Visible = True
            cbo(grd.COL).SetFocus
            RES = SendMessageLong(cbo(grd.COL).hwnd, &H14F, True, 0)
        ElseIf grd.COL = 1 Then
            txtFindRegistro.Text = ""
            txtFindRegistro.Visible = False
            txtFindRegistro.Top = grd.Top + grd.CellTop - 15
            txtFindRegistro.Left = grd.Left + grd.CellLeft - 15
            txtFindRegistro.Width = grd.CellWidth + 30
            'txt.Height = grd.CellHeight
            txtFindRegistro.Text = grd.Text
            txtFindRegistro.Visible = True
            If Len(grd.Text) > 0 Then Azul txtFindRegistro, txtFindRegistro
            txtFindRegistro.SetFocus
        End If
    ElseIf KeyCode = 46 Then
        Dim vbresp As Integer
        vbresp = MsgBox("Esta seguro de Eliminar el Registro", vbQuestion + vbYesNo, Pub_Titulo)
        If vbresp = vbYes Then
            grd.RemoveItem grd.Row
        End If
        grd.SetFocus
        CalculaSubTotales1 grd.Row, 4
        CalculaTotales1 grd.Row, 4
    ElseIf KeyCode = 45 And grd.COL = 4 Then
        Dim dPrecio As String
        dPrecio = InputBox("Ingrese el Precio", Pub_Titulo)
        If Trim(dPrecio) = "" Then Exit Sub
        If dDouble(dPrecio) = 0 Then
            MsgBox "Ingrese un valor correcto"
            Exit Sub
        End If
        If dDouble(dPrecio) < dDouble(grd.TextMatrix(grd.Row, 16)) Then
            MsgBox "El precio ingresado es menor al precios de Costo", vbInformation, Pub_Titulo
            dPrecio = "0.000"
        End If
        grd.TextMatrix(grd.Row, 4) = dDouble(dPrecio)
        CalculaSubTotales1 grd.Row, 4
        CalculaTotales1 grd.Row, 4
        ValRowGrd grd.Row
        If ColumnNext(grd.COL) = -1 And Trim(grd.TextMatrix(grd.Row, 13)) = "X" And (grd.Row = grd.Rows - 1) Then GoTo OtraFila
        If grd.Cols - 1 > grd.COL Then
            grd.COL = grd.COL + ColumnNext(grd.COL)
        ElseIf grd.Cols = grd.COL + 1 Then
            If grd.Row = grd.Rows - 1 Then
            
OtraFila:
                If grd.Rows = 26 Then
                    MsgBox "Se llego al tope de items. maximo 25 items", vbInformation, Pub_Titulo
                    lblItems.Caption = "Nº de Items " & grd.Rows - 1
                    Exit Sub
                End If
                grd.Rows = grd.Rows + 1
            End If
            grd.COL = 1
            grd.Row = grd.Row + 1
        End If
        lblItems.Caption = "Nº de Items " & grd.Rows - 2
        grd.SetFocus
    End If
End Sub
'====================================================
'TXT
Private Sub txt_GotFocus()
    txt.SelStart = 0
    txt.SelLength = Len(txt.Text)
End Sub
Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        txt.Text = ""
        txt.Visible = False
    End If
End Sub
Private Sub Txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And FLAG = 1 Then
        FLAG = 0 'Desactiva Edicion
        grd.Text = txt.Text
        txt.Visible = False
        grd.SetFocus
        If grd.COL = 2 Then
            CalculaSubTotales1 grd.Row, grd.COL
            CalculaTotales1 grd.Row, grd.COL
        End If
        ValRowGrd grd.Row
        If grd.Cols - 1 > grd.COL Then
            grd.COL = grd.COL + ColumnNext(grd.COL)
        ElseIf grd.Cols = grd.COL + 1 Then
            If grd.Row = grd.Rows - 1 Then
                grd.Rows = grd.Rows + 1
            End If
            grd.COL = 1
            grd.Row = grd.Row + 1
        End If
        'ColorCell grd.Row, grd.Col
        ValRowGrd grd.Row
        
        If grd.COL = 7 Then
        If ColumnNext(grd.COL) = -1 And Trim(grd.TextMatrix(grd.Row, 13)) = "X" And (grd.Row = grd.Rows - 1) Then GoTo OtraFila
        
            If grd.Cols - 1 > grd.COL Then
                grd.COL = grd.COL + ColumnNext(grd.COL)
            ElseIf grd.Cols = grd.COL + 1 Then
                If grd.Row = grd.Rows - 1 Then
                
OtraFila:
                    If grd.Rows = 26 Then
                        MsgBox "Se llego al tope de items. maximo 25 items", vbInformation, Pub_Titulo
                        lblItems.Caption = "Nº de Items " & grd.Rows - 1
                        Exit Sub
                    End If
                    grd.Rows = grd.Rows + 1
                End If
                grd.COL = 1
                grd.Row = grd.Row + 1
            End If
        End If
        lblItems.Caption = "Nº de Items " & grd.Rows - 2
        grd.SetFocus
    End If
    If grd.COL = 4 Or grd.COL = 2 Then
        KeyAscii = vNumeric(KeyAscii)
    ElseIf grd.COL = 6 Then
    
    End If
End Sub
'====================================================
'txt2
'Private Sub txt2_GotFocus()
'    txt2.SelStart = 0
'    txt2.SelLength = Len(txt2.Text)
'End Sub
'Private Sub txt2_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        If IsDate(txt2.Text) Or txt2.Text = "__/__/____" Then
'            SendKeys "{tab}"
'            KeyCode = 0
'        Else
'            MsgBox "El Dato ingresado no corresponde a una Fecha Valida!!!", vbCritical, "Control"
'            txt2.SelStart = 0
'            txt2.SelLength = 10
'            txt2.SetFocus
'        End If
'    End If
'End Sub
'Private Sub txt2_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 And FLAG = 1 Then
'        FLAG = 0 'Desactiva Edicion
'        grd.Text = IIf(txt2.Text = "__/__/____", "", txt2.Text)
'        txt2.Visible = False
'        'Call grd_LeaveCell
'        If grd.Cols - 1 > grd.COL Then
'            grd.COL = grd.COL + ColumnNext(grd.COL)
'        ElseIf grd.Cols = grd.COL + 1 Then
'            If grd.Row = grd.Rows - 1 Then
'                grd.Rows = grd.Rows + 1
'            End If
'            grd.COL = 1
'            grd.Row = grd.Row + 1
'        End If
'        'ColorCell grd.Row, grd.Col
'    End If
'End Sub
'============================================
'txtfindregistro
Private Sub txtFindRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
Dim liNumItems As Integer
    If LK_FLAG_ALTERNO = "A" Then Exit Sub
    liNumItems = lvw.ListItems.count
    If liNumItems = 0 Then Exit Sub
    If KeyCode = 40 Then 'ABAJO
        liIndexRegAct = liIndexRegAct + 1
        If liIndexRegAct > liNumItems Then liIndexRegAct = liNumItems
    ElseIf KeyCode = 38 Then 'ARRIBA
        liIndexRegAct = liIndexRegAct - 1
        If liIndexRegAct <= 0 Then liIndexRegAct = 1
    ElseIf KeyCode = 33 Then '33 RE PAG
        liIndexRegAct = liIndexRegAct - 10
        If liIndexRegAct <= 0 Then liIndexRegAct = 1
    ElseIf KeyCode = 34 Then ' 34 AV PAG
        liIndexRegAct = liIndexRegAct + 10
        If liIndexRegAct > liNumItems Then liIndexRegAct = liNumItems
    End If
If liIndexRegAct = 0 Then Exit Sub
    lvw.ListItems.Item(liIndexRegAct).EnsureVisible
    lvw.ListItems.Item(liIndexRegAct).Selected = True
    If KeyCode = 33 Or KeyCode = 34 Or KeyCode = 38 Or KeyCode = 40 Then txtFindRegistro.Text = lvw.ListItems(liIndexRegAct).Text
    If KeyCode = 13 Then
        Call lvw_DblClick
    End If
End Sub
Private Sub txtFindRegistro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If LK_FLAG_ALTERNO = " " Then
            ' CARGA ARTICULO POR DESCRICPION EN LIST VIEW
            If FlagTypeFind = 1 Then
                FLAG = 0 'Desactiva Edicion
                grd.Text = txtFindRegistro.Text
                txtFindRegistro.Visible = False
                grd.SetFocus
                'Call grd_LeaveCell
                If grd.Cols - 1 > grd.COL Then
                    grd.COL = grd.COL + ColumnNext(grd.COL)
                ElseIf grd.Cols = grd.COL + 1 Then
                    If grd.Row = grd.Rows - 1 Then
                        grd.Rows = grd.Rows + 1
                    End If
                    grd.COL = 1
                    grd.Row = grd.Row + 1
                End If
                'ColorCell grd.Row, grd.Col
                'Call FormatObjects(2)
                FLAG = 1
            End If
        ElseIf LK_FLAG_ALTERNO = "A" Then
            'CARGA ARTICULO POR CODIGO ALTERNO
            If Trim(txtFindRegistro.Text) = "" Then Exit Sub
            
            If ExistDato(txtFindRegistro.Text, 1, grd.Row, grd) Then
                MsgBox "El codigo ya se encuentra en Lista!!!", vbInformation, Pub_Titulo
                Call ClearRow(grd.Row, grd)
                txtFindRegistro.SetFocus
                Exit Sub
            End If
            
            pu_alterno = txtFindRegistro.Text
            pu_codcia = LK_CODCIA
            SQ_OPER = 3
            LEER_ART_LLAVE
            If art_llave_alt.EOF Then
'                MsgBox "Error Grave en Arti", 46, Pub_Titulo
                txtFindRegistro.Text = ""
                txtFindRegistro.Visible = False
                Exit Sub
            Else
                grd.TextMatrix(grd.Row, 0) = Trim(art_llave_alt("ART_NOMBRE"))
                grd.TextMatrix(grd.Row, 12) = art_llave_alt("ART_KEY")
                PUB_KEY = art_llave_alt("ART_KEY")
            End If
            grd.Text = txtFindRegistro.Text
            LOADUNIDADES
            LOADLABOALT
            'LOADPROCE
            LOADMARCA
            txtFindRegistro.Visible = False
            grd.SetFocus
            ValRowGrd grd.Row
            If grd.Cols - 1 > grd.COL Then
                grd.COL = grd.COL + ColumnNext(grd.COL)
            ElseIf grd.Cols = grd.COL + 1 Then
                If grd.Row = grd.Rows - 1 Then
                    grd.Rows = grd.Rows + 1
                End If
                grd.COL = 1
                grd.Row = grd.Row + 1
            End If
        End If
    End If
End Sub
Private Sub txtFindRegistro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        lvw.ListItems.Clear
        liIndexRegAct = 0
        txtFindRegistro = ""
        txtFindRegistro.Visible = False
        FormatObjects 2
        grd.SetFocus
    End If
    If KeyCode >= 37 And KeyCode <= 40 Or (KeyCode = 33 Or KeyCode = 34) Then Exit Sub
    If Len(txtFindRegistro) = 1 Then
        'RaiseEvent GetRecordSet(Trim(txtFindRegistro))
        'LoadHeader
        If LK_FLAG_ALTERNO = "A" Then Exit Sub
        LlenaLvw
        FormatObjects 1
    End If
    FlagTypeFind = 1
    If Len(txtFindRegistro) > 1 Then FindItem Trim(txtFindRegistro)
End Sub
'========================================================================
'lvw
Private Sub lvw_DblClick()
    FlagTypeFind = 2
    txtFindRegistro.Text = lvw.SelectedItem.ListSubItems(1).Text
    
    If ExistDato(txtFindRegistro.Text, 1, grd.Row, grd) Then
        FormatObjects 2
        MsgBox "El codigo ya se encuentra en Lista!!!", vbInformation, Pub_Titulo
        Call ClearRow(grd.Row, grd)
        Exit Sub
    End If
    
    PUB_KEY = lvw.SelectedItem.ListSubItems(3).Text
    grd.TextMatrix(grd.Row, 12) = PUB_KEY
    pu_codcia = LK_CODCIA
    SQ_OPER = 1
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
        MsgBox "Error Grave en Arti", 46, Pub_Titulo
    Else
        grd.TextMatrix(grd.Row, 0) = art_LLAVE("ART_NOMBRE")
        grd.TextMatrix(grd.Row, 12) = art_LLAVE("ART_KEY")
    End If
    FLAG = 0 'Desactiva Edicion
    grd.Text = txtFindRegistro.Text
    LOADUNIDADES
    LOADLABO
    'LOADPROCE
    LOADMARCA
    txtFindRegistro.Visible = False
    grd.SetFocus
    ValRowGrd grd.Row
    'Call grd_LeaveCell
    If grd.Cols - 1 > grd.COL Then
        grd.COL = grd.COL + ColumnNext(grd.COL)
    ElseIf grd.Cols = grd.COL + 1 Then
        If grd.Row = grd.Rows - 1 Then
            grd.Rows = grd.Rows + 1
        End If
        grd.COL = 1
        grd.Row = grd.Row + 1
    End If
    'ColorCell grd.Row, grd.Col
    Call FormatObjects(2)
    FLAG = 1
End Sub
Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If txtFindRegistro.Visible Then
        Item.EnsureVisible
        Item.Selected = True
        liIndexRegAct = Item.Index
    End If
End Sub
'===================================================================
'cbo
Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
End Sub
'=====================================================================
'funciones
Private Sub FindItem(ByVal sDato As String)
Dim intSelectedOption As Integer
Dim strFindMe As String
    
    intSelectedOption = lvwText
    Set loListItem = lvw.FindItem(sDato, intSelectedOption, , lvwPartial)
    If loListItem Is Nothing Then
        Exit Sub
    Else
        loListItem.EnsureVisible
        loListItem.Selected = True
        liIndexRegAct = loListItem.Index
        FlagTypeFind = 2
    End If
End Sub
Private Sub LlenaLvw()
Dim i As Long
Dim J As Integer
Dim var As Variant

    var = Asc(txtFindRegistro.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" And Val(txtFindRegistro.Text) <> 0 Then
        'archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND  ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' AND ART_ALTERNO BETWEEN '" & Trim(txtFindRegistro.Text) & "%' AND  '" & var & "' ORDER BY ART_ALTERNO"
        archi = "SELECT ARTI.ART_KEY, ARTI.ART_NOMBRE, ARTI.ART_ALTERNO, ARTICULO.ARM_STOCK, PRECIOS.PRE_EQUIV, TABLAS.TAB_TIPREG, TABLAS_1.TAB_TIPREG AS Expr1, PRECIOS.PRE_FLAG_UNIDAD, ARTI.ART_CODCIA, dbo.TABLAS.TAB_NOMLARGO AS Labo, TABLAS_1.TAB_NOMLARGO AS Marca "
        archi = archi & "FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA INNER JOIN TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB INNER JOIN TABLAS TABLAS_1 ON ARTI.ART_CODCIA = TABLAS_1.TAB_CODCIA AND ARTI.ART_SUBFAM = TABLAS_1.TAB_NUMTAB "
        archi = archi & "WHERE (TABLAS.TAB_TIPREG = 122) AND (TABLAS_1.TAB_TIPREG = 123) AND (PRECIOS.PRE_FLAG_UNIDAD = 'A') AND ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_CALIDAD = 1 AND ARTI.ART_FLAG_STOCK = 'M' AND ARTI.ART_ALTERNO BETWEEN '" & Trim(txtFindRegistro.Text) & "%' AND  '" & var & "' AND ART_SITUACION <> 1 ORDER BY ARTI.ART_ALTERNO"
    Else
    '    archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND  (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' AND ART_NOMBRE BETWEEN '" & Trim(txtFindRegistro.Text) & "%' AND  '" & var & "' ORDER BY ART_NOMBRE"
        archi = "SELECT ARTI.ART_KEY, ARTI.ART_NOMBRE, ARTI.ART_ALTERNO, ARTICULO.ARM_STOCK, PRECIOS.PRE_EQUIV, TABLAS.TAB_TIPREG, TABLAS_1.TAB_TIPREG AS Expr1, PRECIOS.PRE_FLAG_UNIDAD, ARTI.ART_CODCIA, dbo.TABLAS.TAB_NOMLARGO AS Labo, TABLAS_1.TAB_NOMLARGO AS Marca  "
        archi = archi & "FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA INNER JOIN TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB INNER JOIN TABLAS TABLAS_1 ON ARTI.ART_CODCIA = TABLAS_1.TAB_CODCIA AND ARTI.ART_SUBFAM = TABLAS_1.TAB_NUMTAB "
        archi = archi & "WHERE (TABLAS.TAB_TIPREG = 122) AND (TABLAS_1.TAB_TIPREG = 123)  AND (PRECIOS.PRE_FLAG_UNIDAD = 'A') AND ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_CALIDAD = 1 AND ARTI.ART_FLAG_STOCK = 'M' AND ARTI.ART_NOMBRE BETWEEN '" & Trim(txtFindRegistro.Text) & "%' AND  '" & var & "' AND ART_SITUACION <> 1 ORDER BY ARTI.ART_NOMBRE"
    End If
    
    Set PSX = CN.CreateQuery("", archi)
    Set X = PSX.OpenResultset(rdOpenForwardOnly)
    X.Requery
    lvw.ListItems.Clear
    ColHeaders
    Do While Not X.EOF
        i = i + 1
        Set loListItem = lvw.ListItems.Add(, "k" & CStr(i) & CStr(X("ART_ALTERNO")), Trim(X("ART_NOMBRE")))
        loListItem.ListSubItems.Add Key:="Codigo" & CStr(J), Text:=Trim(X("ART_ALTERNO"))
        loListItem.ListSubItems.Add Key:="Stock" & CStr(J), Text:=Format((X("ARM_STOCK") / X("PRE_EQUIV")), "0.000")
        loListItem.ListSubItems.Add Key:="key" & CStr(J), Text:=Trim(X("ART_KEY"))
        
        loListItem.ListSubItems.Add Key:="Fam" & CStr(J), Text:=Trim(X("Labo"))
        'loListItem.ListSubItems.Add Key:="SubFam" & CStr(J), Text:=Trim(X("Proce"))
        loListItem.ListSubItems.Add Key:="Marca" & CStr(J), Text:=Trim(X("Marca"))
        X.MoveNext
    Loop
    If i > 0 Then liIndexRegAct = 1
    FlagLvw = i
End Sub
Private Sub ColHeaders()
On Error Resume Next
    lvw.ColumnHeaders.Add , "Descripcion", "Descripcion", 5000, lvwColumnLeft
    lvw.ColumnHeaders.Add , "Codigo", "Codigo", 1000, lvwColumnLeft
    lvw.ColumnHeaders.Add , "Stock", "Stock", 1000, lvwColumnRight
    lvw.ColumnHeaders.Add , "Key", "Key", 0, lvwColumnRight
    lvw.ColumnHeaders.Add , "Familia", "Familia", 2000, lvwColumnLeft
    lvw.ColumnHeaders.Add , "Subfamilia", "Subfamilia", 2000, lvwColumnRight
    'lvw.ColumnHeaders.Add , "Marca", "Marca", 1500, lvwColumnRight
End Sub

Private Sub ValRowGrd(ByVal lRow As Integer)
On Error GoTo Handler
Dim iErrors As Integer
    
    If Trim(grd.TextMatrix(lRow, 0)) = "" Or Trim(grd.TextMatrix(lRow, 1)) = "" Or Val(grd.TextMatrix(lRow, 10)) = 0 Then
        iErrors = iErrors + 1
    End If
    If iErrors = 0 Then
        grd.TextMatrix(lRow, 13) = "X"
    Else
        grd.TextMatrix(lRow, 13) = " "
    End If
    Exit Sub
Handler:
End Sub

Private Sub txtNroCot_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtOferta.SetFocus
    End If
End Sub

Private Sub txtnumfac_KeyPress(KeyAscii As Integer)
Dim vbresp As Integer
    If KeyAscii = 13 Then
        FlagNumC = 0
        If cmdIngresar.Caption = "&Grabar" Then
            vbresp = MsgBox("Desea inicializar las ordenes de compra desde este numero????", vbYesNo + vbInformation, Pub_Titulo)
            If vbresp = vbYes Then
                FlagNumC = 1
                txtcli.SetFocus
                Exit Sub
            End If
        End If
        LoadCotizacion
    Else
        KeyAscii = vInteger(KeyAscii)
    End If
    
End Sub

Private Sub txtOferta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grd.SetFocus
    End If
End Sub
Private Function ConsisGrl() As Integer
    If Consis1(txtcli.TEXTO) = 1 Then
        MsgBox "Cliente no existe....", 46, Pub_Titulo
        ConsisGrl = 1
        txtcli.SetFoco
        Exit Function
    End If
    If consis2(txtNroCot.Text) = 1 Then
        MsgBox "Falta Ingresar nro de Cotización....", 46, Pub_Titulo
        ConsisGrl = 1
        txtNroCot.SetFocus
        Exit Function
    End If
    If Consis3(grd) = 1 Then
        MsgBox "Falta Ingresar Articulos....", 46, Pub_Titulo
        ConsisGrl = 1
        grd.SetFocus
        Exit Function
    End If
    If Consis4(txtFecha.Text) = 1 Then
        MsgBox "Falta ingresar la fecha....", 46, Pub_Titulo
        ConsisGrl = 1
        txtFecha.SetFocus
        Exit Function
    End If
End Function


Private Sub LoadCotizacion()
Dim iSec As Integer

    SQ_OPER = 1
    PUB_TIPMOV = 300
    pu_codcia = LK_CODCIA
    PUB_PEDSER = txtNumSer.Text
    PUB_PEDFAC = Val(txtNumFac.Text)
    LEER_PED_LLAVE
    If Not ped_llave.EOF Then
        If Trim(ped_llave("PED_ESTADO")) = "E" Then
            MsgBox "DOCUMENTO ESTA EXTORNADO", vbInformation, Pub_Titulo
        ElseIf Trim(ped_llave("PED_SITUACION")) = "P" Then
            MsgBox "DOCUMENTO ESTA PROCESADO", vbInformation, Pub_Titulo
        Else
            cmdIngresar.Enabled = False
            cmdAnular.Enabled = True
            cmdConsulta.Enabled = True
        End If
        cmdCancelar.Enabled = True
        txtcli.Enabled = True
        f1.Enabled = True
        grd.Enabled = True
        chkCambiar.Value = 0
        txtNumFac.Locked = True
        txtFecha.Text = ped_llave("Ped_FECHA")
        txtcli.TEXTO = ped_llave("Ped_CODCLIE")
        Call txtcli_ShowData(txtcli.TEXTO)  'Load Cliente
        
        lblcli.Caption = ped_llave("Ped_NOMCLIE")
        'txtContacto.Text = ped_llave("Ped_CONTACTO")
        txtruc.Text = ped_llave("Ped_RUCCLIE")
        txtOferta.Text = ped_llave("Ped_OFERTA")
        txtNroCot.Text = ped_llave("PED_NUMDOC")
        
        PUB_IMPTO = ped_llave("Ped_IGV")
        PUB_IMPORTE = ped_llave("Ped_BRUTO")
        PUB_SUBTOTAL = ped_llave("Ped_BRUTO") - ped_llave("Ped_IGV")
        
        stbSubtotales.Panels(1).Text = Format(PUB_SUBTOTAL, "##0.000")
        stbSubtotales.Panels(3).Text = Format(PUB_IMPTO, "##0.000")
        stbSubtotales.Panels(5).Text = Format(PUB_IMPORTE, "##0.000")
        
        RES = SendMessageLong(moneda.hwnd, &H14F, False, 0)
        
        If ped_llave("Ped_MONEDA") = "S" Then
            moneda.ListIndex = 0
        ElseIf ped_llave("Ped_MONEDA") = "D" Then
            moneda.ListIndex = 1
        End If
        
        FormatGrid
        grd.Rows = ped_llave.RowCount + 1
        Do While Not ped_llave.EOF
            iSec = iSec + 1
        '    ped_llave("Ped_ESTADO") = "N"
            grd.TextMatrix(iSec, 2) = Format(Nulo_Valor0(ped_llave("Ped_CANTIDAD")) / Nulo_Valor0(ped_llave("Ped_EQUIV")), "0.000")
            grd.TextMatrix(iSec, 3) = ped_llave("Ped_UNIDAD")
            grd.TextMatrix(iSec, 4) = Format(ped_llave("Ped_PRECIO"), "0.000")
            grd.TextMatrix(iSec, 6) = Nulo_Valors(ped_llave("Ped_MARCA"))
            grd.TextMatrix(iSec, 8) = ped_llave("Ped_DESCTO")
            grd.TextMatrix(iSec, 10) = ped_llave("Ped_SUBTOTAL")
            grd.TextMatrix(iSec, 11) = ped_llave("Ped_EQUIV")
            grd.TextMatrix(iSec, 12) = ped_llave("Ped_CODART")
            grd.TextMatrix(iSec, 14) = ped_llave("Ped_NUMSEC")
            grd.TextMatrix(iSec, 15) = ped_llave("Ped_NUMPRE")
            SQ_OPER = 1
            PUB_KEY = ped_llave("Ped_CODART")
            pu_codcia = LK_CODCIA
            LEER_ART_LLAVE
            If Not art_LLAVE.EOF Then
                grd.TextMatrix(iSec, 1) = art_LLAVE("ART_ALTERNO")
                grd.TextMatrix(iSec, 0) = art_LLAVE("ART_NOMBRE")
                grd.Row = iSec
                LOADMARCA
                LOADLABO
'                LOADPROCE
            Else
                MsgBox "Error grave en Arti...", 46, Pub_Titulo
            End If
            ped_llave.MoveNext
        Loop
        txtcli.SetFocus
        cmdConsulta.Caption = "Grabar"
    End If
End Sub
Private Sub LOADUNIDADES()
    SQ_OPER = 2
    pu_codcia = LK_CODCIA
    PUB_CODART = Val(grd.TextMatrix(grd.Row, 12))
    LEER_PRE_LLAVE
    If Not pre_mayor.EOF Then
        cbo(3).Clear
        Do While Not pre_mayor.EOF
            cbo(3).AddItem pre_mayor("Pre_Unidad") & String(60, " ") & pre_mayor("PRE_SECUENCIA")
            pre_mayor.MoveNext
        Loop
    Else
        MsgBox "Error grave en precios....", 46, Pub_Titulo
        Exit Sub
    End If
End Sub
Private Sub LOADPRECIOS()
On Error GoTo Handler
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_CODART = Val(grd.TextMatrix(grd.Row, 12))
    PUB_SECUEN = Val(IIf(Trim(grd.TextMatrix(grd.Row, 15)) = "", "-1", grd.TextMatrix(grd.Row, 15)))
    LEER_PRE_LLAVE
    If Not pre_llave.EOF Then
        cbo(4).Clear
        grd.TextMatrix(grd.Row, 11) = pre_llave("PRE_EQUIV")
        grd.TextMatrix(grd.Row, 16) = pre_llave("PRE_COSTO")
        If PUB_DS = "S" Then
            If pre_llave("PRE_PRE1") <> 0 Then
                cbo(4).AddItem pre_llave("PRE_PRE1")
            End If
            If pre_llave("PRE_PRE2") <> 0 Then
                cbo(4).AddItem pre_llave("PRE_PRE2")
            End If
            If pre_llave("PRE_PRE3") <> 0 Then
                cbo(4).AddItem pre_llave("PRE_PRE3")
            End If
            If pre_llave("PRE_PRE4") <> 0 Then
                cbo(4).AddItem pre_llave("PRE_PRE4")
            End If
            If pre_llave("PRE_PRE5") <> 0 Then
                cbo(4).AddItem pre_llave("PRE_PRE5")
            End If
            If pre_llave("PRE_PRE6") <> 0 Then
                cbo(4).AddItem pre_llave("PRE_PRE6")
            End If
        ElseIf PUB_DS = "D" Then
            If pre_llave("PRE_PRE11") <> 0 Then
                cbo(4).AddItem pre_llave("PRE_PRE11")
            End If
            If pre_llave("PRE_PRE22") <> 0 Then
                cbo(4).AddItem pre_llave("PRE_PRE22")
            End If
            If pre_llave("PRE_PRE33") <> 0 Then
                cbo(4).AddItem pre_llave("PRE_PRE33")
            End If
            If pre_llave("PRE_PRE44") <> 0 Then
                cbo(4).AddItem pre_llave("PRE_PRE44")
            End If
            If pre_llave("PRE_PRE55") <> 0 Then
                cbo(4).AddItem pre_llave("PRE_PRE55")
            End If
            If pre_llave("PRE_PRE66") <> 0 Then
                cbo(4).AddItem pre_llave("PRE_PRE66")
            End If
        End If
    Else
        'MsgBox "Error grave en precios", 46, Pub_Titulo
    End If
    Exit Sub
Handler:
End Sub
Private Sub LOADLABOALT()
    SQ_OPER = 1
    PUB_TIPREG = 122
    PUB_NUMTAB = art_llave_alt("ART_FAMILIA")
    PUB_CODCIA = LK_CODCIA
    LEER_TAB_LLAVE
    If Not tab_llave.EOF Then
        grd.TextMatrix(grd.Row, 5) = tab_llave("tab_nomlargo")
    End If
End Sub
Private Sub LOADLABO()
    SQ_OPER = 1
    PUB_TIPREG = 122
    PUB_NUMTAB = art_LLAVE("ART_FAMILIA")
    PUB_CODCIA = LK_CODCIA
    LEER_TAB_LLAVE
    If Not tab_llave.EOF Then
        grd.TextMatrix(grd.Row, 5) = tab_llave("tab_nomlargo")
    End If
End Sub

Private Sub LOADMARCA()
    SQ_OPER = 1
    PUB_TIPREG = 123
    PUB_NUMTAB = art_LLAVE("ART_SUBFAM")
    PUB_CODCIA = LK_CODCIA
    LEER_TAB_LLAVE
    If Not tab_llave.EOF Then
        grd.TextMatrix(grd.Row, 6) = tab_llave("tab_nomlargo")
    End If
End Sub
'Private Sub LOADPROCE()
'    SQ_OPER = 1
'    PUB_TIPREG = 132
'    PUB_NUMTAB = art_LLAVE("ART_MARCA")
'    PUB_CODCIA = LK_CODCIA
'    LEER_TAB_LLAVE
'    If Not tab_llave.EOF Then
'        grd.TextMatrix(grd.Row, 7) = tab_llave("tab_nomlargo")
'    End If
'End Sub
