VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F6E4F630-E903-11D5-8BB9-0080AD40A177}#1.18#0"; "OSControlsUser.ocx"
Begin VB.Form frmOCCliente 
   Caption         =   "Registro de Orden de Compra  ( Cliente )"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "frmOCCliente.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
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
      Left            =   5910
      TabIndex        =   49
      Top             =   6720
      Width           =   1140
   End
   Begin VB.Frame frmSubtotal 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   0
      TabIndex        =   21
      Tag             =   "cls"
      Top             =   6300
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
      Left            =   5910
      TabIndex        =   0
      Top             =   6330
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
      Left            =   10470
      TabIndex        =   10
      Top             =   6330
      Width           =   915
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
      Left            =   8184
      TabIndex        =   20
      Top             =   6330
      Width           =   1140
   End
   Begin VB.TextBox txtNumFac 
      Height          =   285
      Left            =   8820
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6735
      Width           =   1020
   End
   Begin VB.TextBox txtNumSer 
      Height          =   285
      Left            =   8325
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6720
      Width           =   435
   End
   Begin VB.TextBox txtFindRegistro 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAEDE9&
      ForeColor       =   &H009C3000&
      Height          =   285
      Left            =   2985
      TabIndex        =   18
      Top             =   2625
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F1EC&
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
      Left            =   2985
      TabIndex        =   17
      Top             =   2235
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   3
      Left            =   3945
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4305
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   3900
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   4665
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkCambiar 
      Caption         =   "Cambiar"
      Height          =   240
      Left            =   10095
      TabIndex        =   11
      Top             =   6750
      Width           =   1215
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
      Left            =   7047
      TabIndex        =   8
      Top             =   6330
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
      Left            =   9330
      TabIndex        =   9
      Top             =   6330
      Width           =   1140
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2400
      Left            =   2160
      TabIndex        =   16
      Top             =   3405
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
   Begin MSMask.MaskEdBox txt2 
      Height          =   270
      Left            =   3000
      TabIndex        =   19
      Top             =   3015
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
   Begin OSControlsUser.OSFindItem txtcli 
      Height          =   285
      Left            =   1275
      TabIndex        =   6
      Tag             =   "cls"
      Top             =   1170
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   503
      Enabled         =   0   'False
      Locked          =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4440
      Left            =   90
      TabIndex        =   7
      Top             =   1650
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   7832
      _Version        =   393216
      Enabled         =   0   'False
      AllowUserResizing=   1
   End
   Begin OSControlsUser.OSFindItem Txt_key 
      Height          =   285
      Left            =   4620
      TabIndex        =   44
      Top             =   1215
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   503
      Enabled         =   0   'False
      Locked          =   0   'False
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   225
      Top             =   7035
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame f1 
      Enabled         =   0   'False
      Height          =   1650
      Left            =   90
      TabIndex        =   24
      Top             =   0
      Width           =   11775
      Begin VB.TextBox txtNumSol 
         Height          =   285
         Left            =   1170
         TabIndex        =   1
         Tag             =   "cls"
         Top             =   300
         Width           =   1020
      End
      Begin VB.ComboBox cmdtipo 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOCCliente.frx":000C
         Left            =   8445
         List            =   "frmOCCliente.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2040
         Width           =   2745
      End
      Begin VB.TextBox i_dias 
         Enabled         =   0   'False
         Height          =   300
         Left            =   7080
         TabIndex        =   29
         Top             =   2400
         Width           =   495
      End
      Begin VB.ComboBox i_condi 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOCCliente.frx":0020
         Left            =   7905
         List            =   "frmOCCliente.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2430
         Width           =   2730
      End
      Begin VB.ComboBox i_destino 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOCCliente.frx":0024
         Left            =   240
         List            =   "frmOCCliente.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2370
         Width           =   5490
      End
      Begin VB.ComboBox i_fbg 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOCCliente.frx":0047
         Left            =   5805
         List            =   "frmOCCliente.frx":0051
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   2385
         Width           =   1110
      End
      Begin VB.ComboBox moneda 
         Height          =   315
         ItemData        =   "frmOCCliente.frx":005B
         Left            =   2940
         List            =   "frmOCCliente.frx":0065
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "cls"
         Top             =   750
         Width           =   975
      End
      Begin VB.TextBox txtruc 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6930
         Locked          =   -1  'True
         TabIndex        =   25
         Tag             =   "cls"
         Top             =   1170
         Width           =   1515
      End
      Begin VB.TextBox txtNroDoc 
         Height          =   285
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   4
         Tag             =   "cls"
         Top             =   795
         Width           =   1830
      End
      Begin VB.TextBox txtOferta 
         Height          =   615
         Left            =   6915
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Tag             =   "cls"
         Top             =   480
         Width           =   4770
      End
      Begin OSControlsUser.ctlMaskEdBox txtFecha 
         Height          =   300
         Left            =   1170
         TabIndex        =   2
         Tag             =   "cls"
         Top             =   735
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
      Begin VB.Label lblnroCotizacion 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3645
         TabIndex        =   48
         Tag             =   "cls"
         Top             =   330
         Width           =   1920
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         Caption         =   "Nº Solicitud: "
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
         Left            =   2700
         TabIndex        =   47
         Tag             =   "9999"
         Top             =   360
         Width           =   915
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         Caption         =   "Nº Cotizacion :"
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
         Index           =   12
         Left            =   45
         TabIndex        =   46
         Tag             =   "9999"
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Venta"
         Height          =   255
         Index           =   1
         Left            =   8475
         TabIndex        =   43
         Top             =   2055
         Width           =   2010
      End
      Begin VB.Label lcodart 
         Caption         =   "Dias Cred."
         Height          =   255
         Index           =   11
         Left            =   7110
         TabIndex        =   42
         Tag             =   "9999"
         Top             =   2400
         Width           =   810
      End
      Begin VB.Label lcodart 
         Caption         =   "Condición Venta"
         Height          =   255
         Index           =   10
         Left            =   7935
         TabIndex        =   41
         Tag             =   "9999"
         Top             =   2430
         Width           =   1995
      End
      Begin VB.Label lcodart 
         Caption         =   "   Destino Almacen :"
         Height          =   255
         Index           =   9
         Left            =   150
         TabIndex        =   40
         Tag             =   "9999"
         Top             =   2460
         Width           =   1470
      End
      Begin VB.Label lblven 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5175
         TabIndex        =   39
         Top             =   2010
         Width           =   2955
      End
      Begin VB.Label lcodart 
         Caption         =   "Fact./Bolet."
         Height          =   255
         Index           =   8
         Left            =   5805
         TabIndex        =   38
         Tag             =   "9999"
         Top             =   2430
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Vendedor:"
         Height          =   255
         Index           =   0
         Left            =   5175
         TabIndex        =   37
         Top             =   2025
         Width           =   915
      End
      Begin VB.Label lblcli 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2445
         TabIndex        =   36
         Tag             =   "cls"
         Top             =   1170
         Width           =   4425
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
         Left            =   2235
         TabIndex        =   35
         Tag             =   "9999"
         Top             =   780
         Width           =   720
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
         Left            =   510
         TabIndex        =   34
         Tag             =   "9999"
         Top             =   1185
         Width           =   600
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   0
         Left            =   570
         TabIndex        =   33
         Tag             =   "9999"
         Top             =   780
         Width           =   540
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         Caption         =   "Nº Orden C.: "
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
         Left            =   4020
         TabIndex        =   32
         Tag             =   "9999"
         Top             =   780
         Width           =   990
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
         Left            =   6900
         TabIndex        =   31
         Tag             =   "9999"
         Top             =   270
         Width           =   3900
      End
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
      Left            =   7890
      TabIndex        =   45
      Tag             =   "9999"
      Top             =   6765
      Width           =   330
   End
End
Attribute VB_Name = "frmOCCliente"
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
Dim FlagSolicitud As Integer
Private ColumnNext() As Integer
Dim iCounItemsReser As Integer
Dim Msg As String
Dim ColorCellDefault As Variant

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
        SQL = "UPDATE PEDIDOS SET PED_ESTADO = 'E' WHERE PED_TIPMOV = 400 AND PED_CODCIA='" & LK_CODCIA & "' AND PED_NUMSER='" & txtNumSer.Text & " AND PED_NUMFAC = " & txtNumFac.Text
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
       txtcli.Enabled = False
       cmdIngresar.SetFocus
    End If
    If cmdConsulta.Caption = "&Grabar" Then
       cmdIngresar.Caption = "&Modificar"
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

    iCounItemsReser = 0
    SQ_OPER = 1
    PUB_TIPMOV = 400
    pu_codcia = LK_CODCIA
    PUB_PEDSER = txtNumSer.Text
    PUB_PEDFAC = txtNumFac.Text
    LEER_PED_LLAVE
    If Not ped_llave.EOF Then
         
        pub_cadena = "SELECT * FROM CONTROLL"
        CN.Execute "Begin Transaction", rdExecDirect
        Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
        Msg = "El stock de algunos articulos no completan la cantidad solicitada en esta OC..." & vbCrLf
        iRowsFacar = grd.Rows - 1
        Do While Not ped_llave.EOF
            iCount = iCount + 1
            VerifReserStock iCount
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
                VerifReserStock iCount
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
        MsgBox "Error. Cotizacion no existe", vbCritical, Pub_Titulo
    End If
    If iCounItemsReser <> 0 Then
        MsgBox Msg, vbInformation, Pub_Titulo
    End If
    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical, Pub_Titulo
     con_llave.Close
     CN.Execute "Rollback Transaction", rdExecDirect
End Sub

Private Sub cmdIngresar_Click()
Dim iSecuencia As Long
Dim RowsGrd As Integer
Dim SQL As String
On Error GoTo Handler

    If cmdIngresar.Caption = "&Ingresar" Then
        cmdIngresar.Caption = "&Grabar"
        NumOC
        txtFecha.Text = LK_FECHA_DIA
        f1.Enabled = True
        grd.Enabled = True
        txtcli.Enabled = True
        txtNumSol.SetFocus
        ClearForm Me
        FormatGrid
        stbSubtotales.Panels(1).Text = "0.000"
        stbSubtotales.Panels(3).Text = "0.000"
        stbSubtotales.Panels(4).Text = ""
        stbSubtotales.Panels(5).Text = "0.000"
    Else
        iCounItemsReser = 0
        If FlagSolicitud = 1 Then
            MsgBox "Operación no procede... la Solicitud esta procesada"
            Exit Sub
        ElseIf FlagSolicitud = 3 Then
            MsgBox "Operación no procede... no selecciono una Solicitud"
            Exit Sub
        End If
        
        If ConsisGrl <> 0 Then Exit Sub
         
        pub_cadena = "SELECT * FROM CONTROLL"
        CN.Execute "Begin Transaction", rdExecDirect
        Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
        Msg = "El stock de algunos articulos no completan la cantidad solicitada en esta OC..." & vbCrLf

        RowsGrd = grd.Rows - 1
        For iSecuencia = 1 To RowsGrd
            If Trim(grd.TextMatrix(iSecuencia, 13)) = "X" Then
                VerifReserStock iSecuencia
                ped_llave.AddNew
                AsignaValores iSecuencia, 1
                ped_llave.Update
                SQL = "UPDATE PEDIDOS SET PED_SITUACION = 'P' WHERE PED_NUMFAC = " & Val(txtNumSol.Text) & " AND PED_TIPMOV=300 AND PED_CODCIA='" & LK_CODCIA & "' AND PED_CODART = " & Val(grd.TextMatrix(iSecuencia, 12))
                CN.Execute SQL
            End If
        Next iSecuencia
        
        CN.Execute "Commit Transaction", rdExecDirect
        con_llave.Close
        f1.Enabled = False
        grd.Enabled = False
        txtcli.Enabled = False
        cmdcancelar_Click
        If iCounItemsReser <> 0 Then
            MsgBox Msg, vbInformation, Pub_Titulo
        End If
    End If
    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical, Pub_Titulo
    
     con_llave.Close
     CN.Execute "Rollback Transaction", rdExecDirect
    
End Sub
Private Sub AsignaValores(ByVal fila As Integer, ByVal tipo As Integer)
If tipo = 1 Then
    ped_llave("Ped_CODCIA") = LK_CODCIA
    ped_llave("Ped_FECHA") = txtFecha.Text
    ped_llave("Ped_NUMSER") = txtNumSer.Text
    ped_llave("Ped_NUMFAC") = Val(txtNumFac.Text)
    ped_llave("Ped_NUMSEC") = fila
    ped_llave("Ped_ESTADO") = "N" '
    ped_llave("Ped_TIPMOV") = 400 '
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
End If
    ped_llave("Ped_CANTIDAD") = dDouble(grd.TextMatrix(fila, 2)) * dDouble(grd.TextMatrix(fila, 11))
    ped_llave("Ped_PRECIO") = dDouble(grd.TextMatrix(fila, 4))
    ped_llave("Ped_CODUSU") = LK_CODUSU
    ped_llave("Ped_IGV") = PUB_IMPTO
    ped_llave("Ped_BRUTO") = PUB_IMPORTE
    ped_llave("Ped_CODART") = grd.TextMatrix(fila, 12)
    ped_llave("Ped_UNIDAD") = Trim(grd.TextMatrix(fila, 3))
    ped_llave("Ped_EQUIV") = Val(grd.TextMatrix(fila, 11))
    ped_llave("Ped_CODCLIE") = txtcli.TEXTO
    ped_llave("Ped_HORA") = Time
    ped_llave("Ped_DESCTO") = dDouble(grd.TextMatrix(fila, 8))
    ped_llave("Ped_MONEDA") = PUB_DS
    'ped_llave("Ped_CONTACTO") = Left(txtContacto.Text, 50)
    ped_llave("Ped_NOMCLIE") = Left(lblcli.Caption, 50)
    ped_llave("Ped_RUCCLIE") = Left(txtruc.Text, 15)
    ped_llave("Ped_OFERTA") = txtOferta.Text
    ped_llave("Ped_SUBTOTAL") = dDouble(grd.TextMatrix(fila, 10))
    ped_llave("Ped_NUMPRE") = dDouble(grd.TextMatrix(fila, 15))
    ped_llave("Ped_DESCTO_PRE") = dDouble(grd.TextMatrix(fila, 9))
    ped_llave("PED_NUMDOC") = txtNroDoc.Text
    'ped_llave("Ped_CANATEN") = dDouble(grd.TextMatrix(fila, 16))
    ped_llave("Ped_NUMSER_C") = "0"
    ped_llave("PED_NUMFAC_C") = Val(txtNumSol.Text)
End Sub
Private Sub Form_Load()
    FormatObjects 0
    FormatGrid
'    carga_venta
'    LlenadoCbo cmdtipo, 65
    txtNumSer.Text = 1
    NumOC
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
        txtNroDoc.SetFocus
    End If
End Sub
'Private Sub cmdtipo_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'       txtcli.SetFoco
'    End If
'End Sub

'Private Sub cmdtipo_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 45 Then
'        PUB_TIPREG = 65
'        PUB_CODCIA = LK_CODCIA
'        Load FrmDatArti
'        FrmDatArti.Caption = Trim(Left(cmdtipo.Text, 30)) & " TAB_TIPREG = " & PUB_TIPREG
'        FrmDatArti.Show 1
'        LlenadoCbo cmdtipo, PUB_TIPREG
'        If cmdtipo.ListCount > 0 Then cmdtipo.ListIndex = 0
'        cmdtipo.SetFocus
'        SendKeys "%{UP}"
'        DoEvents
'    End If
'End Sub
'Public Sub carga_venta()
'    SQ_OPER = 2
'    PUB_CODTRA = 2401
'    LEER_SUT_LLAVE
'    i_condi.Clear
'    Do Until SUT_MAYOR.EOF
'        i_condi.AddItem Format(SUT_MAYOR!SUT_SECUENCIA, "00") & ".-" & SUT_MAYOR!sut_descripcion & String(180, " ") & SUT_MAYOR!SUT_SIGNO_CAR & SUT_MAYOR!sut_TIPDOC
'        SUT_MAYOR.MoveNext
'    Loop
'    moneda.Clear
'    If LK_MONEDA = "S" Then
'       moneda.AddItem "S = S/."
'    ElseIf LK_MONEDA = "D" Then
'       moneda.AddItem "D = US$"
'    Else
'       moneda.AddItem "S = S/."
'       moneda.AddItem "D = US$"
'    End If
'    'txtfecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
'End Sub

'===========================================================================
'============Para Presentacion de Datos de Cliente=============
Private Sub txtcli_Cancel()
    txtcli.TEXTO = ""
    lblcli.Caption = ""
    'txtContacto.Text = ""
End Sub

Private Sub txtcli_GetRegistros(ByVal oKeyFind As Variant)
Dim sSql As String
On Error GoTo ErroHandle
    sSql = "SELECT 'Razon Social de la Empredsa'=cli_nombre, 'Codigo'=cli_codcliE, 'RUC'=CLI_RUC_ESPOSO, 'Direccion'=cli_casa_direc FROM CLIENTES WHERE CLI_CP='C' AND CLI_NOMBRE LIKE '" & oKeyFind & "%' AND CLI_CODCIA = '" & LK_CODCIA & "' ORDER BY CLI_NOMBRE"
    txtcli.TypeFind = NameField
    txtcli.SetRecordset = OpenSQLForwardOnly(sSql)
    Exit Sub
ErroHandle:
 MsgBox Err.Description
End Sub

Private Sub txtcli_GotFocus()
    txtcli.ZOrder 0
End Sub

Private Sub txtcli_LostFocus()
'    pub_cadena = "SELECT * FROM DIRCLI WHERE CODCIA=? AND CODCLI=? AND CP=?"
'    Set PSFAR_TRANS = CN.CreateQuery("", pub_cadena)
'    PSFAR_TRANS.rdoParameters(0) = LK_CODCIA
'    PSFAR_TRANS.rdoParameters(1) = Val(txtcli.TEXTO)
'    PSFAR_TRANS.rdoParameters(2) = "C"
'    Set FAR_TRANS = PSFAR_TRANS.OpenResultset(rdOpenKeyset, rdConcurValues)
'    i_destino.Clear
'    Do Until FAR_TRANS.EOF
'      i_destino.AddItem Trim(FAR_TRANS!DIRCOMP) & String(80, " ") & Trim(FAR_TRANS!DIRCLI)
'      FAR_TRANS.MoveNext
'    Loop
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
        'txtContacto.Text = cli_llave("CLI_NOMBRE_ESPOSA")
        grd.SetFocus
    End If
End Sub
'===========================================================================
'============Para Presentacion de Datos de Empleado=============
Private Sub Txt_key_Cancel()
    txt_key.TEXTO = ""
    lblven.Caption = ""
End Sub

Private Sub Txt_key_GetRegistros(ByVal oKeyFind As Variant)
Dim sSql As String
On Error GoTo ErroHandle
    sSql = "SELECT  VEM_NOMBRE AS 'Apellidos y Nombres', vem_codven as CODIGO FROM VEMAEST WHERE VEM_CODCIA = '" & LK_CODCIA & "' ORDER BY VEM_NOMBRE"
    txt_key.TypeFind = NameField
    txt_key.SetRecordset = OpenSQLForwardOnly(sSql)
    Exit Sub
ErroHandle:
 MsgBox Err.Description
End Sub

Private Sub txt_key_GotFocus()
    txt_key.ZOrder 0
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
        .FormatString = "Descripción|Código|Cantidad|Unidad|Precio|División|Familia||||SubTotal"
        .ColWidth(0) = 4500 'descripcion
        .ColWidth(1) = 1200 'codigo
        .ColWidth(2) = 900 'cantidad
        .ColWidth(3) = 1000 'unidad
        .ColWidth(4) = 900 'precio
        .ColWidth(5) = 900 'laboiratorio
        .ColWidth(6) = 900 'marca
        .ColWidth(7) = 0  ' procedencia
        .ColWidth(8) = 0  '% descuento
        .ColWidth(9) = 0 ' descuento
        .ColWidth(10) = 900 'subtotal
        .ColWidth(11) = 0 'equivalencia
        .ColWidth(12) = 0 'cod_art
        .ColWidth(13) = 0 'graba(X) o no( )
        .ColWidth(14) = 0 'secuencia DEL DETALLE
        .ColWidth(15) = 0 'secuencia DEL PRECIO
        .ColWidth(16) = 0 'CANTIDAD ORIGINAL DE SOLICITUD
        
        ColumnNext(1) = 1
        ColumnNext(2) = 1
        ColumnNext(3) = 1
        ColumnNext(4) = -1
        ColumnNext(5) = 5
        ColumnNext(6) = 4
        ColumnNext(7) = 3
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
End Sub
Private Sub CalculaSubTotales1(ByVal RowA As Long, ByVal ColA As Long)
    PUB_SUBTOTAL = dDouble(grd.TextMatrix(RowA, 2)) * dDouble(grd.TextMatrix(RowA, 4))
    grd.TextMatrix(RowA, 10) = Format(PUB_SUBTOTAL, "##0.000")
End Sub
'Private Sub UbicaCombo()
'    'cmbprecios.Visible = True
'    cbmTop = grd.Top + (iRowGrd * 360) '180
'    cbmLeft = grd.Left + 7100
'    cmbprecios.Left = cbmLeft
'    cmbprecios.Top = cbmTop
'End Sub
Private Sub NumOC()

    archi = "SELECT MAX(PED_NUMFAC) AS NUMFAC FROM PEDIDOS WHERE PED_TIPMOV=400 AND Ped_FBG = 'C' AND PED_CODCIA = '" & LK_CODCIA & "'"
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
            txt.Text = grd.Text
            txt.Visible = True
            If Len(grd.Text) > 0 Then Azul txt, txt
            txt.SetFocus
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
                grd.Rows = grd.Rows + 1
            End If
            grd.COL = 1
            grd.Row = grd.Row + 1
        End If
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
    End If
    If grd.COL = 4 Or grd.COL = 2 Then
        KeyAscii = vNumeric(KeyAscii)
    End If
End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        moneda.SetFocus
        RES = SendMessageLong(moneda.hwnd, &H14F, True, 0)
    End If
End Sub
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
            SQ_OPER = 1
            LEER_ART_LLAVE
            LOADUNIDADES
            LOADLABO
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
Dim sText As String
Dim sKey As String
Dim iPos As Integer
    If KeyAscii = 13 Then
        If cbo(Index).ListIndex = -1 Then Exit Sub
        If Len(Trim(cbo(Index).Text)) = 0 Then GoTo SEGUIR
        iPos = InStr(1, cbo(Index).Text, "      ")
        If iPos = 0 Then iPos = Len(cbo(Index).Text)
        sText = Trim(Mid(cbo(Index).Text, 1, iPos))
SEGUIR:
        grd.Text = sText
        If grd.COL = 4 Then
            CalculaSubTotales1 grd.Row, grd.COL
            CalculaTotales1 grd.Row, grd.COL
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
        archi = archi & "FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA INNER JOIN TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB INNER JOIN TABLAS TABLAS_1 ON ARTI.ART_CODCIA = TABLAS_1.TAB_CODCIA AND ARTI.ART_SUBFAM = TABLAS_1.TAB_NUMTAB  "
        archi = archi & "WHERE (TABLAS.TAB_TIPREG = 122) AND (TABLAS_1.TAB_TIPREG = 123) AND (PRECIOS.PRE_FLAG_UNIDAD = 'A') AND ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_CALIDAD = 1 AND ARTI.ART_FLAG_STOCK = 'M' AND ARTI.ART_ALTERNO BETWEEN '" & Trim(txtFindRegistro.Text) & "%' AND  '" & var & "' ORDER BY ARTI.ART_ALTERNO"
    Else
    '    archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND  (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' AND ART_NOMBRE BETWEEN '" & Trim(txtFindRegistro.Text) & "%' AND  '" & var & "' ORDER BY ART_NOMBRE"
        archi = "SELECT ARTI.ART_KEY, ARTI.ART_NOMBRE, ARTI.ART_ALTERNO, ARTICULO.ARM_STOCK, PRECIOS.PRE_EQUIV, TABLAS.TAB_TIPREG, TABLAS_1.TAB_TIPREG AS Expr1, PRECIOS.PRE_FLAG_UNIDAD, ARTI.ART_CODCIA, dbo.TABLAS.TAB_NOMLARGO AS Labo, TABLAS_1.TAB_NOMLARGO AS Marca "
        archi = archi & "FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA INNER JOIN TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB INNER JOIN TABLAS TABLAS_1 ON ARTI.ART_CODCIA = TABLAS_1.TAB_CODCIA AND ARTI.ART_SUBFAM = TABLAS_1.TAB_NUMTAB  "
        archi = archi & "WHERE (TABLAS.TAB_TIPREG = 122) AND (TABLAS_1.TAB_TIPREG = 123) AND (PRECIOS.PRE_FLAG_UNIDAD = 'A') AND ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_CALIDAD = 1 AND ARTI.ART_FLAG_STOCK = 'M' AND ARTI.ART_NOMBRE BETWEEN '" & Trim(txtFindRegistro.Text) & "%' AND  '" & var & "' ORDER BY ARTI.ART_NOMBRE"
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
        loListItem.ListSubItems.Add Key:="Stock" & CStr(J), Text:=Trim(X("ARM_STOCK"))
        loListItem.ListSubItems.Add Key:="key" & CStr(J), Text:=Trim(X("ART_KEY"))
        loListItem.ListSubItems.Add Key:="Labo" & CStr(J), Text:=Trim(X("Labo"))
        'loListItem.ListSubItems.Add Key:="Proc" & CStr(J), Text:=Trim(X("Proce"))
        loListItem.ListSubItems.Add Key:="Marca" & CStr(J), Text:=Trim(X("Marca"))
        
        X.MoveNext
    Loop
    If i > 0 Then liIndexRegAct = 1
    FlagLvw = i
End Sub
Private Sub ColHeaders()
On Error Resume Next
    lvw.ColumnHeaders.Add , "Descripcion", "Descripcion", 4500, lvwColumnLeft
    lvw.ColumnHeaders.Add , "Codigo", "Codigo", 1000, lvwColumnLeft
    lvw.ColumnHeaders.Add , "Stock", "Stock", 1000, lvwColumnRight
    lvw.ColumnHeaders.Add , "Key", "Key", 0, lvwColumnRight
    lvw.ColumnHeaders.Add , "Laboratorio", "Laboratorio", 2000, lvwColumnLeft
    'lvw.ColumnHeaders.Add , "Procedencia", "Procedencia", 1500, lvwColumnRight
    lvw.ColumnHeaders.Add , "Marca", "Marca", 2000, lvwColumnRight
End Sub

Private Sub ValRowGrd(ByVal lRow As Integer)
Dim iErrors As Integer
    
    If Trim(grd.TextMatrix(lRow, 0)) = "" Or Trim(grd.TextMatrix(lRow, 1)) = "" Or Val(grd.TextMatrix(lRow, 10)) = 0 Then
        iErrors = iErrors + 1
    End If
    If iErrors = 0 Then
        grd.TextMatrix(lRow, 13) = "X"
    Else
        grd.TextMatrix(lRow, 13) = " "
    End If
End Sub

Private Sub txtNroCot_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtOferta.SetFocus
    End If
End Sub

Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtOferta.SetFocus
    End If
End Sub

Private Sub txtnumfac_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LoadOC
    Else
        KeyAscii = vInteger(KeyAscii)
    End If
End Sub

Private Sub txtNumSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LoadCotizacion
        On Error Resume Next
        txtFecha.SetFocus
        If FlagSolicitud = 3 Then
            MsgBox "Debe ingresar un solicitud valida...", vbInformation, Pub_Titulo
        End If
    End If
End Sub

Private Sub txtOferta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtcli.SetFocus
    End If
End Sub
Private Function ConsisGrl() As Integer
    If Consis1(txtcli.TEXTO) = 1 Then
        MsgBox "Cliente no existe....", 46, Pub_Titulo
        ConsisGrl = 1
        txtcli.SetFoco
        Exit Function
    End If
    If consis2(txtNroDoc.Text) = 1 Then
        MsgBox "Falta Ingresar nro de Orden de Compra....", 46, Pub_Titulo
        ConsisGrl = 1
        txtNroDoc.SetFocus
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
Dim AcuSolP  As Integer
Dim AcuSol  As Integer

    FlagSolicitud = 0
    SQ_OPER = 1
    PUB_TIPMOV = 300
    pu_codcia = LK_CODCIA
    PUB_PEDSER = 0
    PUB_PEDFAC = Val(txtNumSol.Text)
    LEER_PED_LLAVE
    If Not ped_llave.EOF Then
        If Trim(ped_llave("PED_ESTADO")) = "E" Then
            FlagSolicitud = 1
            MsgBox "DOCUMENTO ESTA EXTORNADO", vbInformation, Pub_Titulo
            cmdcancelar_Click
            Exit Sub
        End If
        FormatGrid
        
        Do While Not ped_llave.EOF
            If ped_llave("PED_SITUACION") = "P" Then
                AcuSolP = AcuSolP + 1
            End If
            If ped_llave("PED_SITUACION") = " " Then
                AcuSol = AcuSol + 1
            End If
            If AcuSol = 1 Then
                 txtcli.Enabled = True
                 f1.Enabled = True
                 grd.Enabled = True
                 chkCambiar.Value = 0
                 txtNumFac.Locked = True
                 txtFecha.Text = ped_llave("Ped_FECHA")
                 txtcli.TEXTO = ped_llave("Ped_CODCLIE")
                 Call txtcli_ShowData(txtcli.TEXTO)  'Load Cliente
                 
                 lblcli.Caption = ped_llave("Ped_NOMCLIE")
                ' txtContacto.Text = ped_llave("Ped_CONTACTO")
                 txtruc.Text = ped_llave("Ped_RUCCLIE")
                 txtOferta.Text = ped_llave("Ped_OFERTA")
                 lblnroCotizacion.Caption = ped_llave("PED_NUMDOC")
                 'txtNroDoc.Text = ped_llave("PED_NUMDOC")
                 
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
                 txtcli.SetFocus
            End If
            If ped_llave("PED_SITUACION") = " " Then
                iSec = iSec + 1
                grd.TextMatrix(iSec, 2) = Format(Nulo_Valor0(ped_llave("Ped_CANTIDAD")) / Nulo_Valor0(ped_llave("Ped_EQUIV")), "0.000")
                grd.TextMatrix(iSec, 3) = ped_llave("Ped_UNIDAD")
                grd.TextMatrix(iSec, 4) = Format(ped_llave("Ped_PRECIO"), "0.000")
                grd.TextMatrix(iSec, 8) = ped_llave("Ped_DESCTO")
                grd.TextMatrix(iSec, 10) = ped_llave("Ped_SUBTOTAL")
                grd.TextMatrix(iSec, 11) = ped_llave("Ped_EQUIV")
                grd.TextMatrix(iSec, 12) = ped_llave("Ped_CODART")
                grd.TextMatrix(iSec, 13) = "X"
                grd.TextMatrix(iSec, 14) = ped_llave("Ped_NUMSEC")
                grd.TextMatrix(iSec, 15) = ped_llave("Ped_NUMPRE")
                grd.TextMatrix(iSec, 16) = ped_llave("Ped_CANTIDAD")
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
                    'LOADPROCE
                Else
                    MsgBox "Error grave en Arti...", 46, Pub_Titulo
                End If
                CalculaTotales1 iSec, 2
                grd.Rows = grd.Rows + 1
            End If
            ped_llave.MoveNext
        Loop
        If AcuSol = 0 Then
            FlagSolicitud = 2
            MsgBox "SOLICITUD YA SE ATENDIO TOTALMENTE", vbInformation, Pub_Titulo
            cmdcancelar_Click
            Exit Sub
        End If
        If AcuSolP <> 0 Then
            FlagSolicitud = 2
            MsgBox "SOLICITUD SE ATENDIO PARCIALMENTE", vbInformation, Pub_Titulo
            'CmdCancelar_Click
            Exit Sub
        End If
    Else
        FlagSolicitud = 3
        MsgBox "Solicitud de Cotización no existe", vbInformation, Pub_Titulo
        cmdcancelar_Click
    End If
End Sub
Private Sub LoadOC()
Dim iSec As Integer
Dim CountItemP As Integer
Dim CountItemA As Integer

    FlagSolicitud = 0
    SQ_OPER = 1
    PUB_TIPMOV = 400
    pu_codcia = LK_CODCIA
    PUB_PEDSER = txtNumSer.Text
    PUB_PEDFAC = txtNumFac.Text
    LEER_PED_LLAVE
    If Not ped_llave.EOF Then
        
        If Trim(ped_llave("PED_ESTADO")) = "E" Then
            FlagSolicitud = 1
            MsgBox "DOCUMENTO ESTA EXTORNADO", vbInformation, Pub_Titulo
        End If
        
        txtcli.Enabled = True
        f1.Enabled = True
        grd.Enabled = True
        chkCambiar.Value = 0
        txtNumFac.Locked = True
        
        txtNumSol.Text = ped_llave("Ped_NUMFAC_C")
        txtFecha.Text = ped_llave("Ped_FECHA")
        txtcli.TEXTO = ped_llave("Ped_CODCLIE")
        Call txtcli_ShowData(txtcli.TEXTO)  'Load Cliente
        lblcli.Caption = ped_llave("Ped_NOMCLIE")
       ' txtContacto.Text = ped_llave("Ped_CONTACTO")
        txtruc.Text = ped_llave("Ped_RUCCLIE")
        txtOferta.Text = ped_llave("Ped_OFERTA")
        lblnroCotizacion.Caption = ped_llave("PED_NUMDOC")
        'txtNroDoc.Text = ped_llave("PED_NUMDOC")
        
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
            If ped_llave("Ped_SITUACION") = "A" Then
                CountItemA = CountItemA + 1
            ElseIf ped_llave("Ped_SITUACION") = "P" Then
                CountItemP = CountItemP + 1
            ElseIf ped_llave("Ped_SITUACION") = " " Then
                
            End If
            grd.TextMatrix(iSec, 2) = Format(Nulo_Valor0(ped_llave("Ped_CANTIDAD")) / Nulo_Valor0(ped_llave("Ped_EQUIV")), "0.000")
            grd.TextMatrix(iSec, 3) = ped_llave("Ped_UNIDAD")
            grd.TextMatrix(iSec, 4) = Format(ped_llave("Ped_PRECIO"), "0.000")
            grd.TextMatrix(iSec, 8) = ped_llave("Ped_DESCTO")
            grd.TextMatrix(iSec, 10) = ped_llave("Ped_SUBTOTAL")
            grd.TextMatrix(iSec, 11) = ped_llave("Ped_EQUIV")
            grd.TextMatrix(iSec, 12) = ped_llave("Ped_CODART")
            grd.TextMatrix(iSec, 13) = "X"
            grd.TextMatrix(iSec, 14) = ped_llave("Ped_NUMSEC")
            grd.TextMatrix(iSec, 15) = ped_llave("Ped_NUMPRE")
            grd.TextMatrix(iSec, 16) = ped_llave("Ped_CANTIDAD")
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
                'LOADPROCE
            Else
                MsgBox "Error grave en Arti...", 46, Pub_Titulo
            End If
            ped_llave.MoveNext
        Loop
        txtcli.SetFocus
        
        If CountItemP > 0 Then
            FlagSolicitud = 2
            MsgBox "Orden esta atendida Parcialmente", vbInformation, Pub_Titulo
            Exit Sub
        End If
        If CountItemA = iSec Then
            FlagSolicitud = 3
            MsgBox "Orden esta atendida Totalmente", vbInformation, Pub_Titulo
            Exit Sub
        End If
        cmdIngresar.Enabled = False
        cmdAnular.Enabled = True
        cmdConsulta.Enabled = True
        cmdConsulta.Caption = "Grabar"
    Else
        MsgBox "ORDEN DE COMPRA no existe....", vbInformation, Pub_Titulo
        cmdcancelar_Click
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
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_CODART = Val(grd.TextMatrix(grd.Row, 12))
    PUB_SECUEN = Val(IIf(Trim(grd.TextMatrix(grd.Row, 15)) = "", "-1", grd.TextMatrix(grd.Row, 15)))
    LEER_PRE_LLAVE
    If Not pre_llave.EOF Then
        cbo(4).Clear
        grd.TextMatrix(grd.Row, 11) = pre_llave("PRE_EQUIV")
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
'            If pre_llave("PRE_PRE5") <> 0 Then
'                cbo(4).AddItem pre_llave("PRE_PRE5")
'            End If
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
'            If pre_llave("PRE_PRE55") <> 0 Then
'                cbo(4).AddItem pre_llave("PRE_PRE55")
'            End If
        End If
    Else
        MsgBox "Error grave en precios", 46, Pub_Titulo
    End If
End Sub

Private Sub VerifReserStock(ByVal lRow As Integer)
Dim CodArt As Long
Dim CanStock As Double
Dim CanAcu As Double
Dim CanOC As Double
Dim DIF As Double

    CanOC = dDouble(grd.TextMatrix(lRow, 2)) * dDouble(grd.TextMatrix(lRow, 11))
    CodArt = grd.TextMatrix(lRow, 12)
    archi = "SELECT SUM(PED_CANTIDAD) as SumCant FROM PEDIDOS WHERE PED_SITUACION=' ' AND PED_TIPMOV=400 AND PED_CODART = " & CodArt & " AND PED_NUMSER <> '" & txtNumSer.Text & "' AND PED_NUMFAC <> " & Val(txtNumFac.Text) & " AND PED_CODCIA= '" & LK_CODCIA & "'"
    Set PSX = CN.CreateQuery("", archi)
    Set X = PSX.OpenResultset(rdOpenKeyset)
    X.Requery
    If X.EOF Then
        CanAcu = 0
    Else
        CanAcu = Nulo_Valor0(X("SumCant"))
    End If
    SQ_OPER = 1
    PUB_CODART = CodArt
    pu_codcia = LK_CODCIA
    LEER_ARM_LLAVE
    If arm_llave.EOF Then
        MsgBox "Error grave en Articulo", vbCritical, Pub_Titulo
        Exit Sub
    Else
        CanStock = arm_llave("arm_stock")
    End If
    DIF = (CanStock - CanAcu) - CanOC
    If (CanStock - CanAcu) < CanOC Then
        Msg = Msg & Trim(grd.TextMatrix(lRow, 0)) & " ==> " & DIF & vbCrLf
        iCounItemsReser = iCounItemsReser + 1
    End If
End Sub


Private Sub CmdImprimir_Click()
On Error GoTo Handler
    Reportes.Connect = PUB_ODBC
    Reportes.ReportFileName = PUB_RUTA_OTRO & "OCCLI.RPT"

    Reportes.WindowTitle = "Orden de Compra de Cliente"
    Reportes.Destination = crptToWindow
    Reportes.SelectionFormula = " {PEDIDOS.PED_NUMFAC} = " & txtNumFac.Text & " AND {PEDIDOS.PED_NUMSER} = '" & txtNumSer.Text & "' AND {PEDIDOS.PED_CODCIA} = '" & LK_CODCIA & "'"
    Reportes.Action = 1
    Exit Sub
Handler:
    MsgBox Err.Description
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

