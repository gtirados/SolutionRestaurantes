VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00F5F1EC&
   Caption         =   "Solution - Administración "
   ClientHeight    =   6510
   ClientLeft      =   -210
   ClientTop       =   735
   ClientWidth     =   9480
   Icon            =   "MDIGeneral.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "MDIGeneral.frx":0442
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1920
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":39E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":3E34
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":4150
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":45A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":49F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":4E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":52A0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   240
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport TablasC 
      Left            =   240
      Top             =   1875
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\Archivos de programa\PROYECTO WINDOWS NT\tablas.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6240
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "MAYÚS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NÚM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "11:35 p.m."
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   6174
            MinWidth        =   6174
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   2831
            MinWidth        =   2822
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   741
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FRMDEF"
            Object.ToolTipText     =   "Definición de Trannsacciones"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tablas"
            Object.ToolTipText     =   "Mant. de Tablas"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Password"
            Object.ToolTipText     =   "Password"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ALTERNO"
            Object.ToolTipText     =   "Parametros de Compañia"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Excel"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CIA"
            ImageIndex      =   7
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   16
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu15 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu16 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin VB.CommandButton TXTCIA 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   5445
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   3600
      End
      Begin VB.TextBox txtderechos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Left            =   5430
         TabIndex        =   3
         Text            =   "Cargando Derechos de Usuario . . ."
         Top             =   0
         Visible         =   0   'False
         Width           =   2775
      End
      Begin MSComctlLib.ProgressBar ProBar 
         Height          =   210
         Left            =   5445
         TabIndex        =   2
         Top             =   195
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Menu menuTit1 
      Caption         =   "&Mantenimiento"
      Enabled         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Definicion de Platos,Combos e Insumos"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^A
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Mantenimiento de Articulos"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^I
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Datos de Clientes y Proveedores"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Datos de Mozos"
         Enabled         =   0   'False
         Index           =   3
         Shortcut        =   ^B
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Definicion de Bancos"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Tablas del Sistema"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "C&ompañia"
         Enabled         =   0   'False
         Index           =   6
         Shortcut        =   ^F
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Empresa &General"
         Enabled         =   0   'False
         Index           =   7
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Tablas"
         Enabled         =   0   'False
         Index           =   8
         Shortcut        =   ^T
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Definición de Transacciones"
         Enabled         =   0   'False
         Index           =   9
         Shortcut        =   ^D
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Calendario de Sistema /T. Cambio"
         Enabled         =   0   'False
         Index           =   10
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Mantenimiento de Zonas"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Asignacion de Mesas"
         Index           =   12
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Actualizacion de Mozos"
         Enabled         =   0   'False
         Index           =   13
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Cam&biar Password"
         Enabled         =   0   'False
         Index           =   14
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Cambiar &Usuario"
         Enabled         =   0   'False
         Index           =   15
         Shortcut        =   ^U
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Reiniciar Aplicación"
         Enabled         =   0   'False
         Index           =   16
         Shortcut        =   ^R
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Salir del Sistema"
         Enabled         =   0   'False
         Index           =   17
         Shortcut        =   +^{F9}
      End
   End
   Begin VB.Menu menuTit2 
      Caption         =   "&Tiempo Real"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu SubmenuTit2 
         Caption         =   "&Operaciones"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^G
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "&Consulta de Comandas"
         Index           =   1
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "&Punto de Venta Salon"
         Index           =   2
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "Punto de Venta Mostrador"
         Index           =   4
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "Venta de Entradas"
         Index           =   5
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "&Tipo de Cambio y Costeo"
         Index           =   6
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "Envio Resumen Diario Sunat"
         Index           =   7
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "&Cierre de Operaciones"
         Enabled         =   0   'False
         Index           =   8
      End
   End
   Begin VB.Menu menutit3 
      Caption         =   "&Contabilidad General"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu submenutit3 
         Caption         =   "Platos por Despachar en Lista"
         Index           =   0
      End
      Begin VB.Menu submenutit3 
         Caption         =   "Platos por Despachar por Mesas"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu submenutit3 
         Caption         =   "Platos despachados en Lista"
         Index           =   2
      End
      Begin VB.Menu submenutit3 
         Caption         =   "Platos despachado por Mesa"
         Index           =   3
      End
      Begin VB.Menu submenutit3 
         Caption         =   "."
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu submenutit3 
         Caption         =   "."
         Index           =   5
      End
      Begin VB.Menu submenutit3 
         Caption         =   "."
         Index           =   6
      End
      Begin VB.Menu submenutit3 
         Caption         =   "."
         Index           =   7
      End
      Begin VB.Menu submenutit3 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu submenutit3 
         Caption         =   "."
         Index           =   9
      End
      Begin VB.Menu submenutit3 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu submenutit3 
         Caption         =   "."
         Enabled         =   0   'False
         Index           =   11
      End
      Begin VB.Menu submenutit3 
         Caption         =   "."
         Index           =   12
      End
      Begin VB.Menu submenutit3 
         Caption         =   "."
         Enabled         =   0   'False
         Index           =   13
      End
      Begin VB.Menu submenutit3 
         Caption         =   "."
         Enabled         =   0   'False
         Index           =   14
      End
      Begin VB.Menu submenutit3 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu submenutit3 
         Caption         =   "Importar Datos"
         Index           =   16
      End
   End
   Begin VB.Menu menutit4 
      Caption         =   "&Reportes"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Ventas"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
         Begin VB.Menu menuVent 
            Caption         =   " "
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Compras"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
         Begin VB.Menu Menucomp 
            Caption         =   " "
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Principales"
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
         Begin VB.Menu menuAlm 
            Caption         =   " "
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Inventarios"
         Enabled         =   0   'False
         Index           =   6
         Visible         =   0   'False
         Begin VB.Menu menuMoli 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "C&ontabilidad"
         Enabled         =   0   'False
         Index           =   8
         Visible         =   0   'False
         Begin VB.Menu MenuContab 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Disponible 1 "
         Index           =   10
         Visible         =   0   'False
         Begin VB.Menu menudis1 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Disponible 2"
         Index           =   12
         Visible         =   0   'False
         Begin VB.Menu menudis2 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Disponible 3"
         Index           =   14
         Visible         =   0   'False
         Begin VB.Menu menudis3 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   15
      End
   End
   Begin VB.Menu menutit5 
      Caption         =   "&Utilidades"
      Enabled         =   0   'False
      Begin VB.Menu submenutit5 
         Caption         =   "&Backups            "
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu submenutit5 
         Caption         =   "&Mostrar ultimo Reporte"
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu submenutit5 
         Caption         =   "&Configuración de Reporte"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu submenutit5 
         Caption         =   "&Grupo de Trabajo"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu submenutit5 
         Caption         =   "&Reporte de Transacciones"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu submenutit5 
         Caption         =   "&Depuracion de extornos"
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu submenutit5 
         Caption         =   "&Programación de Tareas"
         Enabled         =   0   'False
         Index           =   6
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menutit6 
      Caption         =   "&Consultas"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Kardex de Articulos"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Cta.Ctes. Clientes y Proveedores"
         Index           =   1
         Shortcut        =   ^E
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Cta.Ctes. de &Bancos"
         Index           =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Operaciones"
         Index           =   4
         Shortcut        =   ^K
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Consulta de Impresion de Cuentas"
         Index           =   5
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Consulta de Comandas"
         Index           =   6
         Shortcut        =   ^N
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Consulta de Anulaciones"
         Index           =   7
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Consulta de Comensales"
         Enabled         =   0   'False
         Index           =   8
         Shortcut        =   ^P
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Consulta de Ventas por Usuario "
         Index           =   9
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Consulta de Descuentos Otorgados"
         Index           =   10
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Exporta Registro de Ventas al PLE"
         Index           =   11
      End
   End
   Begin VB.Menu menutit7 
      Caption         =   "&Eventos y Banquetes"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu submenutit7 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu submenutit7 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu submenutit7 
         Caption         =   "&Gestion de Contratos"
         Index           =   2
      End
   End
   Begin VB.Menu menuAyuda 
      Caption         =   "&Ayuda"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu menuAyudaSistema 
         Caption         =   "&Acerca de. . .."
      End
      Begin VB.Menu menuRepoAyuda 
         Caption         =   "&Pantalla Activa . . . "
      End
   End
   Begin VB.Menu Menucompania 
      Caption         =   "compania"
      Visible         =   0   'False
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   10
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   11
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   12
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   13
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   14
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   15
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   16
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   17
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   18
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   19
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   20
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   21
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   22
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   23
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   24
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   25
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   26
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   27
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   28
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   29
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   30
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   31
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   32
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   33
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   34
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   35
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   36
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   37
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   38
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   39
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   40
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WF_REINICIAR As Integer
Dim Wfile As String
Dim WFORM As String

Public Sub LLENA_ACCESO_OTROS(WSOTROS As String)
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
WPosV = Len(WSOTROS)
cade = Trim(WSOTROS)
cuenta = 0
wPosF = 1
a = 0
For i = 1 To Len(cade)
 If Mid(cade, i, 1) = "." Then
   a = a + 1
 End If
Next i
lk_OTROS_Count = a
ReDim lk_OTROS(a)
Do Until cuenta = a
   cuenta = cuenta + 1
   wPosF = InStr(wPosF, cade, ".", 1) + 1
   WNUM = Mid(cade, wPosF, 2)
   If Right(WNUM, 1) = "." Then
     WNUM = Left(WNUM, 2)
     wPosF = wPosF - 1
   End If
   lk_OTROS(cuenta) = Format(CStr(WNUM), "00")
Loop

End Sub

Private Sub MDIForm_Load()
Dim WUSU_CIAS As String
Dim WUSU_OTROS As String
Dim h, a, b, C, d, E As Integer
Dim wBAN As Integer
Dim WNUM, wPosF, WPosV, cuenta, cade
Dim Wflag As String * 1
Dim f, g, fg As Integer
Dim wserie As Currency
Dim WFECHA
Dim WS
'On Error GoTo SALE
Wflag = ""
TXTCIA.Enabled = False
PS_GEN(0) = 0
gen.Requery
Pub_Titulo = Trim(gen!GEN_NOMBRE)
PUB_RUTA_REPORTE = Trim(Nulo_Valors(gen!gen_ruta_reportes))
PUB_RUTA_OTRO = Trim(Nulo_Valors(gen!gen_ruta_otros))
LK_IGV = Nulo_Valor0(gen!GEN_IGV)
LK_TIPO_CAMBIO = Nulo_Valor0(gen!gen_tipo_cambio)
TXTCIA.Visible = False
DoEvents
Splash.Hide
DoEvents
MDIForm1.Show
DoEvents
NOTPermisos
PASSFORM.Show 1
DoEvents
Screen.MousePointer = 11
ProBar.Value = 0
ProBar.Min = 0
ProBar.Max = 15
txtderechos.Visible = True
DoEvents
ProBar.Visible = True
DoEvents
If LK_CODCIA = "" Or LK_CODUSU = "" Then
  MsgBox "Existe un Problema con su Acceso , consulte al Administrador ", 48, Pub_Titulo
  GoTo fin_de_pro
End If

StatusBar1.Enabled = False
PS_PAR(0) = " "
par.Requery
If par.EOF Then
  GoTo PASAR
End If
'If LK_CODUSU <> "ADMIN" Then
' MDIForm1.submenutit5(2).Visible = False
 'MDIForm1.submenutit5(3).Visible = False
'Else
' MDIForm1.submenutit5(2).Visible = True
' MDIForm1.submenutit5(3).Visible = True
'End If
usu.Requery
ProBar.Value = ProBar.Value + 1
Do Until usu.EOF
If Trim(usu!usu_key) = LK_CODUSU Then
   WUSU_CIAS = Trim(Nulo_Valors(usu!USU_CIAS))
   WUSU_OTROS = Trim(Nulo_Valors(usu!USU_OTROS))
   LK_PRECIO = Trim(Nulo_Valors(usu!USU_PRECIO))
   
   Exit Do
End If
usu.MoveNext
Loop
If Len(WUSU_CIAS) = 0 Then
  MsgBox "Usted NO Tiene accesos a ninguna compañia , consulte al Administrador ", 48, Pub_Titulo
  GoTo fin_de_pro
End If

If WUSU_OTROS = "" Then
 lk_OTROS_Count = 0
 ReDim lk_OTROS(0)
Else
 LLENA_ACCESO_OTROS WUSU_OTROS
End If
ProBar.Value = ProBar.Value + 1
WNUM = 0
wPosF = 0
WPosV = 0
cuenta = 0
WPosV = Len(WUSU_CIAS)
cade = Trim(WUSU_CIAS)
cuenta = 0
wPosF = 1
a = 0
ProBar.Value = ProBar.Value + 1
For fila = 1 To Len(cade)
    If Mid(cade, fila, 1) = "." Then
        a = a + 1
    End If
Next fila
ProBar.Value = ProBar.Value + 1
wBAN = -1
h = 0
Do Until cuenta = a
    h = h + 1
    cuenta = cuenta + 1
    DoEvents
    wPosF = InStr(wPosF, cade, ".", 1) + 1
    DoEvents
    WNUM = Mid(cade, wPosF, 2)
    If Right(WNUM, 1) = "." Then
        WNUM = Left(WNUM, 2)
        wPosF = wPosF - 1
    End If
    SQ_OPER = 1
    PUB_CODCIA = Format(WNUM, "00")
    LEER_PAR_LLAVE
    If Not par_llave.EOF Then
       MDIForm1.menuCia(cuenta - 1).Caption = PUB_CODCIA & "-" & Trim(par_llave!PAR_NOMBRE)
       '*******************AGREGADO*****
       Toolbar1.Buttons("CIA").ButtonMenus(cuenta).text = PUB_CODCIA & "-" & Trim(par_llave!PAR_NOMBRE)
       '*******************AGREGADO*****
       MDIForm1.menuCia(cuenta - 1).Visible = True
       If LK_CODCIA = par_llave!PAR_CODCIA Then wBAN = h - 1
    Else
       If a = 1 Then ' Cuando la unica compañia de acceso no exite
         MsgBox "Su Compañia de Acceso NO Existe , consulte al Administrador, Procede a Salir del Sistema", 48, Pub_Titulo
         GoTo fin_de_pro
       End If
       MDIForm1.menuCia(cuenta - 1).Caption = "* " & PUB_CODCIA & " , NO Procede"
       MDIForm1.menuCia(cuenta - 1).Visible = True
    End If
Loop
ProBar.Value = ProBar.Value + 1
If wBAN = -1 Then
  For a = 0 To h - 1
    If Left(MDIForm1.menuCia(a).Caption, 1) <> "*" Then
        wBAN = a
        Exit For
    End If
  Next a
  If wBAN = -1 Then
    MsgBox "Su(s) Compañias de Acceso NO Existen , consulte al Administrador, Procede a Salir del Sistema", 48, Pub_Titulo
    GoTo fin_de_pro
  End If
End If

For a = h To 40
  MDIForm1.menuCia(a).Visible = False
Next a
MDIForm1.menuCia(wBAN).Checked = True
'asegura la compañia por defecto
ACTUALIZA_CIA Left(MDIForm1.menuCia(wBAN).Caption, 2)

ProBar.Value = ProBar.Value + 1
PASAR:
PS_GEN(0) = 0
gen.Requery
ProBar.Value = ProBar.Value + 1
Pub_Titulo = Trim(gen!GEN_NOMBRE)
PUB_RUTA_REPORTE = Trim(Nulo_Valors(gen!gen_ruta_reportes))
PUB_RUTA_OTRO = Trim(Nulo_Valors(gen!gen_ruta_otros))
MDIForm1.StatusBar1.Panels(5).text = Trim(gen!GEN_NOMBRE)
MDIForm1.StatusBar1.Panels(6).text = LK_CODUSU
MDIForm1.StatusBar1.Panels(4).text = "Fecha de la Empresa :  " & Format(LK_FECHA_DIA, "dd, mmmm yyyy")
MDIForm1.StatusBar1.Panels(7).text = MDIForm1.menuCia(wBAN).Caption
ProBar.Value = ProBar.Value + 1
MDIForm1.TXTCIA.Caption = MDIForm1.menuCia(wBAN).Caption
'*******************AGREGADO*****
Toolbar1.Buttons("CIA").ToolTipText = Toolbar1.Buttons("CIA").ButtonMenus(wBAN + 1).text
'*******************AGREGADO*****
CONST_SERVER = Trim(Nulo_Valors(gen!GEN_SERVER))
CONST_UID = Trim(Nulo_Valors(gen!GEN_UID))
CONST_PWD = Trim(Nulo_Valors(gen!GEN_PWD))
ProBar.Value = ProBar.Value + 1
If Len(CONST_SERVER) = 0 Or Len(CONST_UID) = 0 Then
  MsgBox "ADVERTENCIA ..!! DEBES DEFINIR EL NOMBRE DEL SERVIDOR , UID , PWD", 48, Pub_Titulo
End If
PUB_ODBC = "ODBC;DRIVER={SQL Server};UID=" & CONST_UID & ";PWD=" & CONST_PWD & ";APP=Data From Designer;WSID=" & CONST_SERVER & ";LANGUAGE=us_english;DATABASE=BDFERRETEROS"

' * Permisos para los Menus *
Screen.MousePointer = 11
Permisos
ProBar.Value = ProBar.Value + 1
Screen.MousePointer = 11
' * Carga menu de Reporte *
SQ_OPER = 2
PUB_INICIO = 8100
LEER_TRA_LLAVE
ProBar.Value = ProBar.Value + 1
a = 0
b = 0
C = 0
d = 0
E = 0
f = 0
g = 0
fg = 0
Do Until tra_menu.EOF
 If tra_menu!TRA_FLAG_EXT <> "A" Then
   GoTo OTRO_MENU
 End If
 If tra_menu(0) >= 8100 And tra_menu(0) < 8299 Then
     If a > 0 Then
        Load MDIForm1.menuVent(a)
     End If
    MDIForm1.menuVent(a).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    a = a + 1
 ElseIf tra_menu(0) >= 8300 And tra_menu(0) < 8499 Then
    If b > 0 Then
        Load MDIForm1.Menucomp(b)
    End If
    MDIForm1.Menucomp(b).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    b = b + 1
 ElseIf tra_menu(0) >= 8500 And tra_menu(0) < 8699 Then
    If C > 0 Then
        Load MDIForm1.menuAlm(C)
    End If
    MDIForm1.menuAlm(C).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    C = C + 1
 ElseIf tra_menu(0) >= 8700 And tra_menu(0) < 8899 Then
    If d > 0 Then
        Load MDIForm1.menuMoli(d)
    End If
    MDIForm1.menuMoli(d).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    d = d + 1
 ElseIf tra_menu(0) >= 8900 And tra_menu(0) < 9099 Then
    If E > 0 Then
        Load MDIForm1.MenuContab(E)
    End If
    MDIForm1.MenuContab(E).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    E = E + 1
 ElseIf tra_menu(0) >= 9100 And tra_menu(0) < 9299 Then
    If f > 0 Then
        Load MDIForm1.menudis1(f)
    End If
    MDIForm1.menudis1(f).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    f = f + 1
 ElseIf tra_menu(0) >= 9300 And tra_menu(0) < 9499 Then
    If g > 0 Then
        Load MDIForm1.menudis2(g)
    End If
    MDIForm1.menudis2(g).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    g = g + 1
 ElseIf tra_menu(0) >= 9500 And tra_menu(0) < 9699 Then
    If fg > 0 Then
        Load MDIForm1.menudis3(fg)
    End If
    MDIForm1.menudis3(fg).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    fg = fg + 1
 End If

OTRO_MENU:
 tra_menu.MoveNext
Loop

SQ_OPER = 2
PUB_TIPREG = 99
PUB_CODCIA = "00"
LEER_TAB_LLAVE
If tab_mayor.EOF Then
 MsgBox " No existe TIPREG 99 en tablas crearlas con titulos del Menu correspondiente..", 48, Pub_Titulo
End If
fila = 0
SQ_OPER = 0
Do Until tab_mayor.EOF
 MDIForm1.SubmenuTit4(fila).Caption = "&" + Trim(tab_mayor!TAB_NOMLARGO)
 If Val(tab_mayor!TAB_NOMCORTO) = 1 Then
  MDIForm1.SubmenuTit4(fila).Visible = True
  MDIForm1.SubmenuTit4(fila + 1).Visible = True
  SQ_OPER = 9
 End If
 fila = fila + 2
 tab_mayor.MoveNext
Loop
If SQ_OPER = 9 Then
 MDIForm1.SubmenuTit4(15).Visible = False
End If

If LK_CODUSU <> "ADMIN" Then
 MDIForm1.SubmenuTit1(8).Caption = "&Usuarios"
 MDIForm1.Toolbar1.Buttons.Item(2).ToolTipText = "Mant. Usuarios"
Else
 MDIForm1.SubmenuTit1(8).Caption = "&Tablas"
 MDIForm1.Toolbar1.Buttons.Item(2).ToolTipText = "Mant. TABLAS"
End If
If Wflag = "A" Then
 MDIForm1.Toolbar1.Buttons.Item(13).Enabled = False
End If
ProBar.Value = ProBar.Value + 1
DoEvents
StatusBar1.Enabled = True
DoEvents
MDIForm1.Visible = True
DoEvents
MDIForm1.Show
DoEvents
ProBar.Value = ProBar.Value + 1
Screen.MousePointer = 11
PASSFORM.Hide
WF_REINICIAR = 0
ProBar.Min = 0
ProBar.Visible = False
txtderechos.Visible = False
TXTCIA.Visible = True
TXTCIA.Enabled = True
If TXTCIA.Enabled And TXTCIA.Visible Then TXTCIA.SetFocus
Screen.MousePointer = 0
MDIForm1.SubmenuTit6(6).Enabled = True

If LK_EMP = "3AA" Then
  WFECHA = #11/30/2000#
  If LK_FECHA_DIA >= WFECHA Then
  If Val(gen!GEN_CBS) <> -9 Then
sigue:
    WS = InputBox("Nº Serial es : " & Format(gen!GEN_CFP, "0000000000000000") & " , Llame al proveedor y pida su nueva serie." & Chr(13) & "{Sí intenta colocar una serie que no es, por quinta vez... el Proveedor no responde por la información almacenada.}", "Acceso Limitado ....")
    If WS = "" Then
      GoTo sigue
    End If
    wserie = Val(WS)
    If Val(PRO_DES(gen!GEN_CFP)) = Val(wserie) Then
       gen.Edit
       gen!GEN_CBS = -9
       gen.Update
    Else
      MsgBox "LLame a su Proveedor de WinSoft y pida su Nº de Serie. Ingreso Incorrecto" & Chr(13) & "{No intentar si no tiene el nuevo Nro. de Serie.}", vbCritical, "Acceso limitado .... "
      End
    End If
  End If
  End If
End If
Exit Sub
'' MsgBox LK_CODCIA & LK_FECHA_DIA
'************************
'MOSTRAR TAREAS
'************************
Dim str As String
str = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=BDATOS;Origen de datos=server"
  On Error GoTo ErrorUser
  rsT.CursorLocation = adUseClient
  rsT.Open "Exec sp_task_user '" & LK_CODUSU & "'", str, adOpenStatic, adLockOptimistic, adCmdText
  If Not rsT.RecordCount = 0 Then
   Set rsT = Nothing
   msg = "TASK"
   frmListTareas.Show 1
   Exit Sub
  Else
   Set rsT = Nothing
  End If
ErrorUser:
  Set rsT = Nothing
'************************
Exit Sub
SALE:
   If Err.Number = 360 Then
     Resume Next
   Else
      MsgBox Err.Description, 48, Pub_Titulo
   End If
   Resume Next
Exit Sub
fin_de_pro:
    Screen.MousePointer = 11
    DoEvents
    MDIForm1.WindowState = 1
    MDIForm1.Caption = "Descargando Modulos..."
    EN.Close
    CN.Close
    Screen.MousePointer = 0
    End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If WF_REINICIAR <> 2 Then
    Screen.MousePointer = 11
    MDIForm1.Caption = "Descargando Modulos..."
    DoEvents
    MDIForm1.WindowState = 1
    EN.Close
    CN.Close
    Unload MDIForm1
    Screen.MousePointer = 0
    End
End If
WF_REINICIAR = 0
End Sub
Private Sub menuAlm_Click(Index As Integer)
On Error GoTo SALE
If Trim(menuAlm(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0
Unload FORMGEN
PUB_CODTRA = Val(Left(menuAlm(Index).Caption, 4))
SQ_OPER = 1
LEER_TRA_LLAVE
If Trim(tra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(tra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(tra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
SALE:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0
End Sub


Private Sub menuAyudaSistema_Click()
MsgBox "AdminSolution  - Módulo de Administración " & Chr(13) & "  " & Chr(13) & "Casa de Desarrollo : " & Chr(13) & "Via Rápida Perú S.A.C. Trujillo - Perú " & Chr(13) & "", vbInformation, "www.viarapidaperu.com"
End Sub

Private Sub menuCia_Click(Index As Integer)
Dim iFormCount As Integer
On Error GoTo sigue
If Left(MDIForm1.menuCia(Index).Caption, 1) = "*" Then
  Exit Sub
End If
Screen.MousePointer = 11
If Forms.Count - 1 > 0 Then
   For iFormCount = Forms.Count - 1 To 1 Step -1
    If iFormCount <> 1 Then
        Unload Forms(iFormCount)
    End If
   Next iFormCount
End If
For fila = 0 To 20
  MDIForm1.menuCia(fila).Checked = False
Next fila
MDIForm1.menuCia(Index).Checked = True
MDIForm1.StatusBar1.Panels(7).text = MDIForm1.menuCia(Index).Caption
MDIForm1.TXTCIA.Caption = Trim(MDIForm1.menuCia(Index).Caption)
'*******************AGREGADO*****
Toolbar1.Buttons("CIA").ToolTipText = Trim(MDIForm1.menuCia(Index).Caption)
'*******************AGREGADO*****
ACTUALIZA_CIA Left(MDIForm1.menuCia(Index).Caption, 2)
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
Screen.MousePointer = 0
Exit Sub
sigue:
Resume Next
End Sub
Private Sub menuComp_Click(Index As Integer)
On Error GoTo SALE
If Trim(Menucomp(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0
Unload FORMGEN
PUB_CODTRA = Val(Left(Menucomp(Index).Caption, 4))
SQ_OPER = 1
LEER_TRA_LLAVE
If Trim(tra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(tra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(tra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
SALE:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0
End Sub

Private Sub MenuContab_Click(Index As Integer)
On Error GoTo SALE
If Trim(MenuContab(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0
Unload FORMGEN
PUB_CODTRA = Val(Left(MenuContab(Index).Caption, 4))
SQ_OPER = 1
LEER_TRA_LLAVE
If Trim(tra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show 1
ElseIf Trim(tra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(tra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
SALE:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0
End Sub

Private Sub menuctacte_Click()
Screen.MousePointer = 11
Load FrmCaa
FrmCaa.Show
Screen.MousePointer = 0
End Sub

Private Sub menuch_Click()
Screen.MousePointer = 11
Load frm_cheques
Screen.MousePointer = 0
frm_cheques.Show 1

End Sub

Private Sub menudis1_Click(Index As Integer)
On Error GoTo SALE
If Trim(menudis1(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0
Unload FORMGEN
SQ_OPER = 1
PUB_CODTRA = Val(Left(menudis1(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(tra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(tra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(tra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
SALE:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0

End Sub

Private Sub menudis2_Click(Index As Integer)
On Error GoTo SALE
If Trim(menudis2(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0
Unload FORMGEN
SQ_OPER = 1
PUB_CODTRA = Val(Left(menudis2(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(tra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(tra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(tra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
SALE:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0

End Sub

Private Sub menudis3_Click(Index As Integer)
On Error GoTo SALE
If Trim(menudis3(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0
Unload FORMGEN
SQ_OPER = 1
PUB_CODTRA = Val(Left(menudis3(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(tra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(tra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(tra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
SALE:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0

End Sub

Private Sub menuKa_Click()
Screen.MousePointer = 11
Load KARDEX
KARDEX.Show
Screen.MousePointer = 0
End Sub

Private Sub menuMoli_Click(Index As Integer)
If Trim(menuMoli(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0
Unload FORMGEN
PUB_CODTRA = Val(Left(menuMoli(Index).Caption, 4))
SQ_OPER = 1
LEER_TRA_LLAVE
If Trim(tra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(tra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(tra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
End Sub

Private Sub menuRepoAyuda_Click()
Dim wl As Object
Dim VAR As String
'Dim ee As Words
Dim WW As Form
Dim i
On Error GoTo OJO
VAR = Trim(MDIForm1.ActiveForm.Name)
If UCase(VAR) = "FORMGEN" Then
  VAR = Trim(MDIForm1.ActiveForm.Name) & "_" & LK_CODTRA
End If
On Error GoTo 0
If wl Is Nothing Then
    Set wl = CreateObject("word.Application")
End If
VAR = "c:\admin\ayudas\" & VAR
On Error GoTo no_existe
wl.Application.WindowState = wdWindowStateMaximize
'wl.Application.ActiveDocument.Windows(1).WindowState = 2
'ee.Windows(0).WindowState = wdWindowStateMaximize
'wl.Windows(1).WindowState = wdWindowStateMaximize
'wl.Documents
wl.Documents.Open FileName:=VAR
wl.Application.Visible = True
Set wl = Nothing
On Error GoTo 0
Exit Sub
no_existe:
'Shell "C:\Archivos de programa\Microsoft Office\Office\winword.EXE", "c:\admin\ayudas\FORMGEN_2401.doc""
MsgBox Err.Description
OJO:
If Err.Number <> 91 Then
 MsgBox "Verificar " & Err.Description, 48, Pub_Titulo
ElseIf Err.Number = 91 Then
  VAR = "MDIForm1"
  On Error GoTo 0
  If wl Is Nothing Then
    Set wl = CreateObject("word.Application")
  End If
  VAR = "c:\admin\ayudas\" & VAR
  On Error GoTo no_existe2
  wl.Application.WindowState = wdWindowStateMaximize
  wl.Documents.Open FileName:=VAR
  wl.Application.Visible = True
  On Error GoTo 0
End If
no_existe2:
End Sub

Private Sub menuVent_Click(Index As Integer)
On Error GoTo SALE
If Trim(menuVent(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0
Unload FORMGEN
SQ_OPER = 1
PUB_CODTRA = Val(Left(menuVent(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(tra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(tra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(tra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
SALE:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 ElseIf Err.Number = 40009 Then
   MsgBox "No se encontro el Registro del Reporte .Reiniciar", 48, Pub_Titulo
 End If
 Screen.MousePointer = 0
End Sub
Private Sub SubmenuTit1_Click(Index As Integer)
Dim h, wBAN As Integer
Dim iFormCount As Integer
On Error GoTo SALE
Select Case Index
    Case 0
        Screen.MousePointer = 11
        Load frmARTI
        frmARTI.Show
        Screen.MousePointer = 0
    Case 1
        Screen.MousePointer = 11
        Load frmopciones
        frmopciones.Show
        Screen.MousePointer = 0
    Case 2
        Screen.MousePointer = 11
        Load FrmVen
        FrmVen.Show
        Screen.MousePointer = 0
    Case 3
        Screen.MousePointer = 11
        Load frmBancos
        frmBancos.Show
        Screen.MousePointer = 0
    Case 5
        Screen.MousePointer = 11
        Load frmTablas
        frmTablas.Show
        Screen.MousePointer = 0
    Case 6
        Screen.MousePointer = 11

        If Forms.Count - 1 > 0 Then
           For iFormCount = Forms.Count - 1 To 1 Step -1
            If iFormCount <> 1 Then
                Unload Forms(iFormCount)
            End If
           Next iFormCount
        End If
        Load FrmParGen
        FrmParGen.Show
        Screen.MousePointer = 0
        
    Case 7
 '       Screen.MousePointer = 11
'        Load FrmTiempoReal
'        FrmTiempoReal.Show
'        Screen.MousePointer = 0
        Screen.MousePointer = 0
        Load frmGeneral
        frmGeneral.Show 1
     
    Case 8
        Screen.MousePointer = 11
        Load FrmTabla1
        FrmTabla1.Show
        Screen.MousePointer = 0
    Case 9
        Screen.MousePointer = 11
        Load FrmSub
        FrmSub.Show
        Screen.MousePointer = 0
    Case 10
        Screen.MousePointer = 11
        Load GENCAL
        GENCAL.Show
        Screen.MousePointer = 0
    Case 14
        Screen.MousePointer = 11
        Load FrmClave
        FrmClave.Show
        Screen.MousePointer = 0
    Case 15
       Dim iFrmCount As Integer
       Screen.MousePointer = 11
       For iFrmCount = Forms.Count - 1 To 1 Step -1
         If iFrmCount <> 1 Then
            Unload Forms(iFrmCount)
         End If
       Next iFrmCount
       NOTPermisos
       For fila = 0 To 20
         MDIForm1.menuCia(fila).Checked = False
       Next fila
       PUB_FLAG = 1
       Screen.MousePointer = 11
       MDIForm_Load
       Screen.MousePointer = 0
    Case 16
        'REINICIAR SISTEMA
        WF_REINICIAR = 2 'PARA VERIFICAR EN MDIFORM_UNLOAD
        Screen.MousePointer = 11
        MDIForm1.Caption = "Descargando Modulos..."
        CN.Cancel
        CN.Close
        MDIForm1.WindowState = 1
        MDIForm1.Hide
        Unload MDIForm1
        PUB_FLAG = 0
        Unload Splash
        Screen.MousePointer = 0
        Load Splash
    Case 17
        'SALIR
        Screen.MousePointer = 11
        DoEvents
        MDIForm1.WindowState = 1
        MDIForm1.Caption = "Descargando Modulos..."
        EN.Close
        CN.Close
        Screen.MousePointer = 0
        End
    Case Else
        MsgBox "INDICE FUERA DE RANGO ... Favor de llamar a Computo ", 48, Pub_Titulo
End Select
Exit Sub
SALE:
'MsgBox Err.Description
If Err.Number = 424 Then
 Screen.MousePointer = 0
 MsgBox "No tiene Acceso ", 48, Pub_Titulo
End If

End Sub

Private Sub SubmenuTit2_Click(Index As Integer)
On Error GoTo SALE
Select Case Index
    Case 0
        Screen.MousePointer = 11
        Load FORMGEN
        FORMGEN.Show
        Screen.MousePointer = 0
    Case 1
       Screen.MousePointer = 11
       Load Costos
       Costos.Show
       Screen.MousePointer = 0
    Case 3
       Screen.MousePointer = 11
       Load CIERRE_DIA
       Screen.MousePointer = 0
       CIERRE_DIA.Show 1
    Case Else
        MsgBox "INDICE FUERA DE RANGO ...Favor de llamar a Computo ..!1", 48, Pub_Titulo
End Select
Exit Sub
SALE:
If Err.Number = 424 Then
 Screen.MousePointer = 0
 MsgBox "No tiene Acceso ", 48, Pub_Titulo
End If

End Sub

Private Sub SubmenuTit6_Click(Index As Integer)
On Error GoTo SALE
Select Case Index
    Case 0
     Exit Sub
    Case 1
        Screen.MousePointer = 11
        Load frmDocu
        Screen.MousePointer = 0
        frmDocu.Show
    Case 2
        Screen.MousePointer = 11
        Load FrmPreUnidad
        FrmPreUnidad.Show
        Screen.MousePointer = 0
    Case 4
        Screen.MousePointer = 11
        Load FrmTiempoReal
        Screen.MousePointer = 0
        FrmTiempoReal.Show 1

    Case 8
       Load frmPesos
       frmPesos.Show 1
       If LKCHEK = False Then
         GoTo sale2
        End If
        Dim ArchOrigen, ArchDestino As String
        Dim Momento As String
        Screen.MousePointer = 0
        pub_mensaje = "Inserte el Disco de la Portatil en la Unidad A: , Luego Aceptar.."
        Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
        If Pub_Respuesta = vbYes Then
          Screen.MousePointer = 0
          Exit Sub
        End If
        Screen.MousePointer = 11
disco:
        On Error GoTo sale2
        
        ArchOrigen = "C:\ADMIN\PEDIDOS\VER.TXT"
        ArchDestino = "a:\VER"
        FileCopy ArchOrigen, ArchDestino
        
        Call Shell("C:\ADMIN\PEDIDOS\COPIAR.PIF", vbNormalFocus)
        Momento = Timer + 3#
        Do While Momento > Timer
         DoEvents
        Loop
        Screen.MousePointer = 0
        'msgbox "Proceso Terminado..", 48, pub_titulo
        Exit Sub
sale2:
        Screen.MousePointer = 0
    Case 6
       Screen.MousePointer = 11
       Load FORM_PED
       FORM_PED.Show
       Screen.MousePointer = 0
    Case Else
        MsgBox "INDICE FUERA DE RANGO ... Favor de llamar a Computo ", 48, Pub_Titulo
End Select
Exit Sub
SALE:
Screen.MousePointer = 0
If Err.Number = 53 Then
    MsgBox "NO esta Instalado  ...", 48, Pub_Titulo
    Exit Sub
End If
If Err.Number = 424 Then
 Screen.MousePointer = 0
 MsgBox "No tiene Acceso ", 48, Pub_Titulo
End If

End Sub

Private Sub SubmenuTit5_Click(Index As Integer)
On Error GoTo SALE
Select Case Index
    Case 0
       Call Shell("C:\MSSQL\BINN\SQLEW.EXE", 3) ' Ejecuta SQL
    Case 1
        Load frmeditor
        frmeditor.Show 1
    Case 2
      If gen!gen_bloqueo = "A" Then
        MsgBox "EN CONSTRUCCION...", 48, Pub_Titulo
        Exit Sub
      End If
       Screen.MousePointer = 11
       Load frmConfigRepo
       Screen.MousePointer = 0
       frmConfigRepo.Show 1
    Case 3
        If LK_CODUSU <> "ADMIN" Then
          Load FrmGrupos
          FrmGrupos.Show 1
        Else
        If LK_FLAG_SOS = "A" Then
          MsgBox "EN CONSTRUCCION ", 48, Pub_Titulo
          Exit Sub
        End If
          pub_mensaje = "ADMIN , desea el Formulario de Imprimer BDATOS "
          Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
          If Pub_Respuesta = vbYes Then
            Load FrmImprime
            FrmImprime.Show 1
          Else
            Load FrmGrupos
            FrmGrupos.Show 1
          End If
       End If
    Case 4
        Load frmDiseno
        frmDiseno.Show 1
    Case 5
        Load FrmDepura
        FrmDepura.Show 1
     Case 6
       Load frmTareas
       frmTareas.Show 1

    Case Else
        MsgBox "INDICE FUERA DE RANGO ...Favor de llamar a Computo ..!1", 48, Pub_Titulo
End Select

Exit Sub
SALE:
If Err.Number = 53 Then
 Screen.MousePointer = 0
 MsgBox "Si Elijio Backup, Este No Procede en Terminales. Solo en su Servidor , Consulte a su Proveedor para mas Detalles", 48, Pub_Titulo
 Exit Sub
End If
If Err.Number = 424 Then
 Screen.MousePointer = 0
 MsgBox "No tiene Acceso ", 48, Pub_Titulo
End If
End Sub
Private Sub SubmenuTit3_Click(Index As Integer)
On Error GoTo SALE
Select Case Index
    Case 0
        Screen.MousePointer = 11
        Load FrmCoPar
        Screen.MousePointer = 0
        FrmCoPar.Show 1
    Case 1
        Screen.MousePointer = 11
        Load FORM_CONT
        FORM_CONT.Show
        Screen.MousePointer = 0
    Case 2
        Screen.MousePointer = 11
        Load FORM_CONTA
        FORM_CONTA.Show
        Screen.MousePointer = 0
    Case 3
        Screen.MousePointer = 11
        Load frm_mayoriz
        Screen.MousePointer = 0
        frm_mayoriz.Show 1
    Case 4
        Screen.MousePointer = 11
        Load frm_mes_cierre
        Screen.MousePointer = 0
        frm_mes_cierre.Show 1
    
    Case 5
        Screen.MousePointer = 11
        Load GenContab
        Screen.MousePointer = 0
        GenContab.Show 1
        
    Case Else
        MsgBox "INDICE FUERA DE RANGO ...Favor de llamar a Computo ..!1", 48, Pub_Titulo
End Select
Exit Sub
SALE:
If Err.Number = 424 Then
 Screen.MousePointer = 0
 MsgBox "No tiene Acceso ", 48, Pub_Titulo
End If


End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo SAL
Dim iFormCount As Integer
Select Case Button.Key
  Case "FRMDEF"
      Screen.MousePointer = 11
      Load FrmSub
      FrmSub.Show
      Screen.MousePointer = 0
  Case "Tablas"
      Screen.MousePointer = 11
      Load FrmTabla1
      FrmTabla1.Show
      Screen.MousePointer = 0
  Case "KARDEX"
      Screen.MousePointer = 11
      Load KARDEX
      KARDEX.Show
      Screen.MousePointer = 0
  Case "PRECIOS"
      Screen.MousePointer = 11
      Load FrmPreUnidad
      FrmPreUnidad.Show
      Screen.MousePointer = 0
  Case "ARTI"
      Screen.MousePointer = 11
      Load frmARTI
      frmARTI.Show
      Screen.MousePointer = 0
  Case "CLIENTES"
      Screen.MousePointer = 11
      Load frmCLI
      frmCLI.Show
      Screen.MousePointer = 0
  Case "CUENTAS"
      Load FrmCaa
      FrmCaa.Show
  Case "CHEQUES"
      Load frm_cheques
      frm_cheques.Show 1
  Case "Password"
      Load FrmClave
      FrmClave.Show 1
  Case "BANCOS"
      Load frmBancos
      frmBancos.Show
  Case "Salir"
        Screen.MousePointer = 11
        DoEvents
        MDIForm1.WindowState = 1
        MDIForm1.Caption = "Descargando Modulos..."
        EN.Close
        CN.Close
        Screen.MousePointer = 0
        End
  Case "ALTERNO"
        Screen.MousePointer = 11
        Load FrmParGen
        FrmParGen.Show
        Screen.MousePointer = 0
  
   
  Case "Excel"
     Call Shell("c:\Archivos de programa\Microsoft Office\Office\EXCEL.EXE", 3)    ' Ejecuta Excel
     'Call Shell("c:\Archivos de programa\Microsoft Office\Office\EXCEL.EXE", 3)    ' Ejecuta Excel
     
End Select
Exit Sub

SAL:
  
  Select Case Err.Number
  Case 53
    MsgBox "No se encontro el Archivo ....!!!!..", 48, Pub_Titulo
  Case 424
    MsgBox "No existe el Formulario ....!!!!..r", 48, Pub_Titulo
  Case Else
    MsgBox Err.Number & "  " & Err.Description & "    !!!!... ", 48, Pub_Titulo
  End Select
 Screen.MousePointer = 0
  Exit Sub


End Sub

Private Sub Toolbar1_ButtonDropDown(ByVal Button As MSComctlLib.Button)
MDIForm1.PopupMenu Menucompania, , 2455, 450
End Sub

'***********************
'*************************
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Toolbar1.Buttons("CIA").ToolTipText = ButtonMenu.text

End Sub

Private Sub TXTCIA_Click()
  MDIForm1.PopupMenu Menucompania
  'MDIForm1.menuCia(
End Sub
Public Function PRO_DES(WCFP As Currency) As Currency
Dim WRES As Currency
WRES = (WCFP * 9) + 9
PRO_DES = WRES
End Function

