VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Solution-Software de Gestión para Restaurantes"
   ClientHeight    =   6510
   ClientLeft      =   -210
   ClientTop       =   735
   ClientWidth     =   9480
   Icon            =   "MDIGeneral.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "MDIGeneral.frx":0442
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3555
      Top             =   1335
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":E944
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1425
      Top             =   1545
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":ED98
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":F1EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":F640
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":FA94
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":FEE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":1033C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":10790
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":10BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":11038
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":11354
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":117A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":11BFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":12050
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":124A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":128F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":12D4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":131A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":135F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   688
      ButtonWidth     =   661
      ButtonHeight    =   635
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "FORMGEN"
            Object.ToolTipText     =   "Operaciones del Sistema"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Tablas"
            Object.ToolTipText     =   "Mant. Tablas"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "ARTI"
            Object.ToolTipText     =   "Maestro de Articulos"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "PRECIOS"
            Object.ToolTipText     =   "Actualización de Precios"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "CLIENTES"
            Object.ToolTipText     =   "Maestro de Clientes"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "BANCOS"
            Object.ToolTipText     =   "Registro de Bancos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "CUENTAS"
            Object.ToolTipText     =   "Cta. Cte. de Clientes y Proveedores"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "CHEQUES"
            Object.ToolTipText     =   "Ctas. Cte. de Bancos"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "KARDEX"
            Object.ToolTipText     =   "Kardex de Articulos"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Password"
            Object.ToolTipText     =   "Cambio de Claves de Usuarios"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "consultas"
            Object.ToolTipText     =   "Consultas de Documentos"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "ALTERNO"
            Object.ToolTipText     =   "Modo Alterno"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Excel"
            Object.ToolTipText     =   "Calculadora"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CAL"
            Object.ToolTipText     =   "Suma y Resta"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "COTI"
            Object.ToolTipText     =   "Cotizaciones"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ORDEN"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MSGBOX"
            Object.ToolTipText     =   "Mensageria OSBussines"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   16
         EndProperty
      EndProperty
      Begin VB.TextBox lblciagen 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   10785
         TabIndex        =   4
         Top             =   45
         Width           =   915
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
         Left            =   7320
         TabIndex        =   2
         Text            =   "Identificando sus accesos al sistema..."
         Top             =   0
         Visible         =   0   'False
         Width           =   4215
      End
      Begin MSComctlLib.ProgressBar ProBar 
         Height          =   180
         Left            =   7320
         TabIndex        =   1
         Top             =   180
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   135
      Top             =   2070
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   135
      Top             =   1290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport TablasC 
      Left            =   615
      Top             =   1290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\Archivos de programa\PROYECTO WINDOWS NT\tablas.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.StatusBar stb_EB 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   6180
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1765
            MinWidth        =   1765
            Text            =   "LOCAL :"
            TextSave        =   "LOCAL :"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   8819
            MinWidth        =   8819
            Picture         =   "MDIGeneral.frx":13A48
            Key             =   "cia"
            Object.ToolTipText     =   "LOCAL ACTIVO"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
            Picture         =   "MDIGeneral.frx":13D9A
            Key             =   "date"
            Object.ToolTipText     =   "Fecha del sistema"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Picture         =   "MDIGeneral.frx":140EC
            Key             =   "user"
            Object.ToolTipText     =   "USUARIO"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu menuTit1 
      Caption         =   "&Maestros"
      Enabled         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Mantenimiento y Formulacion de Platos"
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Mantenimiento de Clientes y Proveedores"
         Index           =   1
         Shortcut        =   ^I
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Mantenimiento de Series"
         Index           =   2
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Mantenimiento de Mozos"
         Index           =   3
         Shortcut        =   ^D
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Mantenimiento de Zonas"
         Index           =   4
         Shortcut        =   ^B
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Mantenimiento de Unidades de Medida"
         Index           =   5
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Mantenimiento de Familias"
         Index           =   6
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Mantenimiento de SubFamilias"
         Index           =   7
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Mantenimiento de Documentos"
         Index           =   8
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Tiempos de Envio a Cocina y Bar"
         Index           =   9
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Situacion de Mozos"
         Index           =   10
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Asignacion de Mesas a Mozos"
         Index           =   11
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Mantenimiento de Caracteristicas"
         Index           =   12
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Motivos de Anulación"
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
      Caption         =   "&Procesos"
      Enabled         =   0   'False
      Begin VB.Menu SubmenuTit2 
         Caption         =   "&Operaciones"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^G
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "Punto de Venta Carniceria"
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "Punto de Venta &Salón"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "Punto de Venta &Mostrador"
         Index           =   4
         Shortcut        =   ^M
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "Cambia Formas de Pago"
         Index           =   5
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "&Tipo de Cambio"
         Index           =   6
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "Envio Resumen Diario SUNAT"
         Index           =   7
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "&Cierre de Operaciones del Día"
         Enabled         =   0   'False
         Index           =   8
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "Listado de Productos"
         Enabled         =   0   'False
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "Mantenimiento de Zonas"
         Enabled         =   0   'False
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "Elaboración de Contratos"
         Enabled         =   0   'False
         Index           =   11
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menugrifo 
      Caption         =   "&Grifos"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu submenugrifo 
         Caption         =   "&Grifos"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu submenugrifo 
         Caption         =   "&Islas y Responsables"
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu submenugrifo 
         Caption         =   "&Parte Diario de Grifero"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu submenugrifo 
         Caption         =   "&Stock por Regla"
         Enabled         =   0   'False
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu submenugrifo 
         Caption         =   "-"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu submenugrifo 
         Caption         =   "&Transacciones II"
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menutit3 
      Caption         =   "&Controles"
      Begin VB.Menu submenutit3 
         Caption         =   "Productos por Despachar en Lista"
         Index           =   0
      End
      Begin VB.Menu submenutit3 
         Caption         =   "Productos por Despachar por Mesas"
         Index           =   1
      End
      Begin VB.Menu submenutit3 
         Caption         =   "Productos Despachados"
         Index           =   2
      End
      Begin VB.Menu submenutit3 
         Caption         =   "&Productos Despachados por Mesas"
         Index           =   3
      End
      Begin VB.Menu submenutit3 
         Caption         =   "Cierre &Del Periodo de Trabajo "
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu submenutit3 
         Caption         =   "Pase a Con&tabilidad"
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menutit4 
      Caption         =   "&Reportes"
      Enabled         =   0   'False
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
      Caption         =   "&Modulo de Deliverys"
      Begin VB.Menu submenutit5 
         Caption         =   "&Mantenimiento de Clientes Delivery"
         Index           =   0
      End
      Begin VB.Menu submenutit5 
         Caption         =   "Mantenimiento de Zonas de Reparto"
         Index           =   1
      End
      Begin VB.Menu submenutit5 
         Caption         =   "Mantenimiento de Repartidores"
         Index           =   2
      End
      Begin VB.Menu submenutit5 
         Caption         =   "Ingreso de Ventas Delivery"
         Index           =   3
      End
      Begin VB.Menu submenutit5 
         Caption         =   "Liquidacion de Repartidores"
         Index           =   4
      End
   End
   Begin VB.Menu menutit6 
      Caption         =   "&Consultas"
      Enabled         =   0   'False
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Consulta del Kardex de Productos"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^K
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Consulta de Cuentas Por Pagar y Cobrar"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^E
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Consulta de Cuentas de Bancos"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Consulta de Operaciones"
         Enabled         =   0   'False
         Index           =   4
         Shortcut        =   ^N
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Consulta de Impresion de Cuentas"
         Index           =   5
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Consulta de Comandas"
         Index           =   6
         Shortcut        =   ^L
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Consulta de Anulaciones"
         Index           =   7
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Consulta de Comensales"
         Index           =   8
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Consulta de Eventos"
         Enabled         =   0   'False
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Consulta de Descuentos Otorgados"
         Index           =   10
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Exporta Registro de Ventas al PLE"
         Index           =   11
      End
   End
   Begin VB.Menu menutit7 
      Caption         =   "&Módulo de Produccion"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu SubmenuTit7 
         Caption         =   "&Procesa Producción"
         Index           =   0
         Shortcut        =   ^T
      End
      Begin VB.Menu SubmenuTit7 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit7 
         Caption         =   "&Gestion de Contratos"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^Z
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit7 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit7 
         Caption         =   "&Registro de Requerimientos"
         Enabled         =   0   'False
         Index           =   4
         Shortcut        =   ^O
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit7 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit7 
         Caption         =   "&Bloqueo de Periodos"
         Enabled         =   0   'False
         Index           =   6
         Shortcut        =   ^Q
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit7 
         Caption         =   "-"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit7 
         Caption         =   "&Recepcion de Requerimientos"
         Enabled         =   0   'False
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit7 
         Caption         =   "-"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit7 
         Caption         =   "Generar Boletas y &Facturas"
         Enabled         =   0   'False
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit7 
         Caption         =   "-"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit7 
         Caption         =   "Datos de Transportista"
         Enabled         =   0   'False
         Index           =   12
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menuAyuda 
      Caption         =   "Deliverys"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu submenuayuda 
         Caption         =   "&Mantenimiento de Clientes Delivery"
         Index           =   0
      End
      Begin VB.Menu submenuayuda 
         Caption         =   "Mantenimiento de Zonas de Reparto"
         Index           =   1
      End
      Begin VB.Menu submenuayuda 
         Caption         =   "Mantenimiento de Repartidores"
         Index           =   2
      End
      Begin VB.Menu submenuayuda 
         Caption         =   "Ingreso de Ventas Delivery"
         Index           =   3
      End
      Begin VB.Menu submenuayuda 
         Caption         =   "Liquidacion de Repartidores"
         Index           =   4
      End
      Begin VB.Menu submenuayuda 
         Caption         =   "Formulario de Ventas"
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu submenuayuda 
         Caption         =   "Ingreso de Contratas"
         Enabled         =   0   'False
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu submenuayuda 
         Caption         =   "Reporte de Consumos Diarios"
         Enabled         =   0   'False
         Index           =   7
         Visible         =   0   'False
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
Dim sal As Boolean
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
Private Sub cmdbuscacia_Click()
  MDIForm1.PopupMenu Menucompania, , 1800, 6300
End Sub
Private Sub lblciagen_Click()
 MDIForm1.PopupMenu Menucompania, , 1800, 6300
End Sub
Private Sub MDIForm_Load()

 xParametros.Server = Trim(LeerIni("SERVER", "SERVIDOR", App.Path & "\CONFIG.DAT"))
    xParametros.User = Trim(LeerIni("SERVER", "USR", App.Path & "\CONFIG.DAT"))
    xParametros.DataBase = Trim(LeerIni("SERVER", "DB", App.Path & "\CONFIG.DAT"))
    xParametros.Password = Trim(LeerIni("SERVER", "PWD", App.Path & "\CONFIG.DAT"))
xParametros.CiaRegistro = Trim(LeerIni("SERVER", "CR", App.Path & "\CONFIG.DAT"))
    
    xParametros.CadenaConexion = "Provider=SQLOLEDB;data Source=" + xParametros.Server + ";Initial Catalog=" + xParametros.DataBase + ";User Id=" + xParametros.User + ";Password=" + xParametros.Password
    



Dim ciades As String * 80
Dim WUSU_CIAS As String
Dim WUSU_OTROS As String
Dim h, a, b, c, D, E As Integer
Dim wBAN As Integer
Dim WNUM, wPosF, WPosV, cuenta, cade
Dim WFLAG As String * 1
Dim f, g, fg As Integer
Dim wserie As Currency
Dim wfecha
Dim WS
''cmdbuscacia.Enabled = False
ciades = ""
'On Error GoTo SALE
WFLAG = ""
''TXTCIA.Enabled = False
PS_GEN(0) = 0
 GEN.Requery
Pub_Titulo = Trim(GEN!GEN_NOMBRE)
PUB_RUTA_REPORTE = Trim(Nulo_Valors(GEN!gen_ruta_reportes))
PUB_RUTA_OTRO = Trim(Nulo_Valors(GEN!gen_ruta_otros))
LK_IGV = Nulo_Valor0(GEN!GEN_IGV)
LK_DIG_RUC = Nulo_Valor0(GEN!GEN_DIG_RUC)
LK_TIPO_CAMBIO = Nulo_Valor0(GEN!gen_tipo_cambio)
LK_ACTIVA = Nulo_Valors(GEN!GEN_ACTIVA)
LK_FLAG_SOS = Nulo_Valors(GEN!gen_bloqueo)
LK_CONTROLA_STOCK = Nulo_Valors(GEN!gen_controla_stock)
LK_PASA_BOLETAS = Nulo_Valors(GEN!gen_pasa_boletas)
'LK_FACTORMAN = Nulo_Valor0(GEN!GEN_FACTORMAN)

'MDIForm1.StatusBar1.Panels(3).Text = "T.C.= S/. " + Format(LK_TIPO_CAMBIO, "0.0000")
 If LK_FLAG_GRIFO = "A" Then
  MDIForm1.menugrifo.Visible = True
  MDIForm1.menugrifo.Enabled = True
Else
  MDIForm1.menugrifo.Visible = False
  MDIForm1.menugrifo.Enabled = False
End If

'TXTCIA.Visible = False
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
ProBar.max = 15
txtderechos.Visible = True
DoEvents
ProBar.Visible = True
DoEvents
If LK_CODCIA = "" Or LK_CODUSU = "" Then
  MsgBox "Existe un Problema con su Acceso , consulte al Administrador ", 48, Pub_Titulo
  GoTo fin_de_pro
End If

'StatusBar1.Enabled = False
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
If Trim(usu!USU_KEY) = LK_CODUSU Then
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
PSUSU_LLAVE(0) = LK_CODUSU
usu_llave.Requery
WUSU_CIAS = usu_llave!USU_CIAS
'LK_FLAG_SOS = Trim(Nulo_Valors(usu_llave!USU_FLAG_CONT))
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
       ciades = PUB_CODCIA & "-" & Trim(par_llave!PAR_NOMBRE)
       MDIForm1.menuCia(cuenta - 1).Caption = ciades
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
GEN.Requery
ProBar.Value = ProBar.Value + 1
Pub_Titulo = Trim(GEN!GEN_NOMBRE)
PUB_RUTA_REPORTE = Trim(Nulo_Valors(GEN!gen_ruta_reportes))
PUB_RUTA_OTRO = Trim(Nulo_Valors(GEN!gen_ruta_otros))
'MDIForm1.StatusBar1.Panels(1).Text = "Tipo Cambio : S/." & Format(LK_TIPO_CAMBIO, "0.0000")
''MDIForm1.gfecha.Caption = Format(LK_FECHA_DIA, "dd/mm/yyyy")
'MDIForm1.StatusBar1.Panels(3).Text = Trim(GEN!gen_nombre)
''MDIForm1.gcodusu.Caption = LK_CODUSU
'MDIForm1.StatusBar1.Panels(5).Text = MDIForm1.menuCia(wBAN).Caption
ProBar.Value = ProBar.Value + 1
''MDIForm1.stb_EB.Panels("cia") = Trim(MDIForm1.menuCia(wBAN).Caption)
''MDIForm1.lblciagen.Caption = Trim(MDIForm1.menuCia(wBAN).Caption)

MDIForm1.stb_EB.Panels("cia").Text = Trim$(MDIForm1.menuCia(wBAN).Caption)
MDIForm1.stb_EB.Panels("user").Text = LK_CODUSU
MDIForm1.stb_EB.Panels("date").Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
'*******************************************************************


CONST_SERVER = Trim(Nulo_Valors(GEN!GEN_SERVER))
CONST_UID = Trim(Nulo_Valors(GEN!GEN_UID))
CONST_PWD = Trim(Nulo_Valors(GEN!GEN_PWD))
CONST_CIACENTRAL = Trim(Nulo_Valors(GEN!GEN_TIPO_CIA))
ProBar.Value = ProBar.Value + 1
If Len(CONST_SERVER) = 0 Or Len(CONST_UID) = 0 Then
  MsgBox "ADVERTENCIA ..!! DEBES DEFINIR EL NOMBRE DEL SERVIDOR , UID , PWD", 48, Pub_Titulo
End If
PUB_ODBC = "ODBC;DRIVER={SQL Server};UID=" & CONST_UID & ";PWD=" & CONST_PWD & ";APP=Data From Designer;WSID=" & CONST_SERVER & ";LANGUAGE=us_english;DATABASE=bdatos"
'CONn$ = "dsn=" & wdsn & ";uid=sa;pwd=" & wAcceso & ";database=bdatos2;"
' * Permisos para los Menus *
'Debug.Print PUB_ODBC
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
c = 0
D = 0
E = 0
f = 0
g = 0
fg = 0
Do Until tra_menu.EOF
 If tra_menu!TRA_FLAG_EXT <> "A" Then
   GoTo OTRO_MENU
 End If
 If LK_FLAG_SOS = "A" And tra_menu!TRA_ACT14 = 1 Then GoTo OTRO_MENU
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
    If c > 0 Then
        Load MDIForm1.menuAlm(c)
    End If
    MDIForm1.menuAlm(c).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    c = c + 1
 ElseIf tra_menu(0) >= 8700 And tra_menu(0) < 8899 Then
    If D > 0 Then
        Load MDIForm1.menuMoli(D)
    End If
    MDIForm1.menuMoli(D).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    D = D + 1
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
 MDIForm1.SubmenuTit4(fila).Caption = "&" + Trim(tab_mayor!tab_NOMLARGO)
 If val(tab_mayor!tab_nomcorto) = 1 Then
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

'If LK_CODUSU <> "ADMIN" Then
' MDIForm1.SubmenuTit1(8).Caption = "&Usuarios"
' MDIForm1.Toolbar1.Buttons.Item(2).ToolTipText = "Mant. Usuarios"
'Else
' MDIForm1.SubmenuTit1(8).Caption = "&Tablas"
' MDIForm1.Toolbar1.Buttons.Item(2).ToolTipText = "Mant. TABLAS"
'End If
If LK_FLAG_SOS = "A" Then
 MDIForm1.SubmenuTit1(3).Visible = False
End If
If WFLAG = "A" Then
 MDIForm1.Toolbar1.Buttons.Item(13).Enabled = False
End If

ProBar.Value = ProBar.Value + 1
DoEvents
'StatusBar1.Enabled = True
'DoEvents
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
'TXTCIA.Visible = True
''cmdbuscacia.Enabled = True
''TXTCIA.Enabled = True
''If TXTCIA.Enabled And TXTCIA.Visible Then TXTCIA.SetFocus

Screen.MousePointer = 0

If LK_EMP = "HER" Then
  wfecha = #8/31/2099#
  If LK_FECHA_DIA >= wfecha Then
  If val(GEN!GEN_CBS) <> 9 Then
sigue:
    WS = InputBox("Nº Serial es : " & Format(GEN!GEN_CFP, "0000000000000000") & " , Llame al proveedor y pida su nueva serie." & Chr(13) & "{Sí intenta colocar una serie que no es, por quinta vez... el Proveedor no responde por la información almacena.}", "Acceso Limitado ....")
    If WS = "" Then
      GoTo sigue
    End If
    wserie = val(WS)
    If val(PRO_DES(GEN!GEN_CFP)) = val(wserie) Then
       GEN.Edit
       GEN!GEN_CBS = 9
       GEN.Update
    Else
      MsgBox "LLame a su Proveedor de Solution y pida su Nº de Serie. Ingreso Incorrecto" & Chr(13) & "{No intentar si no tiene el nuevo Nro. de Serie.}", vbCritical, "Acceso limitado .... "
      End
    End If
  End If
  End If
End If


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
    Pub_ConnAdo.Close
    Screen.MousePointer = 0
    End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If LK_EMP = "PLA" Then
  If XLL Is Nothing Then
  Else
        XLL.Workbooks.Close
        XLL.APPLICATION.Quit
    End If
  Set XLL = Nothing
End If
If WF_REINICIAR <> 2 Then
    Screen.MousePointer = 11
    MDIForm1.Caption = "Descargando Modulos..."
    DoEvents
    MDIForm1.WindowState = 1
    EN.Close
    CN.Close
    Pub_ConnAdo.Close
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

pu_codtra = val(Left(menuAlm(Index).Caption, 4))
SQ_OPER = 3
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
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


'Private Sub menuAyudaSistema_Click()
'MsgBox "Vía Rápida Perú S.A.C. - www.viarapidaperu.com. Cel.949433704 RPM *366258 ", vbInformation, "www.viarapidaperu.com"
'End Sub

Private Sub menuCia_Click(Index As Integer)
Dim iFormCount As Integer
On Error GoTo sigue
If Left(MDIForm1.menuCia(Index).Caption, 1) = "*" Then
  Exit Sub
End If
Screen.MousePointer = 11
If Forms.count - 1 > 0 Then
   For iFormCount = Forms.count - 1 To 1 Step -1
    If iFormCount <> 1 Then
        Unload Forms(iFormCount)
    End If
   Next iFormCount
End If
For fila = 0 To 20
  MDIForm1.menuCia(fila).Checked = False
Next fila
MDIForm1.menuCia(Index).Checked = True
'MDIForm1.StatusBar1.Panels(7).Text = MDIForm1.menuCia(Index).Caption
''MDIForm1.stb_EB.Panels("cia") = Trim(MDIForm1.menuCia(Index).Caption)
''MDIForm1.lblciagen.Caption = Trim(MDIForm1.menuCia(Index).Caption)

MDIForm1.stb_EB.Panels("cia").Text = Trim$(MDIForm1.menuCia(Index).Caption)

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
       tab_precioss(CONTA) = tab_mayor!tab_NOMLARGO
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

pu_codtra = val(Left(Menucomp(Index).Caption, 4))
SQ_OPER = 3
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
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

pu_codtra = val(Left(MenuContab(Index).Caption, 4))
SQ_OPER = 3
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
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
If LK_FLAG_SOS = "A" Then
  Exit Sub
End If
Screen.MousePointer = 11
Load frm_cheques
Screen.MousePointer = 0
frm_cheques.Show

End Sub

Private Sub menudis1_Click(Index As Integer)
On Error GoTo SALE
If Trim(menudis1(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0

SQ_OPER = 3
pu_codtra = val(Left(menudis1(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
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

SQ_OPER = 3
pu_codtra = val(Left(menudis2(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
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

SQ_OPER = 3
pu_codtra = val(Left(menudis3(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
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

pu_codtra = val(Left(menuMoli(Index).Caption, 4))
SQ_OPER = 3
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
End Sub

Private Sub menuRepoAyuda_Click()
Dim WL As Object
Dim VAR As String
'Dim ee As Words
MsgBox "" & Chr(13) & "Visitenos en www.viarapidaperu.com", vbInformation, "Asesoria y Soporte"
Exit Sub
Dim ww As Form
Dim i
On Error GoTo OJO
VAR = Trim(MDIForm1.ActiveForm.Name)
If UCase(VAR) = "FORMGEN" Then
  VAR = Trim(MDIForm1.ActiveForm.Name) & "_" & LK_CODTRA
End If
On Error GoTo 0
If WL Is Nothing Then
    Set WL = CreateObject("word.Application")
End If
VAR = Left(PUB_RUTA_OTRO, 2) + "\admin\ayudas\" & VAR
On Error GoTo no_existe
WL.APPLICATION.WindowState = wdWindowStateMaximize
'wl.Application.ActiveDocument.Windows(1).WindowState = 2
'ee.Windows(0).WindowState = wdWindowStateMaximize
'wl.Windows(1).WindowState = wdWindowStateMaximize
'wl.Documents
WL.Documents.Open FileName:=VAR
WL.APPLICATION.Visible = True
Set WL = Nothing
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
  If WL Is Nothing Then
    Set WL = CreateObject("word.Application")
  End If
  VAR = Left(PUB_RUTA_OTRO, 2) + "\admin\ayudas\" & VAR
  On Error GoTo no_existe2
  WL.APPLICATION.WindowState = wdWindowStateMaximize
  WL.Documents.Open FileName:=VAR
  WL.APPLICATION.Visible = True
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

SQ_OPER = 3
pu_codtra = val(Left(menuVent(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
SALE:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 ElseIf Err.Number = 40009 Then
   MsgBox "No se encontro el Registro del Reporte Reiniciar", 48, Pub_Titulo
 End If
 Screen.MousePointer = 0
End Sub

Private Sub stb_EB_PanelClick(ByVal Panel As MSComctlLib.Panel)
 Select Case Panel.key
  Case "cia"
    MDIForm1.PopupMenu Menucompania, , 1750, 6000
 End Select
End Sub

Private Sub submenugrifo_Click(Index As Integer)
Select Case Index
Case 0
    'Screen.MousePointer = 11
    'PUB_TIPREG = 2102
    'PUB_CODCIA = LK_CODCIA
    'Load FrmDatArti
    'FrmDatArti.Caption = "TURNOS DE GRIFOS -  TAB_TIPREG = " & PUB_TIPREG
    'Screen.MousePointer = 0
    'FrmDatArti.Show 1
    
Case 1
    Screen.MousePointer = 11
    Load Frmislas
    Frmislas.Show
    Screen.MousePointer = 0
    Exit Sub
Case 2
    Screen.MousePointer = 11
    Load frmPARTE
    frmPARTE.Show
    Screen.MousePointer = 0
    Exit Sub
Case 3
    Screen.MousePointer = 11
    Load frmStockG
    Screen.MousePointer = 0
    frmStockG.Show 1
    Exit Sub
Case 5
    Screen.MousePointer = 11
    Load FORM_GRIFO
    FORM_GRIFO.Show
    Screen.MousePointer = 0
    Exit Sub
End Select


End Sub

Private Sub SubmenuTit1_Click(Index As Integer)
Dim h, wBAN As Integer
Dim iFormCount As Integer
On Error GoTo SALE
Select Case Index
    Case 0
        Screen.MousePointer = 11
        frmListaProd.Show
        'frmMclientes.Show
         Screen.MousePointer = 0
    Case 1
        Screen.MousePointer = 11
        Load frmCLI
        frmCLI.Show

       ' frmTiemposEnvio.Show
        ' frmMantFamilia.Show
        Screen.MousePointer = 0
    Case 2
        Screen.MousePointer = 11
        frmUsuariosSeries.Show
'        frmClientes.vTIPO = "P"
'        frmClientes.Show

        Screen.MousePointer = 0
    Case 3
        Screen.MousePointer = 11
        Load FrmVen
        FrmVen.Show
        Screen.MousePointer = 0
    Case 4
        Screen.MousePointer = 11
        frmMantZonas.Show
        'Load frmBancos
        'Unload FORMGEN
        'frmBancos.Show
        Screen.MousePointer = 0
    Case 5
        Screen.MousePointer = 11
        'Load frmTablas
        'frmTablas.Show
        frmUnidadMedida.Show
        Screen.MousePointer = 0
    Case 6
        Screen.MousePointer = 11

        'If Forms.count - 1 > 0 Then
        '   For iFormCount = Forms.count - 1 To 1 Step -1
        '    If iFormCount <> 1 Then
        '        Unload Forms(iFormCount)
        '    End If
        '   Next iFormCount
        'End If
        'Load FrmParGen
        'FrmParGen.Show
        'frmBancos.Show
        frmMantFamilia.Show
        Screen.MousePointer = 0
        
    Case 7
        Screen.MousePointer = 11
'        Load FrmTiempoReal
'        FrmTiempoReal.Show
         frmMantSubFamilia.Show
        Screen.MousePointer = 0
        'Screen.MousePointer = 0
        'Load frmGeneral
        'frmGeneral.Show 1
     
    Case 8
        Screen.MousePointer = 11
        'Load FrmTabla1
        'FrmTabla1.Show
        frmTipoDocumento.Show
        Screen.MousePointer = 0
    Case 9
        Screen.MousePointer = 11
        frmTiemposEnvio.Show
        Screen.MousePointer = 0
    Case 10
        Screen.MousePointer = 11
        frmMozosEstado.Show
        Screen.MousePointer = 0
    Case 11
        Screen.MousePointer = 11
        frmMesasMozo.Show
        Screen.MousePointer = 0
    Case 12
        Screen.MousePointer = 11
        frmsubfamiliacaracteristica.Show
        Screen.MousePointer = 0
    
    Case 13
        Screen.MousePointer = 11
        frmMantMotivos.Show
       ' frmUnidadMedida.Show
        Screen.MousePointer = 0
    
    
    
    Case 14
        Screen.MousePointer = 11
        Load FrmClave
        FrmClave.Show
        Screen.MousePointer = 0
    Case 15
       Dim iFrmCount As Integer
       Screen.MousePointer = 11
       For iFrmCount = Forms.count - 1 To 1 Step -1
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
        Pub_ConnAdo.Close
        MDIForm1.WindowState = 1
        MDIForm1.Hide
        Unload MDIForm1
        PUB_FLAG = 0
        Unload Splash
        Screen.MousePointer = 0
        Load Splash
    Case 17
        'SALIR
        If LK_EMP = "PLA" Then
          If XLL Is Nothing Then
          Else
                XLL.Workbooks.Close
                XLL.APPLICATION.Quit
            End If
          Set XLL = Nothing
        End If
        Screen.MousePointer = 11
        DoEvents
        MDIForm1.WindowState = 1
        MDIForm1.Caption = "Descargando Modulos..."
        EN.Close
        CN.Close
        Pub_ConnAdo.Close
        Screen.MousePointer = 0
        End
    Case Else
        MsgBox "INDICE FUERA DE RANGO ... ", 48, Pub_Titulo
End Select
Exit Sub
SALE:
'MsgBox Err.Description
Resume Next
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
        frmVentas2.Show
        Screen.MousePointer = 0
    Case 2
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_VERIFICA_FORMA_COMANDA"
    Dim ORSfORMA As ADODB.Recordset
    Set ORSfORMA = oCmdEjec.Execute(, LK_CODUSU)
    
    If Not ORSfORMA.EOF Then
        If ORSfORMA!Comanda Then
        frmMainMozos.Show
        Else
            frmZonas.Show
        End If
        Else
        MsgBox "Error", vbCritical, Pub_Titulo
    End If
        
        'frmComanda2.Show
    Case 3
        Screen.MousePointer = 11
        Load frmletras 'frmPesos
        Screen.MousePointer = 0
        frmletras.Show
        
    Case 4
        frmComanda2.Show
    Case 5
        frmChangeFormasPago.Show
    Case 6
       Screen.MousePointer = 11
       Load Frmcospro
       Frmcospro.Show
       Screen.MousePointer = 0
     Case 7
       frmResumenDiario.Show
     Case 8
       Screen.MousePointer = 11
       Load PRODIA
       Screen.MousePointer = 0
       PRODIA.Show 1
    Case 9
        frmListaProd.Show
    Case 10
        frmMantZonas.Show
        Case 11
        'frmContratos.Show
       ' frmContratosList.Show
    Case Else
        MsgBox "INDICE FUERA DE RANGO ...", 48, Pub_Titulo
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
        Screen.MousePointer = 11
        Load KARDEX
        KARDEX.Show
        Screen.MousePointer = 0
    Case 1
        Screen.MousePointer = 11
        Load FrmCaa
        FrmCaa.Show
        Screen.MousePointer = 0
    Case 2
        Screen.MousePointer = 11
        Load frm_cheques
        frm_cheques.Show
        Screen.MousePointer = 0
    Case 4
        Screen.MousePointer = 11
        Load frmdocu
        frmdocu.Show
        Screen.MousePointer = 0
    Case 5
        Screen.MousePointer = 11
        frmConsultaPrintCuentaLog.Show
        Screen.MousePointer = 0
    Case 6
        Screen.MousePointer = 11
        Load FORM_COT
        FORM_COT.Show
        Screen.MousePointer = 0
    Case 7
        Screen.MousePointer = 11
        'Load frmplatdesp
        Screen.MousePointer = 0
        'frmplatdesp.Show 1
        frmConsultaPlatosExtornados.Show
        
    Case 8
        Screen.MousePointer = 11
        'Load frmbar
        frmConsultaComensales.Show
        Screen.MousePointer = 0
       ' frmbar.Show 1
 
    Case 9
        Screen.MousePointer = 11
        frmEventoConsulta.Show
        Screen.MousePointer = 0
        
    Case 10
        Screen.MousePointer = 11
        frmResumenDescuentos.Show
        Screen.MousePointer = 0
        
    Case 11
        Screen.MousePointer = 11
        frmPLE.Show
        Screen.MousePointer = 0
        
        
        
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
Private Sub Submenuayuda_Click(Index As Integer)
On Error GoTo SALE
Select Case Index
    Case 0
        Screen.MousePointer = 11
        'frmConversion.Show
         'frmTurno.Show
         frmClientes.vTIPO = "C"
         frmClientes.Show
        Screen.MousePointer = 0
    Case 1
        Screen.MousePointer = 11
        'frmProgramacionMain.Show
        frmZonaReparto.Show
        Screen.MousePointer = 0
    Case 2
        Screen.MousePointer = 11
       'frmProgramacionRequerida.Show
       frmRepartidores.Show
        Screen.MousePointer = 0
   Case 3
        Screen.MousePointer = 11
       'frmProcesarIngreso.Show
       frmDeliveryMain.Show
        Screen.MousePointer = 0
   Case 4
        Screen.MousePointer = 11
       'frmConsultaAtenciones.Show
       frmConsultaRepartidores.Show
        Screen.MousePointer = 0
  Case 5
        Screen.MousePointer = 11
       'frmConsultaAtenciones1.Show
        frmVenta.Show
        Screen.MousePointer = 0
  Case 6
        Screen.MousePointer = vbHourglass
        frmContrata.Show
        Screen.MousePointer = vbDefault
        Case 7
         Screen.MousePointer = vbHourglass
        frmConsultaAtenciones2.Show
        Screen.MousePointer = vbDefault
        
        
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
'    Case 0
'       Call Shell("C:\MSSQL\BINN\SQLEW.EXE", 3) ' Ejecuta SQL
'    Case 1
'        Load frmeditor
'        frmeditor.Show 1
'    Case 2
'       Screen.MousePointer = 11
'       Unload FORMGEN
'       Load frmConfigRepo
'       Screen.MousePointer = 0
'       frmConfigRepo.Show 1
'    Case 3
'        If LK_CODUSU <> "ADMIN" Then
'          Load FrmGrupos
'          FrmGrupos.Show 1
'        Else
'          pub_mensaje = "ADMIN , desea el Formulario de Imprimer bdatos "
'          Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
'          If Pub_Respuesta = vbYes Then
'            Load FrmImprime
'            FrmImprime.Show 1
'          Else
'            Load FrmGrupos
'            FrmGrupos.Show 1
'          End If
'       End If
'    Case 4
'        Load frmDiseno
'        frmDiseno.Show 1
'    Case Else
'        MsgBox "INDICE FUERA DE RANGO ...", 48, Pub_Titulo

Case 0
        Screen.MousePointer = 11
        'frmConversion.Show
         'frmTurno.Show
         frmClientes.vTIPO = "C"
         frmClientes.Show
        Screen.MousePointer = 0
    Case 1
        Screen.MousePointer = 11
        'frmProgramacionMain.Show
        frmZonaReparto.Show
        Screen.MousePointer = 0
    Case 2
        Screen.MousePointer = 11
       'frmProgramacionRequerida.Show
       frmRepartidores.Show
        Screen.MousePointer = 0
   Case 3
        Screen.MousePointer = 11
       'frmProcesarIngreso.Show
       frmDeliveryMain.Show
        Screen.MousePointer = 0
   Case 4
        Screen.MousePointer = 11
       'frmConsultaAtenciones.Show
       frmConsultaRepartidores.Show
        Screen.MousePointer = 0
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
        frmPorDespachar.Show
        Screen.MousePointer = 0
    Case 1
        Screen.MousePointer = 11
        frmPorDespachar1.Show
        Screen.MousePointer = 0
    Case 2
        Screen.MousePointer = 11
        frmdespachados.Show
        Screen.MousePointer = 0
    Case 3
        Screen.MousePointer = 11
        frmConsultaDespachos.Show
        Screen.MousePointer = 0
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
        MsgBox "INDICE FUERA DE RANGO ...", 48, Pub_Titulo
End Select
Exit Sub
SALE:
If Err.Number = 424 Then
 Screen.MousePointer = 0
 MsgBox "No tiene Acceso ", 48, Pub_Titulo
End If


End Sub

Private Sub submenutit7_Click(Index As Integer)
  Select Case Index
    Case 0
       Screen.MousePointer = 11
'       Load frmContratosList
'       frmContratosList.Show
        frmProduccion.Show
       Screen.MousePointer = 0
    Case 2
       'MsgBox "Error de componente ... ", 48, Pub_Titulo
       frmContratosList2.Show
       'Screen.MousePointer = 11
       'Load frmCotizacion  'FORM_COTfrmOCProv
       'frmCotizacion.Show
       'Screen.MousePointer = 0
    Case 4
        Screen.MousePointer = 11
       frmRequerimiento.Show
       Screen.MousePointer = 0
   Case 6
       Screen.MousePointer = 11
       Load frmOpcion
       Screen.MousePointer = 0
       frmOpcion.Show 1
    Case 8
       Screen.MousePointer = 11
        frmRequerimientoRecepcion.Show
       Screen.MousePointer = 0
      
    Case 10
       Screen.MousePointer = 11
       PUB_TIPOPRINT = 1
       Load FrmRepo
       FrmRepo.Caption = "Impresion por Lotes de Facturas y Boletas"
       FrmRepo.cheguia.Visible = False
       FrmRepo.cmdMostrar.Caption = "&Imprimir Documentos"
       Screen.MousePointer = 0
       FrmRepo.Show 1
    Case 12
       Screen.MousePointer = 11
       Load frmTransportista
       Screen.MousePointer = 0
       frmTransportista.Show
    Case Else
        MsgBox "INDICE FUERA DE RANGO ... ", 48, Pub_Titulo
End Select
End Sub


'Private Sub Timer1_Timer()
'On Error GoTo SALE
'Timer1.Enabled = False
'If Trim(MDIForm1.gfecha.Caption) <> "" Then
'  MDIForm1.gfecha.Caption = ""
'Else
'  MDIForm1.gfecha.Caption = Format(LK_FECHA_DIA, "dd/mm/yyyy")
'End If
'PSMSG_LLAVE(0) = LK_CODUSU
'msg_llave.Requery
'If Trim(msg_llave!USU_MSGBOX) <> "" Then
'   MsgBox "Usted. Tiene un Mensaje de :" & Chr(13) & Chr(13) & Trim(msg_llave!USU_MSGBOX), vbExclamation, "Mensajeria Solution"
'   msg_llave.Edit
'   msg_llave!USU_MSGBOX = ""
'   msg_llave.Update
'End If
'SALE:
'
'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RetVal
Dim wclave As String
On Error GoTo sal
Dim iFormCount As Integer
Dim WCAMBIA
Select Case Button.key
  Case "FORMGEN"
      Screen.MousePointer = 11
      Load FORMGEN
      FORMGEN.Show
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
      If LK_FLAG_SOS = "A" Then
         MsgBox "En Contrucción", 48, Pub_Titulo
         Exit Sub
      End If

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
  Case "MSGBOX"
      Load FrmMsgbox
      FrmMsgbox.Show 1
  Case "CHEQUES"
     If LK_FLAG_SOS = "A" Then
        Screen.MousePointer = 0
        Exit Sub
      End If
      Load frm_cheques
      frm_cheques.Show
  Case "Password"
      Load FrmClave
      FrmClave.Show 1
  Case "BANCOS"
      If LK_FLAG_SOS = "A" Then
        Screen.MousePointer = 0
        Exit Sub
      End If
      Load frmBancos
      frmBancos.Show
  Case "consultas"
      Screen.MousePointer = 11
      Load frmdocu
      frmdocu.Show
      Screen.MousePointer = 0
  Case "CAL"
      Load FrmCal
      FrmCal.Show 1
  Case "COTI"
      Screen.MousePointer = 11
      Load FORM_COT
      FORM_COT.Show
      Screen.MousePointer = 0
  Case "ORDEN"
      Screen.MousePointer = 11
      Load FORM_PED
      FORM_PED.Show
      Screen.MousePointer = 0
  Case "Salir"
        Screen.MousePointer = 11
        DoEvents
        MDIForm1.WindowState = 1
        MDIForm1.Caption = "Descargando Modulos..."
        EN.Close
        CN.Close
        Pub_ConnAdo.Close
        Screen.MousePointer = 0
        If LK_EMP = "PLA" Then
          If XLL Is Nothing Then
          Else
                XLL.Workbooks.Close
                XLL.APPLICATION.Quit
            End If
          Set XLL = Nothing
         End If
         End
  Case "ALTERNO"
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      Toolbar1.Buttons.Item(12).ToolTipText = "Busqueda nombre de articulo."
      LK_FLAG_ALTERNO = " "
      Toolbar1.Buttons.Item(12).Image = 18 '********** tenr en cuent
      pub_cadena = "SELECT ART_KEY FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY <> 0 order by art_NOMBRE"
      'DoEvents
    Else
      Toolbar1.Buttons.Item(12).ToolTipText = "Busqueda sub-codigo de articulo."
      LK_FLAG_ALTERNO = "A"
      Toolbar1.Buttons.Item(12).Image = 5
      pub_cadena = "SELECT ART_ALTERNO FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY <> 0 order by art_alterno"
    End If
    Set PSMANO_CODI = CN.CreateQuery("", pub_cadena)
    PSMANO_CODI.rdoParameters(0) = 0
    Set mano_CODI = PSMANO_CODI.OpenResultset(rdOpenKeyset, rdConcurValues)
    PSMANO_CODI(0) = LK_CODCIA
    mano_CODI.Requery
      
   If Forms.count - 1 > 0 Then
     For iFormCount = Forms.count - 1 To 1 Step -1
        If iFormCount <> 1 Then
           If UCase(Forms(iFormCount).Name) = "FRMARTI" Then
                'MsgBox UCase(Forms(iFormCount).Name)
                frmARTI.PROCESO_CANCELAR
                frmARTI.PROCESO_ARTI
           ElseIf UCase(Forms(iFormCount).Name) = "KARDEX" Then
                KARDEX.ListView1.Visible = False
                KARDEX.i_codart2.Text = ""
           ElseIf UCase(Forms(iFormCount).Name) = "FRMPREUNIDAD" Then
                FrmPreUnidad.PRO_GOTFOCUS
                FrmPreUnidad.ListView1.Visible = False
                FrmPreUnidad.Txt_key.Text = ""
           ElseIf UCase(Forms(iFormCount).Name) = "FRMREPO" Then
                FrmRepo.ListView1.Visible = False
                FrmRepo.Txt_key.Text = ""
           ElseIf UCase(Forms(iFormCount).Name) = "FORMGEN" Then
                FORMGEN.LV_ART.Visible = False
                FORMGEN.TEXTOVAR.Text = ""
           End If
        End If
       Next iFormCount
   End If

  Case "Excel"
   '  If Trim(LK_FLAG_SOS) = "" Then
   '    LK_FLAG_SOS = "A"
   '    MsgBox "activo"
   '  Else
   '    LK_FLAG_SOS = ""
   '    MsgBox "Des activo"
   '   End If
   '   LK_ACTIVA = "A"
   RetVal = Shell("C:\WINDOWS\CALC.EXE", 1)    ' Ejecuta Calculadora.
    Exit Sub
   If LK_EMP <> "CAM" Then
     ACT_TIPO_CAMBIO
     Exit Sub
   End If
    pub_cadena = "SELECT * FROM CONTROLL"
    Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
    GEN.Requery
    If Trim(LK_FLAG_SOS) <> Trim(UCase(GEN!gen_bloqueo)) Then
        con_llave.Close
        MsgBox "Listo!!!", 48, Pub_Titulo
        End
    End If
    If Trim(UCase(GEN!gen_bloqueo)) = "A" Then
      If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
          wclave = InputBox("Desactivar:", "Activar/Desactivar", "**********")
          If wclave = Trim(GEN!GEN_PWD) Then
            GEN.Edit
            GEN!gen_bloqueo = " "
            GEN.Update
            con_llave.Close
            MsgBox "Control Desactivado", 48, Pub_Titulo
            End
          Else
            con_llave.Close
            Exit Sub
          End If
      Else
'         MsgBox "Control fue Activado!!!", 48, Pub_Titulo
         con_llave.Close
         Exit Sub
      End If
    End If
    GEN.Edit
    GEN!gen_bloqueo = "A"
    GEN.Update
    con_llave.Close
    If Trim(UCase(GEN!gen_bloqueo)) = "A" Then
       MsgBox "Activado Control", 48, Pub_Titulo
       End
    End If
  
    'If LK_FLAG_GRIFO = "A" Then
    '     WCAMBIA = InputBox("Ingrese Fecha: " & LK_FECHA_DIA & Chr(13))
    '     If WCAMBIA = "" Then Exit Sub
    '     If IsDate(WCAMBIA) = False Then Exit Sub
    '     LK_FECHA_DIA = WCAMBIA
    '     MDIForm1.StatusBar1.Panels(2).Text = "Fecha de Empresa : " & Format(LK_FECHA_DIA, "dd, mmmm yyyy")
    '     Exit Sub
    'End If
    'ACT_TIPO_CAMBIO
End Select
Exit Sub

sal:
  
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

Private Sub TXTCIA_Click()
  MDIForm1.PopupMenu Menucompania, 0, 7400, True
End Sub
Public Function PRO_DES(WCFP As Currency) As Currency
Dim WRES As Currency
WRES = (WCFP * 9) + 9
PRO_DES = WRES
End Function


Public Sub VER_MENSAJES()
PSMEN_LLAVE(0) = ""
men_llave.Requery
End Sub
