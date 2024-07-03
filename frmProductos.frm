VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form frmProductos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información del Producto"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   9720
      TabIndex        =   58
      Top             =   7680
      Width           =   1335
   End
   Begin TabDlg.SSTab stabMain 
      Height          =   7455
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   13150
      _Version        =   393216
      TabOrientation  =   3
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      TabCaption(0)   =   "Datos del Producto"
      TabPicture(0)   =   "frmProductos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Formulación de Insumos y/o Materia Prima"
      TabPicture(1)   =   "frmProductos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "REGISTRO DE INFORMACIÓN"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   37
         Top             =   80
         Width           =   10095
         Begin VB.Label lblUniMedF 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2160
            TabIndex        =   43
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lblDescripcionF 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2160
            TabIndex        =   42
            Top             =   720
            Width           =   6615
         End
         Begin VB.Label lblCodigoF 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2160
            TabIndex        =   41
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "UNIDAD DE MEDIDA:"
            Height          =   195
            Index           =   18
            Left            =   120
            TabIndex        =   40
            Top             =   1125
            Width           =   1845
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPCIÓN:"
            Height          =   195
            Index           =   17
            Left            =   600
            TabIndex        =   39
            Top             =   765
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CODIGO:"
            Height          =   195
            Index           =   16
            Left            =   1080
            TabIndex        =   38
            Top             =   405
            Width           =   825
         End
      End
      Begin VB.Frame fra1 
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   40
         Width           =   9975
         Begin VB.ComboBox ComCboTipoProd 
            Height          =   315
            ItemData        =   "frmProductos.frx":0038
            Left            =   840
            List            =   "frmProductos.frx":0045
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtCodAlt 
            Height          =   285
            Left            =   7920
            TabIndex        =   0
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CODIGO:"
            Height          =   195
            Index           =   4
            Left            =   3360
            TabIndex        =   22
            Top             =   285
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CODIGO ALTERNO:"
            Height          =   195
            Index           =   2
            Left            =   6000
            TabIndex        =   21
            Top             =   285
            Width           =   1680
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "TIPO:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   285
            Width           =   495
         End
         Begin VB.Label lblCodigo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4320
            TabIndex        =   19
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame fra2 
         Height          =   4095
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   9975
         Begin VB.TextBox txtCodBolsa 
            Height          =   285
            Left            =   6000
            TabIndex        =   80
            Top             =   3600
            Width           =   975
         End
         Begin VB.TextBox txtBolsas 
            Height          =   285
            Left            =   3120
            MaxLength       =   2
            TabIndex        =   79
            Text            =   "0"
            Top             =   3600
            Width           =   975
         End
         Begin VB.TextBox txtCosto 
            Height          =   285
            Left            =   3480
            TabIndex        =   74
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox txtPorcion 
            Height          =   285
            Left            =   8760
            TabIndex        =   71
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox chkPorcion 
            Caption         =   "Porcion"
            Height          =   255
            Left            =   7680
            TabIndex        =   70
            Top             =   600
            Width           =   975
         End
         Begin MSDataListLib.DataCombo DatUM 
            Height          =   315
            Left            =   240
            TabIndex        =   68
            Top             =   1320
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.CheckBox chkPrioritario2 
            Caption         =   "Inafecto IGV"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   8160
            TabIndex        =   67
            Top             =   3600
            Width           =   1815
         End
         Begin VB.CommandButton cmdImagenDEL 
            Height          =   615
            Left            =   9000
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   1080
            Width           =   855
         End
         Begin VB.CommandButton cmdImagen 
            Height          =   615
            Left            =   8040
            Picture         =   "frmProductos.frx":0062
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtMax 
            Height          =   285
            Left            =   3480
            TabIndex        =   61
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton cmdSubFam 
            Caption         =   "..."
            Height          =   255
            Left            =   2760
            TabIndex        =   60
            Top             =   2520
            Width           =   495
         End
         Begin VB.CommandButton cmdFam 
            Caption         =   "..."
            Height          =   255
            Left            =   2760
            TabIndex        =   59
            Top             =   1920
            Width           =   495
         End
         Begin VB.CheckBox chkPri 
            Caption         =   "Afecto ICBPER"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   3600
            Width           =   1815
         End
         Begin VB.CheckBox chkSituacion 
            Caption         =   "Habilitado"
            Height          =   255
            Left            =   8520
            TabIndex        =   56
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtpvd6 
            Height          =   285
            Left            =   6600
            TabIndex        =   16
            Top             =   3000
            Width           =   975
         End
         Begin VB.TextBox txtpvd5 
            Height          =   285
            Left            =   6600
            TabIndex        =   15
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox txtpvd4 
            Height          =   285
            Left            =   6600
            TabIndex        =   14
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox txtpvd3 
            Height          =   285
            Left            =   6600
            TabIndex        =   13
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox txtpvd2 
            Height          =   285
            Left            =   6600
            TabIndex        =   12
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox txtpvd1 
            Height          =   285
            Left            =   6600
            TabIndex        =   11
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtpv1 
            Height          =   285
            Left            =   5520
            TabIndex        =   5
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   285
            Left            =   240
            TabIndex        =   1
            Top             =   600
            Width           =   7335
         End
         Begin VB.TextBox txtpv2 
            Height          =   285
            Left            =   5520
            TabIndex        =   6
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox txtpv3 
            Height          =   285
            Left            =   5520
            TabIndex        =   7
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox txtpv5 
            Height          =   285
            Left            =   5520
            TabIndex        =   9
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox txtpv6 
            Height          =   285
            Left            =   5520
            TabIndex        =   10
            Top             =   3000
            Width           =   975
         End
         Begin VB.TextBox txtpv4 
            Height          =   285
            Left            =   5520
            TabIndex        =   8
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox txtMin 
            Height          =   285
            Left            =   3480
            TabIndex        =   4
            Top             =   1320
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo dcboFam 
            Height          =   315
            Left            =   240
            TabIndex        =   2
            Top             =   1920
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcboSubFam 
            Height          =   315
            Left            =   240
            TabIndex        =   3
            Top             =   2520
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.CheckBox chkStock 
            Caption         =   "Controla Stock"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   3240
            Width           =   2895
         End
         Begin VB.TextBox txtProporcion 
            Height          =   285
            Left            =   7800
            TabIndex        =   62
            Top             =   3000
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código Bolsa:"
            Height          =   195
            Index           =   21
            Left            =   4680
            TabIndex        =   81
            Top             =   3645
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BOLSAS:"
            Height          =   195
            Index           =   20
            Left            =   2160
            TabIndex        =   78
            Top             =   3645
            Width           =   780
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "% COSTO"
            Height          =   195
            Left            =   3480
            TabIndex        =   76
            Top             =   2880
            Width           =   870
         End
         Begin VB.Label lblPorcentaje 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   3480
            TabIndex        =   75
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "COSTO:"
            Height          =   195
            Left            =   3480
            TabIndex        =   73
            Top             =   2280
            Width           =   705
         End
         Begin VB.Image ipic 
            Height          =   1575
            Left            =   8040
            Stretch         =   -1  'True
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "P Delivery"
            Height          =   195
            Left            =   6600
            TabIndex        =   55
            Top             =   960
            Width           =   885
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "P Salón"
            Height          =   195
            Left            =   5640
            TabIndex        =   54
            Top             =   960
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "PV1"
            Height          =   195
            Index           =   14
            Left            =   5040
            TabIndex        =   53
            Top             =   1245
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "STOCK MIN.:"
            Height          =   195
            Index           =   9
            Left            =   3480
            TabIndex        =   34
            Top             =   1080
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "STOCK MAX.:"
            Height          =   195
            Index           =   8
            Left            =   3480
            TabIndex        =   33
            Top             =   1680
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "SUBFAMILIA:"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   32
            Top             =   2280
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "FAMILIA:"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   31
            Top             =   1680
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "UNIDAD DE MEDIDA:"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   30
            Top             =   1080
            Width           =   1845
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "PV2"
            Height          =   195
            Index           =   3
            Left            =   5040
            TabIndex        =   29
            Top             =   1605
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPCIÓN DEL PRODUCTO:"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "PV5"
            Height          =   195
            Index           =   11
            Left            =   5040
            TabIndex        =   27
            Top             =   2685
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "PV4"
            Height          =   195
            Index           =   12
            Left            =   5040
            TabIndex        =   26
            Top             =   2325
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "PV3"
            Height          =   195
            Index           =   13
            Left            =   5040
            TabIndex        =   25
            Top             =   1965
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "PV6"
            Height          =   195
            Index           =   15
            Left            =   5040
            TabIndex        =   24
            Top             =   3045
            Width           =   330
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "PROPORCIÓN TOTAL"
            Height          =   435
            Index           =   26
            Left            =   7740
            TabIndex        =   63
            Top             =   2520
            Visible         =   0   'False
            Width           =   1320
         End
      End
      Begin VB.Frame fra3 
         Height          =   2655
         Left            =   120
         TabIndex        =   35
         Top             =   4680
         Width           =   9975
         Begin VB.CommandButton cmdQuitarC 
            Caption         =   "==>"
            Height          =   375
            Left            =   9240
            TabIndex        =   51
            Top             =   720
            Width           =   615
         End
         Begin MSComctlLib.ListView lvComposicion 
            Height          =   2175
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   3836
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.CommandButton cmdAddComp 
            Caption         =   "<=="
            Height          =   375
            Left            =   9240
            TabIndex        =   47
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "COMPOSICIÓN DE ELEMENTOS Y/O PIEZAS QUE INTEGRAN EL PRODUCTO"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   36
            Top             =   120
            Width           =   6450
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5400
         Left            =   -74880
         TabIndex        =   44
         Top             =   1460
         Width           =   10095
         Begin VB.CommandButton cmdQuitarF 
            Caption         =   "==>"
            Height          =   375
            Left            =   9360
            TabIndex        =   52
            Top             =   960
            Width           =   615
         End
         Begin VB.CommandButton cmdAddForm 
            Caption         =   "<=="
            Height          =   375
            Left            =   9360
            TabIndex        =   49
            Top             =   600
            Width           =   615
         End
         Begin MSComctlLib.ListView lvFormulacion 
            Height          =   4335
            Left            =   120
            TabIndex        =   48
            Top             =   600
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   7646
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label lblTot 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7920
            TabIndex        =   72
            Top             =   5040
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "FORMULACIÓN DE INSUMOS Y/O MATERIA PRIMA"
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   4290
         End
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   8280
      TabIndex        =   50
      Top             =   7680
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   480
      TabIndex        =   77
      Top             =   4080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
End
Attribute VB_Name = "frmProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VNuevo As Boolean
Private oRS As ADODB.Recordset
Private oRsCombos As ADODB.Recordset
Public vCodigo As Double
Public cTipo As String

Public vGraba As Boolean
Public vSit As Boolean
Private oRsFam As ADODB.Recordset
Private oRsSubFam As ADODB.Recordset
Private vGRABAi As Boolean 'para grabar la imagen
Private oRSum As ADODB.Recordset
Private vCarga As Boolean

Private Sub CreaRS()
Set oRsFam = New ADODB.Recordset
 oRsFam.Fields.Append "Codigo", adInteger
oRsFam.Fields.Append "Familia", adVarChar, 100
oRsFam.CursorLocation = adUseClient
oRsFam.Open

Set oRsSubFam = New ADODB.Recordset
 oRsSubFam.Fields.Append "Codigo", adInteger 'codigo de sub familia
 oRsSubFam.Fields.Append "codFam", adInteger 'codigo e sub familia
oRsSubFam.Fields.Append "SubFamilia", adVarChar, 100
oRsSubFam.CursorLocation = adUseClient
oRsSubFam.Open
End Sub

Private Sub ConfiguraFormulacion()
With Me.lvFormulacion
    .ColumnHeaders.Add , , "Codigo", 800
    .ColumnHeaders.Add , , "Descripción", 3000
    .ColumnHeaders.Add , , "Unidad"
    If frmListaProd.cboTipoProd.ListIndex = 3 Then 'materia prima
    .ColumnHeaders.Add , , "Proporción"
    ElseIf frmListaProd.cboTipoProd.ListIndex = 1 Then 'INSUMO
    .ColumnHeaders.Add , , "Cantidad/Kg."
    .ColumnHeaders.Add , , "CostoUnit"
    .ColumnHeaders.Add , , "Costo"
    Else
    .ColumnHeaders.Add , , "Cantidad/Kg."
    .ColumnHeaders.Add , , "CostoUnit"
    .ColumnHeaders.Add , , "Costo"
    End If
    .View = lvwReport
    .LabelEdit = lvwManual
   .Gridlines = True
    .FullRowSelect = True
End With
End Sub

Private Sub ConfiguraComposicion()
With Me.lvComposicion
    .ColumnHeaders.Add , , "Codigo", 900
    .ColumnHeaders.Add , , "Descripción", 4000
    .ColumnHeaders.Add , , "Unidad"
    .ColumnHeaders.Add , , "Cantidad"
    .ColumnHeaders.Add , , "CostoUnit"
    .ColumnHeaders.Add , , "Costo"
    .View = lvwReport
    .LabelEdit = lvwManual
    .Gridlines = True
    .FullRowSelect = True
End With
End Sub

Private Sub cboforma_Click()
 
End Sub


Private Sub chkSituacion_Click()
vSit = Me.chkSituacion.Value
End Sub

Private Sub cmdAddComp_Click()
If cTipo <> "M" And cTipo <> "P" Then
    Select Case cTipo
    Case "P": frmBusProd.vTIPO = "M"
    Case "C": frmBusProd.vTIPO = "P"
    Case "M": frmBusProd.vTIPO = "M"
    End Select
    frmBusProd.Caption = "Elementos y/o Piezas que integran el Producto"
    frmBusProd.Show vbModal
Else
    MsgBox "No se permite la operación", vbInformation, "Error"
End If
End Sub

Private Sub cmdAddForm_Click()
If cTipo = "P" Or cTipo = "M" Then
    frmBusProd.vTIPO = "M"
    frmBusProd.Caption = "Insumos y/o Materia Prima"
    frmBusProd.Show vbModal
Else
    MsgBox "No se permite la operación", vbInformation, "Error"
End If
End Sub

Private Function ArmaXmlComposicion() As String

    Dim valor As String

    Dim Item  As Object

    If Me.lvComposicion.ListItems.count > 0 Then

        valor = "<r>"

        For Each Item In Me.lvComposicion.ListItems

            valor = valor & "<d "
            valor = valor & "idp=""" & Trim(Item.Text) & """ "
            valor = valor & "c=""" & Trim(Item.SubItems(3)) & """ "
            valor = valor & "/>"
        Next

        valor = valor & "</r>"
    ElseIf Me.lvFormulacion.ListItems.count > 0 Then
        valor = "<r>"

        For Each Item In Me.lvFormulacion.ListItems

            valor = valor & "<d "
            valor = valor & "idp=""" & Trim(Item.Text) & """ "
            valor = valor & "c=""" & Trim(Item.SubItems(3)) & """ "
            valor = valor & "/>"
        Next

        valor = valor & "</r>"
    End If

    ArmaXmlComposicion = valor
End Function

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub Almacena()
    'VALIDA QUE EL CODIDO ALTERNATIVO NO ESTE REPETIDO
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPPRODUCTO_VALIDACODIGOALT_REPETIDO"

    Dim oRSr As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@alterno", adVarChar, adParamInput, 50, Trim(Me.txtCodAlt.Text))

    If Not VNuevo Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codigo", adDouble, adParamInput, , CDbl(Me.lblCodigo.Caption))

    End If

    Set oRSr = oCmdEjec.Execute

    If Not oRSr.EOF Then
        If CBool(oRSr!valor) Then
            MsgBox "El Codigo Alternativo ya se encuentra registrado.", vbCritical, Pub_Titulo

            Exit Sub

        End If

    End If

    LimpiaParametros oCmdEjec

    Dim vxmlCompform As String

    vxmlCompform = ArmaXmlComposicion
    
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Descrip", adVarChar, adParamInput, 70, Me.txtDescripcion.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@alterno", adVarChar, adParamInput, 20, Trim(Me.txtCodAlt.Text))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@unidad", adVarChar, adParamInput, 20, Me.DatUM.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codfam", adInteger, adParamInput, , CInt(Me.dcboFam.BoundText))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codsubfam", adInteger, adParamInput, , CInt(Me.dcboSubFam.BoundText))

    If Len(Trim(Me.txtMin.Text)) = 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@stockmin", adDouble, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@stockmin", adDouble, adParamInput, , CDbl(Me.txtMin.Text))

    End If

    If Len(Trim(Me.txtMax.Text)) = 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@stockmax", adDouble, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@stockmax", adDouble, adParamInput, , CDec(Me.txtMax.Text))

    End If
    
    If Len(Trim(Me.txtpv1.Text)) = 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp1", adDouble, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp1", adDouble, adParamInput, , CDec(Me.txtpv1.Text))

    End If

    If Len(Trim(Me.txtpv2.Text)) = 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp2", adDouble, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp2", adDouble, adParamInput, , CDec(Me.txtpv2.Text))

    End If

    If Len(Trim(Me.txtpv3.Text)) = 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp3", adDouble, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp3", adDouble, adParamInput, , CDec(Me.txtpv3.Text))

    End If

    If Len(Trim(Me.txtpv4.Text)) = 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp4", adDouble, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp4", adDouble, adParamInput, , CDec(Me.txtpv4.Text))

    End If
        
    If Len(Trim(Me.txtpv5.Text)) = 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp5", adDouble, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp5", adDouble, adParamInput, , CDec(Me.txtpv5.Text))

    End If

    If Len(Trim(Me.txtpv6.Text)) = 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp6", adDouble, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp6", adDouble, adParamInput, , CDec(Me.txtpv6.Text))

    End If
    
    If Len(Trim(Me.txtpvd1.Text)) = 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp11", adDouble, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp11", adDouble, adParamInput, , CDec(Me.txtpvd1.Text))

    End If
    
    If Len(Trim(Me.txtpvd2.Text)) = 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp22", adDouble, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp22", adDouble, adParamInput, , CDec(Me.txtpvd2.Text))

    End If
    
    If Len(Trim(Me.txtpvd3.Text)) = 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp33", adDouble, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp33", adDouble, adParamInput, , CDec(Me.txtpvd3.Text))

    End If

    If Len(Trim(Me.txtpvd4.Text)) = 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp44", adDouble, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp44", adDouble, adParamInput, , CDec(Me.txtpvd4.Text))

    End If

    If Len(Trim(Me.txtpv5.Text)) = 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp55", adDouble, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp55", adDouble, adParamInput, , CDec(Me.txtpv5.Text))

    End If

    If Len(Trim(Me.txtpvd6.Text)) = 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp66", adDouble, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pp66", adDouble, adParamInput, , CDec(Me.txtpv6.Text))

    End If

    'proporcion
    If Len(Trim(Me.txtProporcion.Text)) = 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@propocion", adDouble, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@propocion", adDouble, adParamInput, , CDbl(Me.txtProporcion.Text))

    End If
  
    If Me.chkSituacion.Value = 1 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@sit", adInteger, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@sit", adInteger, adParamInput, , 1)

    End If
  
    If Me.chkPri.Value = 1 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pri", adInteger, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pri", adInteger, adParamInput, , 1)

    End If
  
    If Me.chkPrioritario2.Value = 1 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pri2", adInteger, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@pri2", adInteger, adParamInput, , 1)

    End If
    
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@xmlCompform", adBSTR, adParamInput, 20000, IIf(Len(Trim(vxmlCompform)) = 0, vbNullString, Trim(vxmlCompform)))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@stock", adBoolean, adParamInput, , Me.chkStock.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PORCION", adBoolean, adParamInput, , Me.chkPorcion.Value)

    'agregado gts
    If Len(Trim(Me.txtPorcion.Text)) = 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@preporcion", adDouble, adParamInput, , 0)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PREPORCION", adDouble, adParamInput, , CDec(Me.txtPorcion.Text))

    End If
          
    Pub_ConnAdo.BeginTrans
    
    On Error GoTo gSAVE

    If VNuevo Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@flagstock", adChar, adParamInput, 1, cTipo)
        
        If Len(Trim(Me.txtCosto.Text)) = 0 Then
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@costo", adDouble, adParamInput, , 0)
        Else
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@costo", adDouble, adParamInput, , CDbl(Me.txtCosto.Text))

        End If

        'BOLSAS
        If Len(Trim(Me.txtBolsas.Text)) = 0 Then
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@BOLSAS", adInteger, adParamInput, , 0)
        Else
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@BOLSAS", adInteger, adParamInput, , Me.txtBolsas.Text)

        End If
        
        'COD ART BOLSAS
        If Len(Trim(Me.txtCodBolsa.Text)) = 0 Then
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODBOLSA", adInteger, adParamInput, , 0)
        Else
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODBOLSA", adInteger, adParamInput, , Me.txtCodBolsa.Text)

        End If


        'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@COSTO", adCurrency, adParamInput, , Me.txtCosto.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MaxCod", adInteger, adParamOutput, , 0)
        
        oCmdEjec.CommandText = "SpRegistraProducto"
        oCmdEjec.Execute
        VNuevo = False
        Me.lblCodigo.Caption = oCmdEjec.Parameters("@MaxCod").Value
        Me.lblCodigoF.Caption = oCmdEjec.Parameters("@MaxCod").Value
        
        Dim itemN As Object

        Set itemN = frmListaProd.lvData.ListItems.Add(, , Me.lblCodigo.Caption)
        itemN.SubItems(1) = Trim(Me.txtDescripcion.Text)
        itemN.SubItems(2) = Trim(Me.DatUM.Text)
        itemN.SubItems(3) = Me.dcboFam.Text
        itemN.SubItems(4) = Me.txtpv1.Text
        itemN.SubItems(5) = Me.txtpv2.Text
        itemN.SubItems(6) = Me.txtpv3.Text
        itemN.SubItems(7) = Me.txtpv4.Text
        itemN.SubItems(8) = Me.txtpv5.Text
        itemN.SubItems(9) = Me.txtpv6.Text
        itemN.SubItems(10) = 0
        itemN.SubItems(11) = Me.dcboFam.BoundText
        frmListaProd.lvData.ListItems(frmListaProd.lvData.ListItems.count).Selected = True   'julio 09/01/2011
        itemN.SubItems(12) = Me.dcboSubFam.BoundText
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codigo", adDouble, adParamInput, , CDbl(Me.lblCodigo.Caption))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adChar, adParamInput, 1, cTipo)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@COSTO", adCurrency, adParamInput, , Me.txtCosto.Text)

        'BOLSAS
        If Len(Trim(Me.txtBolsas.Text)) = 0 Then
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@BOLSAS", adInteger, adParamInput, , 0)
        Else
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@BOLSAS", adInteger, adParamInput, , Me.txtBolsas.Text)

        End If
        'COD ART BOLSAS
        If Len(Trim(Me.txtCodBolsa.Text)) = 0 Then
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODBOLSA", adInteger, adParamInput, , 0)
        Else
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODBOLSA", adInteger, adParamInput, , Me.txtCodBolsa.Text)

        End If
        oCmdEjec.CommandText = "SpModificaProducto"
        oCmdEjec.Execute

        With frmListaProd.lvData.SelectedItem
            .SubItems(1) = Trim(Me.txtDescripcion.Text)
            .SubItems(2) = Trim(Me.DatUM.Text)
            .SubItems(3) = Me.dcboFam.Text
            .SubItems(4) = Me.txtpv1.Text
            .SubItems(5) = Me.txtpv2.Text
            .SubItems(6) = Me.txtpv3.Text
            .SubItems(7) = Me.txtpv4.Text
            .SubItems(8) = Me.txtpv5.Text
            .SubItems(9) = Me.txtpv6.Text
            .SubItems(10) = 0
            .SubItems(11) = Me.dcboFam.BoundText
            .SubItems(12) = Me.dcboSubFam.BoundText

        End With

    End If

    'actualiza la foto
    If vGRABAi Then
    
        Dim varray As ADODB.Stream
        
        If Len(Trim(Me.ipic.Tag)) <> 0 Then
            Set varray = Imagen_Array(Me.ipic.Tag)
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SP_PRODUCTO_ACTUALIZAIMAGEN"
            '            oCmdEjec.Execute , Array(Me.lblCodigo, LK_CODCIA, varray.Read)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDARTICULO", adBigInt, adParamInput, , Me.lblCodigo.Caption)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IMG", adLongVarBinary, adParamInput, 1000000, varray.Read)
            oCmdEjec.Execute

            If varray.State = adStateOpen Then varray.Close
            If Not varray Is Nothing Then Set varray = Nothing
            
        Else
            
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SP_PRODUCTO_ACTUALIZAIMAGEN"
            oCmdEjec.Execute , Array(Me.lblCodigo, LK_CODCIA)

        End If
        
    End If

    vGraba = True
    Pub_ConnAdo.CommitTrans
    MsgBox "Datos Almacenados Correctamente", vbInformation, "Información"
    Unload Me
    Exit Sub

gSAVE:
    Pub_ConnAdo.RollbackTrans
    MsgBox Err.Description

End Sub

Private Sub cmdFam_Click()
frmAgreFamSubFam.VEsFam = True
frmAgreFamSubFam.lblfamilia.Visible = False
frmAgreFamSubFam.Show vbModal

If frmAgreFamSubFam.vAcepta Then
    
    oRsFam.AddNew
    oRsFam.Fields(0).Value = frmAgreFamSubFam.vCodFam
    oRsFam.Fields(1).Value = frmAgreFamSubFam.vDescripcion
    oRsFam.Update
    Me.dcboFam.BoundText = frmAgreFamSubFam.vCodFam
    
End If
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo Graba
If Len(Trim(Me.txtCodAlt.Text)) = 0 Then
    MsgBox "Debe ingresar el codigo alterno", vbInformation, "Error de Datos"
    Me.txtCodAlt.SetFocus
ElseIf Len(Trim(Me.txtDescripcion.Text)) = 0 Then
    MsgBox "Debe ingresar la Descripcion del Producto", vbInformation, "Error de Datos"
    Me.txtDescripcion.SetFocus
ElseIf Me.DatUM.BoundText = "" Then
    MsgBox "Debe ingresar la Unidad", vbInformation, "Error de Datos"
    Me.DatUM.SetFocus
ElseIf Me.dcboFam.BoundText = "" Then
    MsgBox "Debe elegir la familia", vbInformation, "Error de Datos"
    Me.dcboFam.SetFocus
ElseIf Me.dcboSubFam.BoundText = "" Then
    MsgBox "Debe elegir la Sub Familia", vbInformation, "ERror de Datos"
    Me.dcboSubFam.SetFocus
    'ElseIf Not IsNumeric(Me.txtProporcion.Text) Then
   ' MsgBox "El valor ingresado es incorrecto.", vbInformation, Pub_Titulo
    'Me.txtProporcion.SetFocus
    'Me.txtProporcion.SelLength = Len(Me.txtProporcion.Text)
Else
    Almacena

   
End If
Exit Sub
Graba:
    MsgBox Err.Description, vbCritical + vbOKOnly, "ERROR"
End Sub

Private Sub cmdImagen_Click()
'Abrimos el Commondialog con ShowOpen
Me.CommonDialog1.Filter = "jpg|*.jpg"
CommonDialog1.ShowOpen

'Si seleccionamos un archivo mostramos la ruta
If CommonDialog1.FileName <> "" Then

   vGRABAi = True

  
   Dim imagen As IPictureDisp
    'cargamos la imagen con LoadPicture
    Set imagen = LoadPicture(Me.CommonDialog1.FileName)
    
 alto = Round(Me.ScaleY(imagen.Height, vbHimetric, vbPixels))
      
    ancho = Round(Me.ScaleX(imagen.Width, vbHimetric, vbPixels))
      
    
    If alto > 60 Or ancho > 60 Then
    MsgBox "Las Dimenciones de la Imagen exceden las permitidas" & vbCrLf & "60px / 60px", vbInformation
    Else
     Me.ipic.Tag = Me.CommonDialog1.FileName
   Me.ipic.Picture = LoadPicture(Me.CommonDialog1.FileName)
    End If

Else
   'Si no mostramos un texto de advertencia de que no se seleccionó _
   ninguno, ya que FileName devuelve una cadena vacía
  Me.ipic.Tag = ""

End If
End Sub

Private Sub cmdImagenDEL_Click()
vGRABAi = True
Me.ipic.Tag = ""
Set Me.ipic.Picture = Nothing
End Sub

Private Sub cmdQuitarC_Click()
If Me.lvComposicion.ListItems.count > 0 Then
If Not Me.lvComposicion.SelectedItem Is Nothing Then
    Me.lvComposicion.ListItems.Remove Me.lvComposicion.SelectedItem.Index
End If
End If
End Sub

Private Sub cmdQuitarF_Click()
If Me.lvFormulacion.ListItems.count > 0 Then
If Not Me.lvFormulacion.SelectedItem Is Nothing Then
    Me.lvFormulacion.ListItems.Remove Me.lvFormulacion.SelectedItem.Index
End If
End If

End Sub

Private Sub cmdSubFam_Click()
If Me.dcboFam.BoundText = "" Then
    MsgBox "Debe elegir la Familia", vbCritical, "Error"
    Exit Sub
End If
frmAgreFamSubFam.VEsFam = False
frmAgreFamSubFam.lblfamilia.Tag = Me.dcboFam.BoundText
frmAgreFamSubFam.lblfamilia.Caption = Me.dcboFam.Text
'frmAgreFamSubFam.lblfamilia.Visible = False
frmAgreFamSubFam.Show vbModal


If frmAgreFamSubFam.vAcepta Then
    oRsSubFam.Filter = ""
    oRsSubFam.AddNew
    oRsSubFam.Fields(0).Value = frmAgreFamSubFam.vCodSubFam
    oRsSubFam.Fields(1).Value = Me.dcboFam.BoundText
    oRsSubFam.Fields(2).Value = frmAgreFamSubFam.vDescripcion
    oRsSubFam.Update
    
    oRsSubFam.Filter = "CodFam=" & Me.dcboFam.BoundText
    Dim orsSubFamTMP As New ADODB.Recordset
orsSubFamTMP.Fields.Append "codsubfam", adInteger
orsSubFamTMP.Fields.Append "Subfamilia", adVarChar, 60
orsSubFamTMP.CursorLocation = adUseClient
orsSubFamTMP.Open

Do While Not oRsSubFam.EOF
    orsSubFamTMP.AddNew
    orsSubFamTMP.Fields(0).Value = oRsSubFam.Fields(0).Value
    orsSubFamTMP.Fields(1).Value = oRsSubFam.Fields(2).Value
    orsSubFamTMP.Update
    oRsSubFam.MoveNext
Loop

    Set Me.dcboSubFam.RowSource = orsSubFamTMP
    Me.dcboSubFam.ListField = orsSubFamTMP.Fields(1).Name
    Me.dcboSubFam.BoundColumn = orsSubFamTMP.Fields(0).Name
    Me.dcboSubFam.BoundText = frmAgreFamSubFam.vCodSubFam
End If
End Sub

Private Sub ComCboTipoProd_Click()

    If Me.ComCboTipoProd.ListIndex = 0 Then 'Producto
        cTipo = "P"
    ElseIf Me.ComCboTipoProd.ListIndex = 1 Then 'Insumo
        cTipo = "M"
    ElseIf Me.ComCboTipoProd.ListIndex = 2 Then 'Combo
        cTipo = "C"
    ElseIf Me.ComCboTipoProd.ListIndex = 3 Then 'Materia Prima
        cTipo = "I"
    End If

    If Not vCarga Then
        Me.lvComposicion.ListItems.Clear
        Me.lvFormulacion.ListItems.Clear
    End If

End Sub

Private Sub DatUM_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.dcboFam.SetFocus
End Sub

Private Sub dcboFam_Change()
If Me.dcboFam.BoundText <> "" Then
oRsSubFam.Filter = "CodFam = " & Me.dcboFam.BoundText
If Not oRsSubFam.EOF Then
Set Me.dcboSubFam.RowSource = Nothing
Me.dcboSubFam.BoundText = ""

Dim orsSubFamTMP As New ADODB.Recordset
orsSubFamTMP.Fields.Append "codsubfam", adInteger
orsSubFamTMP.Fields.Append "Subfamilia", adVarChar, 60
orsSubFamTMP.CursorLocation = adUseClient
orsSubFamTMP.Open

Do While Not oRsSubFam.EOF
    orsSubFamTMP.AddNew
    orsSubFamTMP.Fields(0).Value = oRsSubFam.Fields(0).Value
    orsSubFamTMP.Fields(1).Value = oRsSubFam.Fields(2).Value
    orsSubFamTMP.Update
    oRsSubFam.MoveNext
Loop
'Set orsa = oRsCombos
Set Me.dcboSubFam.RowSource = orsSubFamTMP
Me.dcboSubFam.BoundColumn = orsSubFamTMP.Fields(0).Name
Me.dcboSubFam.ListField = orsSubFamTMP.Fields(1).Name
Else
Set Me.dcboSubFam.RowSource = Nothing
Me.dcboSubFam.BoundText = ""
End If
End If
End Sub

Private Sub dcboFam_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.dcboSubFam.SetFocus
End Sub

Private Sub dcboSubFam_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.txtMin.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    CreaRS
    vGraba = False
    vGRABAi = False
    vCarga = True
    ConfiguraComposicion
    ConfiguraFormulacion
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpCargarInfoProductos"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    Set oRsCombos = oCmdEjec.Execute

    Do While Not oRsCombos.EOF
        oRsFam.AddNew
        oRsFam.Fields(0).Value = oRsCombos!NUMERO
        oRsFam.Fields(1).Value = Trim(oRsCombos!Familia)
        oRsFam.Update
        oRsCombos.MoveNext
    Loop

    Set Me.dcboFam.RowSource = oRsFam
    Me.dcboFam.BoundColumn = oRsFam.Fields(0).Name
    Me.dcboFam.ListField = oRsFam.Fields(1).Name
    'llenando sub familias

    Set oRsCombos = oRsCombos.NextRecordset

    Do While Not oRsCombos.EOF
        oRsSubFam.AddNew
        oRsSubFam.Fields(0).Value = oRsCombos!NUMERO 'codigo de sub familia
        oRsSubFam.Fields(1).Value = oRsCombos!codfam 'codigo de familia
        oRsSubFam.Fields(2).Value = Trim(oRsCombos!Familia) 'descripcion de la sub familia
        oRsSubFam.Update
        oRsCombos.MoveNext
    Loop

    Set oRSum = oRsCombos.NextRecordset

    Set Me.DatUM.RowSource = oRSum
    Me.DatUM.BoundColumn = oRSum.Fields(0).Name
    Me.DatUM.ListField = oRSum.Fields(1).Name
    '
    'If Not ORS.EOF Then
    'Set Me.dcboSubFam.RowSource = ORS
    'Me.dcboSubFam.BoundColumn = ORS.Fields(0).Name
    'Me.dcboSubFam.ListField = ORS.Fields(2).Name
    'End If

    If VNuevo Then
        Me.chkSituacion.Value = 1
        Me.chkStock.Value = 0
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPPRODUCTO_ULTIMONUMERO"

        Dim orsN As ADODB.Recordset

        Set orsN = oCmdEjec.Execute(, Array(LK_CODCIA))

        Do While Not orsN.EOF
            Me.txtCodAlt.Text = orsN!NUMERO
            orsN.MoveNext
        Loop

    Else
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SpCargarInfoAuxProductos"
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Codigo", adInteger, adParamInput, , vCodigo)
        Set oRS = oCmdEjec.Execute
        Me.txtpv1.Text = oRS!pv1
        Me.txtpv2.Text = oRS!pv2
        Me.txtpv3.Text = oRS!pv3
        Me.txtpv4.Text = oRS!pv4
        Me.txtpv5.Text = oRS!pv5
        Me.txtpv6.Text = oRS!pv6
        Me.txtpvd1.Text = oRS!pvd1
        Me.txtpvd2.Text = oRS!pvd2
        Me.txtpvd3.Text = oRS!pvd3
        Me.txtpvd4.Text = oRS!pvd4
        Me.txtpvd5.Text = oRS!pvd5
        Me.txtpvd6.Text = oRS!pvd6
        Me.txtCodAlt.Text = oRS!codalt
        Me.lblCodigo.Caption = vCodigo
        Me.txtMin.Text = oRS!minimo
        Me.txtMax.Text = oRS!maximo
        Me.lblCodigoF.Caption = vCodigo
        Me.chkPorcion.Value = IIf(CBool(oRS!Porcion), 1, 0)
        Me.txtPorcion.Text = oRS!preporcion
        Me.txtCosto.Text = oRS!costo
        Me.txtBolsas.Text = oRS!BOLSAS

        If oRS!SIT = 0 Then
            Me.chkSituacion.Value = 1
        Else
            Me.chkSituacion.Value = 0
        End If

        If oRS!pri = 1 Then
            Me.chkPri.Value = 0
        Else
            Me.chkPri.Value = 1
    
        End If

        If oRS!pri2 = 1 Then
            Me.chkPrioritario2.Value = 0
        Else
            Me.chkPrioritario2.Value = 1
        End If

        Me.chkStock.Value = IIf(oRS!stock, 1, 0)
        Me.txtProporcion.Text = oRS!proporcion
    
        Dim i As Integer

        i = 0

        Dim Vid As Integer

        Do While Not oRSum.EOF
            oRSum.Filter = "DENO= '" & Trim(oRS!UM) & "'"

            If oRSum.RecordCount = 1 Then
                Vid = oRSum!IDE

                Exit Do

            End If
        
            i = i + 1
        Loop
    
        Me.DatUM.BoundText = Vid
    
        'MOSTRANDO LA IMAGEN EN CASO TUVIERA
        Dim sIMG As ADODB.Stream

        ' Nuevo objeto Stream para poder leer el campo de imagen
        Set sIMG = New ADODB.Stream
      
        ' Especifica el tipo de datos ( binario )
        sIMG.Type = adTypeBinary
        sIMG.Open
      
        ' verifica con la función IsNull que el campo no tenga _
          un valor Nulo ya que si no da error, en ese caso sale de la función

        If Not IsNull(oRS!Datoimagen) Then
            ' Graba los datos en el objeto stream
            sIMG.Write oRS.Fields!Datoimagen
        
            ' este método graba un  archivo temporal  en disco _
              ( en el app.path que luego se elimina )
            sIMG.SaveToFile App.Path & "\temp.jpg", adSaveCreateOverWrite
        
            ' Retorna la imagen a la función
            'Set Me.Picture1.Picture = LoadPicture(App.Path & "\temp.jpg")
            Me.ipic.Picture = LoadPicture(App.Path & "\temp.jpg")
        
            ' Elimina el archivo temporal
            Kill App.Path & "\temp.jpg"
        
            If sIMG.State = adStateOpen Then sIMG.Close
            If Not sIMG Is Nothing Then Set sIMG = Nothing
        End If
    
        '===================================
        Set oRS = oRS.NextRecordset
        Dim vTotalCosto As Double
        vTotalCosto = 0
        If Not oRS.EOF Then

            Dim itemX As Object
        
            Do While Not oRS.EOF

                If frmListaProd.cboTipoProd.ListIndex = 2 Then
                    Set itemX = Me.lvComposicion.ListItems.Add(, , oRS!Codigo)
            
                ElseIf frmListaProd.cboTipoProd.ListIndex = 0 Then
                    Set itemX = Me.lvFormulacion.ListItems.Add(, , oRS!Codigo)
                ElseIf frmListaProd.cboTipoProd.ListIndex = 1 Then 'NUEVO
                    Set itemX = Me.lvFormulacion.ListItems.Add(, , oRS!Codigo) 'NUEVO
                End If

                itemX.SubItems(1) = Trim(oRS!DESCRIPCION)
                itemX.SubItems(2) = Trim(oRS!UNIDAD)
                itemX.SubItems(3) = oRS!Cantidad
                itemX.SubItems(4) = oRS!costounit
                itemX.SubItems(5) = oRS!costo
                vTotalCosto = vTotalCosto + oRS!costo
                oRS.MoveNext
            Loop

        End If
        Me.lblTot.Caption = vTotalCosto
        '--------------------- AGREGADO GTS 01/05/2016--------------
        If Me.lblTot <> 0 Then
        Me.txtCosto.Text = vTotalCosto
        End If
        '------------------- AGREGADO GTS 01/05/2016--------------
        If val(Me.txtpv1.Text) <> 0 Then
        Me.lblPorcentaje.Caption = Round((CDec(Me.txtCosto.Text) * 100) / CDec(Me.txtpv1.Text), 2)
        Else
        Me.lblPorcentaje.Caption = "0"
        End If

        '====AGREGADO GTS PARA EL MOCHICA PARA Q NO PUEDAN MODIFICAR PRECIOS USUARIOS SIN PERMISO=======
        For fila = 1 To lk_OTROS_Count

            If val(lk_OTROS(fila)) = 6 Then ' bloque de precios en mastros de articulos
                loc_flag_bloq = "A"
            End If

        Next fila

        If loc_flag_bloq = "A" Then
            txtpv1.Locked = True
            txtpv2.Locked = True
            txtpv3.Locked = True
            txtpv4.Locked = True
            txtpv5.Locked = True
            txtpv6.Locked = True
            txtpvd1.Locked = True
            txtpvd2.Locked = True
            txtpvd3.Locked = True
            txtpvd4.Locked = True
            txtpvd5.Locked = True
            txtpvd6.Locked = True
        End If

        '====AGREGADO GTS PARA EL MOCHICA PARA Q NO PUEDAN MODIFICAR PRECIOS USUARIOS SIN PERMISO=======
    End If

End Sub



Private Sub txtBolsas_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtCodAlt_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.txtDescripcion.SetFocus
End Sub

Private Sub txtCosto_Change()
If val(Me.txtpv1.Text) <> 0 Then
Me.lblPorcentaje.Caption = Round((CDec(Me.txtCosto.Text) * 100) / CDec(Me.txtpv1.Text), 2)
Else
Me.lblPorcentaje.Caption = "0"
End If
End Sub

Private Sub txtDescripcion_Change()
Me.lblDescripcionF.Caption = Me.txtDescripcion.Text
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.DatUM.SetFocus
'KeyAscii = Mayusculas(KeyAscii)
End Sub

Private Sub txtMax_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.txtpv1.SetFocus
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.txtMax.SetFocus
End Sub

Private Sub txtProporcion_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
'KeyAscii = SoloNumeros(KeyAscii)
End Sub

Private Sub txtProporcion_LostFocus()
If Len(Trim(Me.txtProporcion.Text)) = 0 Then Me.txtProporcion.Text = 0
End Sub

Private Sub txtpv1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.txtpv2.SetFocus
End Sub

Private Sub txtpv2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.txtpv3.SetFocus
End Sub

Private Sub txtpv3_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.txtpv4.SetFocus
End Sub

Private Sub txtpv4_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.txtpv5.SetFocus
End Sub

Private Sub txtpv5_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.txtpv6.SetFocus
End Sub



