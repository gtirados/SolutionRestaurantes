VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form FrmImp2 
   Caption         =   "Reportes"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8445
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1695
      Left            =   240
      TabIndex        =   60
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
      Begin VB.TextBox txtserie 
         Height          =   375
         Left            =   960
         TabIndex        =   63
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox fbtxt 
         Height          =   375
         Left            =   1080
         TabIndex        =   62
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox max 
         Height          =   375
         Left            =   1080
         TabIndex        =   61
         Text            =   "65"
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblmoneda1 
         Caption         =   "Serie:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblmoneda1 
         Caption         =   "F/B"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   65
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblmoneda1 
         Caption         =   "Lineas X pag."
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   64
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Data transfer 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Admin\Standar\TRANSFER.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "FACART"
      Top             =   3720
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Frame fzonas 
      Caption         =   "Zonas"
      Height          =   2535
      Left            =   5520
      TabIndex        =   58
      Top             =   720
      Visible         =   0   'False
      Width           =   2895
      Begin VB.ListBox solozonas 
         Height          =   2205
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   59
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CheckBox chestands 
      Caption         =   "Multiplicar Monto por Nº de Stands"
      Height          =   255
      Left            =   4200
      TabIndex        =   57
      Top             =   2640
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   2280
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CheckBox cheasiento 
      Caption         =   "Pasar a Contabilidad"
      Height          =   255
      Left            =   4200
      TabIndex        =   55
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   615
      Left            =   1800
      TabIndex        =   20
      Top             =   -120
      Width           =   4455
      Begin VB.Label lblreporte 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   3945
      End
   End
   Begin VB.Frame fradeudas 
      Height          =   1215
      Left            =   5520
      TabIndex        =   31
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CheckBox Check1 
         Caption         =   "Fechas de Nuevo Vcto."
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   120
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox Cheop2 
         Caption         =   "Deudas por Cobrar Vencidas"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   600
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Cheop1 
         Caption         =   "Deudas por Cobrar del Dia"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   360
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Cheop3 
         Caption         =   "Deudas por Cobrar por Vencer"
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
      _Version        =   327680
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
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame fcontab 
      Height          =   2535
      Left            =   0
      TabIndex        =   42
      Top             =   840
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CheckBox chenivel 
         Caption         =   "Nivel 6"
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   49
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox chenivel 
         Caption         =   "Nivel 5"
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   48
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chenivel 
         Caption         =   "Nivel 4"
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   47
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chenivel 
         Caption         =   "Nivel 3"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox chenivel 
         Caption         =   "Nivel 2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chenivel 
         Caption         =   "Nivel 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblcontab 
         Caption         =   "Seleccione los Niveles para impresión"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame fracli 
      Height          =   2535
      Left            =   3480
      TabIndex        =   34
      Top             =   840
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
         Caption         =   "2.- "
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   41
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblcli 
         Caption         =   "1.- "
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   40
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblcli 
         Caption         =   "Ingrese 2 Rangos :"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   39
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CheckBox Chesup 
      Caption         =   "Suprimir las Columnas con 0"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   840
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame frazonas 
      Height          =   2535
      Left            =   0
      TabIndex        =   32
      Top             =   840
      Visible         =   0   'False
      Width           =   8415
      Begin VB.OptionButton opzonas 
         Caption         =   "Zonas"
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
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   2175
      End
      Begin VB.OptionButton opzonas 
         Caption         =   "Distrito"
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
         Height          =   1950
         Left            =   3720
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblzonas 
         AutoSize        =   -1  'True
         Caption         =   "Zonas :"
         Height          =   195
         Left            =   2040
         TabIndex        =   33
         Top             =   600
         Width           =   1380
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
      Height          =   2220
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CheckBox chestock 
      Caption         =   "Incluir Valorado"
      Height          =   255
      Left            =   6240
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame FRASTOCK 
      Height          =   2415
      Left            =   3120
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   5175
      Begin VB.ComboBox CmbCalidad 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1080
         Width           =   2055
      End
      Begin VB.ListBox subfami 
         Appearance      =   0  'Flat
         Height          =   705
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   1560
         Width           =   2775
      End
      Begin VB.ListBox fami 
         Appearance      =   0  'Flat
         Height          =   930
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblcalidad 
         AutoSize        =   -1  'True
         Caption         =   "Calidad"
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
         Caption         =   "Sub Familia :"
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
      Height          =   2220
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox PROV 
      Height          =   2085
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   960
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
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
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
      _Version        =   327680
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
      _Version        =   327680
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
      Caption         =   "Fecha :"
      Height          =   195
      Left            =   1680
      TabIndex        =   52
      Top             =   600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblmoneda 
      Caption         =   "Moneda :"
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
      Caption         =   "Proveedores"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblProceso 
      Alignment       =   2  'Center
      Caption         =   "Procesando ..."
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
Dim VAR
Dim nom
If Trim((listVen.Text)) = "" Then Exit Sub
lblProceso.Visible = True
DoEvents
nom = Trim(Left(listVen.Text, 3))
VAR = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\OFFICE\PLANVEN" & nom & ".XLS"

If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
End If
On Error GoTo no_existe
xl.Workbooks.Open VAR, 0, False, 4
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
'lblmoneda.Visible = True
'cmdMoneda.Visible = True
If Trim(par_llave!PAR_DEFAULT_FAC) = "D" Then
'  cmdMoneda.ListIndex = 1
Else
'  cmdMoneda.ListIndex = 0
End If

If Wfile = "IMP_STOCK" Then
 chestock.Visible = True
 lblstock.Visible = True
 pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CP = 'P'  AND CLI_CODCIA = ? ORDER BY CLI_NOMBRE"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
 PS_REP01(0) = LK_CODCIA
 llave_rep01.Requery
 Do Until llave_rep01.EOF
     PROV.AddItem llave_rep01!cli_nombre & String(25, " ") & llave_rep01!cli_codclie
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
If Wfile = "IMP_PLANILLA" Or Wfile = "IMP_CONSUPLAN" Or Wfile = "IMP_COMISION" Or Wfile = "IMP_COMI_NETAS" Then
 Dim codi As String * 5
 lblstock.Caption = "Vendedor : "
 lblstock.Visible = True
 pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = ? ORDER BY VEM_CODVEN"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 PS_REP01(0) = 0
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
  txtCampo2.Mask = "##/##/####"
  txtCampo2.Visible = True
 End If
 listVen.TabIndex = 0
End If

If Wfile = "PARTE_DIARIO" Or Wfile = "VENTA_X_VEND" Or Wfile = "VEN_NEGOCIOS" Or Wfile = "VEN_NEGOCIOS_CLI" Or Wfile = "HER_NEGOCIOS" Or Wfile = "HER_VENDEDOR" Or Wfile = "HER_REGISTRO" Or Wfile = "REGCOMP" Or Wfile = "RESUMEN_TRANSA" Or Wfile = "RESU_VENTA_DIA" Or Wfile = "TRANSFERENCIA" Then
 If Wfile = "HER_REGISTRO" Or Wfile = "REGCOMP" Or Wfile = "RESU_VENTA_DIA" Then
   If par_llave!PAR_CONTABILIDAD = "A" Then
    txtCampo1.Text = Format(cop_llave!cop_fecha_proceso, "dd/mm/yyyy")
    txtCampo2.Text = Format(cop_llave!COP_FECHA_PROCESO2, "dd/mm/yyyy")
    cheasiento.Visible = True
   Else
    txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    txtCampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    max.Text = 65
    Frame2.Visible = True
   End If
 Else
   txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 End If

 lblcampo1.Caption = "Fecha de Inicial : "
 lblcampo1.Visible = True
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
 lblcampo2.Caption = "Fecha de Final: "
 lblcampo2.Visible = True
 txtCampo2.Mask = "##/##/####"
 txtCampo2.Visible = True
 If Wfile = "VENTA_X_VEND" Or Wfile = "HER_VENDEDOR" Or Wfile = "VEN_NEGOCIOS" Or Wfile = "VEN_NEGOCIOS_CLI" Then
    lblstock.Caption = "Vendedor : "
    lblstock.Visible = True
    pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = ? ORDER BY VEM_CODVEN"
    Set PS_REP01 = CN.CreateQuery("", pub_cadena)
    PS_REP01(0) = 0
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
 If Wfile = "PARTE_DIARIO" Then
    cheasiento.Visible = False
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
    
    lblcampo1.Caption = "Fecha de Inicial : "
    lblcampo1.Visible = True
    txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    txtCampo1.Mask = "##/##/####"
    txtCampo1.Visible = True
    lblcampo2.Caption = "Fecha de Final: "
    lblcampo2.Visible = True
    txtCampo2.Mask = "##/##/####"
    txtCampo2.Visible = True
ElseIf Wfile = "COMISIONES" Then
    pantalla.TabIndex = 0
    lblcampo1.Caption = "Fecha de Inicial : "
    lblcampo1.Visible = True
    txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    txtCampo1.Mask = "##/##/####"
    txtCampo1.Visible = True
    lblcampo2.Caption = "Fecha de Final: "
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
    lblcampo1.Caption = "Fecha de Inicial : "
    lblcampo1.Visible = True
    txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    txtCampo1.Mask = "##/##/####"
    txtCampo1.Visible = True
    lblcampo2.Caption = "Fecha de Final: "
    lblcampo2.Visible = True
    txtCampo2.Mask = "##/##/####"
    txtCampo2.Visible = True
End If
If Wfile = "REPO_CAJA_GEN" Then
' lblmoneda.Visible = True
 cmdMoneda.Visible = True
 cmdMoneda.ListIndex = 0
 cmdMoneda.TabIndex = 0
 lblfecha.Visible = True
 txtfecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtfecha.Mask = "##/##/####"
 txtfecha.Visible = True
End If
If Wfile = "REPO_CAJA_DET" Or Wfile = "REPO_CAJA_DET2" Then
 lblfecha.Visible = True
 txtfecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtfecha.Mask = "##/##/####"
 txtfecha.Visible = True
End If
If Wfile = "REPO_CAJA" Then
' lblmoneda.Visible = True
 cmdMoneda.Visible = True
 cmdMoneda.ListIndex = 0
 cmdMoneda.TabIndex = 0
 lblfecha.Visible = True
 txtfecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtfecha.Mask = "##/##/####"
 txtfecha.Visible = True
End If
If Wfile = "ULTIMAS_VENTAS" Then
 PUB_CODCIA = "00"
 LLENADOS zonas, 35
 frazonas.Visible = True
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
        If tab_mayor!TAB_NOMLARGO <> "TE" And tab_mayor!TAB_NOMLARGO <> "CT" Then
          Combo1.AddItem tab_mayor!TAB_NOMLARGO & "=" & tab_mayor!tab_nomcorto & String(30, " ") & Left(tab_mayor!TAB_contable2, 1)
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
    lblmoneda.Visible = False
    cmdMoneda.Visible = False
End If
If Wfile = "FECHA_STOCK" Then
  PUB_CODCIA = LK_CODCIA
  LLENADOS listVen, 122
  listVen.Visible = True
  lblcampo1.Caption = "Fecha para el Stock : "
  lblcampo1.Visible = True
  txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
  txtCampo1.Mask = "##/##/####"
  txtCampo1.Visible = True
End If
If Wfile = "INTERES" Or Wfile = "PLAZA_GRAU" Then
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
        'If tab_mayor!TAB_NOMLARGO <> "TE" And tab_mayor!TAB_NOMLARGO <> "CT" Then
          Combo1.AddItem Trim(tab_mayor!TAB_NOMLARGO) & "=" & tab_mayor!tab_nomcorto & String(30, " ") & Left(tab_mayor!TAB_contable2, 1)
        'End If
        tab_mayor.MoveNext
    Loop
    Combo1.ListIndex = 0
End If

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

   ElseIf Wfile = "IMP_CONSUPLAN" Then

   ElseIf Wfile = "IMP_COMISION" Then

   End If
   
ElseIf Wfile = "PARTE_DIARIO" Or Wfile = "VENTA_X_VEND" Or Wfile = "VEN_NEGOCIOS" Or Wfile = "VEN_NEGOCIOS_CLI" Or Wfile = "HER_NEGOCIOS" Or Wfile = "HER_VENDEDOR" Or Wfile = "HER_REGISTRO" Or Wfile = "REGCOMP" Or Wfile = "RESU_VENTA_DIA" Or Wfile = "TRANSFERENCIA" Then
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

   ElseIf Wfile = "HER_VENDEDOR" Then

   ElseIf Wfile = "HER_REGISTRO" Then
    
   ElseIf Wfile = "TRANSFERENCIA" Then
    Call Transferencia
    
    
   End If
End If
If Wfile = "DEDVENC" Then

ElseIf Wfile = "BALANCE" Then
 
ElseIf Wfile = "REPO_CAJA_GEN" Then
 
ElseIf Wfile = "PLAZAG" Then
 
ElseIf Wfile = "REPO_CAJA" Then
  
ElseIf Wfile = "SALDO_CAR" Then
 
ElseIf Wfile = "SALDO_CAR_IT" Then
 
ElseIf Wfile = "SALDO_CAR_TE" Then
 
ElseIf Wfile = "SALDO_TOT_TE" Then
 
ElseIf Wfile = "SALDO_CAR_IT" Then
 
ElseIf Wfile = "PRO_CO" Then
 
ElseIf Wfile = "VENTAS_DEUDA" Then
 
ElseIf Wfile = "ULTIMAS_VENTAS" Then
 
ElseIf Wfile = "REPO_CAJA_DET" Then
 
ElseIf Wfile = "REPO_CAJA_DET2" Then
 
ElseIf Wfile = "CARGA_CARTERA" Then

ElseIf Wfile = "FECHA_STOCK" Then
 
ElseIf Wfile = "IMP_COMI_NETAS" Then
 
End If

Exit Sub
SALE:
ProgBar.Visible = False
lblProceso.Visible = False
pantalla.Enabled = True
cerrar.Enabled = True
MsgBox Err.Description + "Intente Nuevamente.", 48, Pub_Titulo
End Sub

Private Sub solozonas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If pantalla.Enabled Then pantalla.SetFocus
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
        cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!tab_numtab
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
        cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!tab_numtab
        tab_mayor.MoveNext
    Loop
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
Private Sub zonas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If pantalla.Enabled Then pantalla.SetFocus
End If
End Sub
Private Sub Transferencia()
Dim RR
Dim PSFAR As rdoQuery
Dim far_r  As rdoResultset
Dim CONTADOR As Integer
Dim ArchOrigen, ArchDestino As String
Dim pub_mensaje, estilo, respuesta
Dim PSfar_menorx As rdoQuery
Dim far_menorx As rdoResultset


Dim WLUGAR As String
Dim WZONA As String
Dim WSUBZONA As String

'On Error GoTo SALE

pub_cadena = "SELECT * FROM facart WHERE FAR_CODCIA= ? AND FAR_FECHA>= ? AND FAR_FECHA<= ? AND FAR_TIPMOV = 10  AND FAR_ESTADO<>'E'  ORDER BY FAR_FECHA, FAR_NUMOPER"
Set PSfar_menorx = CN.CreateQuery("", pub_cadena)
Set far_menorx = PSfar_menorx.OpenResultset(rdOpenKeyset, rdConcurValues)

PSfar_menorx.rdoParameters(0) = LK_CODCIA
PSfar_menorx.rdoParameters(1) = txtCampo1.Text
PSfar_menorx.rdoParameters(2) = txtCampo2.Text
far_menorx.Requery

Screen.MousePointer = 11

DoEvents
transfer.RecordSource = "SELECT * FROM FACART"

transfer.Refresh
Do Until transfer.Recordset.EOF
   transfer.Recordset.Edit
   transfer.Recordset.Delete
   transfer.Recordset.MoveNext
Loop

SALTA:
FrmImp2.lblProceso.Visible = True
FrmImp2.ProgBar.Visible = True
FrmImp2.lblProceso.Caption = "Transferencia de Ventas"
FrmImp2.ProgBar.Min = 0
FrmImp2.ProgBar.Value = 0
FrmImp2.ProgBar.max = far_menorx.RowCount
CONTADOR = 0
If Not far_menorx.EOF Then
   PUB_FECHA = far_menorx!FAR_fecha
End If

CN.Execute "Begin Transaction", rdExecDirect
      
      Do Until far_menorx.EOF
         transfer.Recordset.AddNew
         transfer.Recordset!FAR_CODCIA = far_menorx!FAR_CODCIA
         transfer.Recordset!far_fbg = far_menorx!far_fbg
         transfer.Recordset!far_numser = far_menorx!far_numser
         transfer.Recordset!far_numfac = far_menorx!far_numfac
         transfer.Recordset!far_numsec = far_menorx!far_numsec
         transfer.Recordset!FAR_fecha = far_menorx!FAR_fecha
         transfer.Recordset!far_codclie = 0
         transfer.Recordset!far_codart = 0
         transfer.Recordset!far_PRECIO = far_menorx!far_PRECIO
         transfer.Recordset!FAR_pordescto1 = far_menorx!FAR_pordescto1
         transfer.Recordset!FAR_IMPTO = far_menorx!FAR_IMPTO
         transfer.Recordset!FAR_SUBTOTAL = far_menorx!FAR_SUBTOTAL
         transfer.Recordset!FAR_codven = far_menorx!FAR_codven
         transfer.Recordset!FAR_CANTIDAD = far_menorx!FAR_CANTIDAD
         transfer.Recordset!FAR_SIGNO_CAR = far_menorx!FAR_SIGNO_CAR
         transfer.Recordset!FAR_MONEDA = far_menorx!FAR_MONEDA
         transfer.Recordset!far_SERGUIA = far_menorx!far_SERGUIA
         transfer.Recordset!far_numGUIA = far_menorx!far_numGUIA
         transfer.Recordset!FAR_DIAS = far_menorx!FAR_DIAS
         transfer.Recordset!far_DESCRI = far_menorx!far_DESCRI
         transfer.Recordset!far_PESO = far_menorx!far_PESO
         transfer.Recordset!FAR_SIGNO_ARM = far_menorx!FAR_SIGNO_ARM
         transfer.Recordset!FAR_gastos = far_menorx!FAR_gastos
         transfer.Recordset!far_numser_c = far_menorx!far_numser_c
         transfer.Recordset!far_numfac_c = far_menorx!far_numfac_c
         transfer.Recordset!FAR_CP = far_menorx!FAR_CP
         transfer.Recordset!far_cod_sunat = far_menorx!far_cod_sunat
         transfer.Recordset!FAR_TIPDOC = far_menorx!FAR_TIPDOC
         transfer.Recordset!FAR_BRUTO = far_menorx!FAR_BRUTO
         
         SQ_OPER = 1
         pu_codclie = far_menorx!far_codclie
         pu_cp = "C"
         pu_codcia = LK_CODCIA
         LEER_CLI_LLAVE
         PUB_KEY = far_menorx!far_codart
         LEER_ART_LLAVE
         transfer.Recordset!FAR_NOMART = Trim(art_LLAVE!ART_NOMBRE)
         transfer.Recordset!FAR_NOMCLI = cli_llave!cli_nombre
         transfer.Recordset!far_ALTERNO = art_LLAVE!ART_ALTERNO
         If Trim(cli_llave!cli_RUC_ESPOSO) <> "" Then
            transfer.Recordset!FAR_RUC = Trim(cli_llave!cli_RUC_ESPOSO)
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
        WLUGAR = Trim(tab_llave!TAB_NOMLARGO)
        End If
        
        PUB_NUMTAB = cli_llave!cli_TRAB_ZONA
        PUB_TIPREG = 20
        LEER_TAB_LLAVE
        WZONA = ""
        If Not tab_llave.EOF Then
        WZONA = Trim(tab_llave!TAB_NOMLARGO)
        End If
        PUB_NUMTAB = cli_llave!cli_TRAB_SUBZONA
        PUB_TIPREG = 35
        LEER_TAB_LLAVE
        WSUBZONA = ""
        If Not tab_llave.EOF Then
        WSUBZONA = Trim(tab_llave!TAB_NOMLARGO)
        End If
        If Val(cli_llave!CLI_TRAB_NUM) <> 0 Then
           transfer.Recordset!FAR_DIREC = Trim(WLUGAR) + " " + Trim(cli_llave!CLI_TRAB_DIREC) + " # " + Trim(cli_llave!CLI_TRAB_NUM) & "  " & WZONA & "  " & WSUBZONA
        Else
           transfer.Recordset!FAR_DIREC = Trim(WLUGAR) + " " + Trim(cli_llave!CLI_TRAB_DIREC) & "  " & WZONA & "  " & WSUBZONA
        End If
         transfer.Recordset!FAR_DIREC = Left(Trim(transfer.Recordset!FAR_DIREC), 60)
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
  MsgBox "Proceso Terminado .. Pero EL proceso de Enpaquetar hacerlo Manual"
  Screen.MousePointer = 0
  Unload FrmImp2

  
End Sub

