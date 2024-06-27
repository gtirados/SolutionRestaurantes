VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Frmformula 
   Caption         =   "Formulación"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   1620
   ClientWidth     =   9480
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9480
   Begin VB.Frame Frame2 
      Caption         =   "Articulos para Procesos"
      Height          =   4215
      Left            =   6600
      TabIndex        =   2
      Top             =   120
      Width           =   2895
      Begin VB.ListBox LISARTI 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3840
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton Command1 
         Caption         =   "&Calcular"
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
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtPeso 
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txttotal 
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
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Total Peso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Total  ( S/.)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
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
      Left            =   3840
      TabIndex        =   7
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Lista de Articulos para Procesos"
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   6615
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Frmformula.frx":0000
         Height          =   2055
         Left            =   0
         OleObjectBlob   =   "Frmformula.frx":0011
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.Label momento 
         Alignment       =   2  'Center
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
         Left            =   1800
         TabIndex        =   6
         Top             =   840
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Procesos :"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2895
      Begin VB.CommandButton Command3 
         Caption         =   "Crear &Procesos"
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
         Left            =   600
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox LISPROCESOS 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
   Begin MSRDC.MSRDC MSRDC1 
      Height          =   330
      Left            =   1080
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   327680
      Options         =   0
      CursorDriver    =   0
      BOFAction       =   0
      EOFAction       =   0
      RecordsetType   =   1
      LockType        =   3
      QueryType       =   0
      Prompt          =   3
      Appearance      =   1
      QueryTimeout    =   30
      RowsetSize      =   100
      LoginTimeout    =   15
      KeysetSize      =   0
      MaxRows         =   0
      ErrorThreshold  =   -1
      BatchSize       =   15
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      ReadOnly        =   0   'False
      Appearance      =   -1  'True
      DataSourceName  =   "DSN_DATOS"
      RecordSource    =   ""
      UserName        =   "ABEL"
      Password        =   "ABEL"
      Connect         =   ""
      LogMessages     =   ""
      Caption         =   "MSRDC1"
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
   Begin VB.Label Label1 
      Caption         =   "Editar Procesos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Frmformula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cod As String
Dim kPASE As Boolean


Private Sub Command1_Click()
If kPASE Then
  SUMTOTAL
  Command1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Unload Frmformula
End Sub

Private Sub Command3_Click()
Load FrmProcesos
FrmProcesos.Show 1
'LISTA_PROCESOS
End Sub


Private Sub DBGrid1_AfterDelete()
''SUMTOTAL
End Sub

Private Sub DBGrid1_BeforeUpdate(Cancel As Integer)
Dim wcantidad, wprecio, t
If Not IsNumeric(DBGrid1.Columns(1)) Then
  Cancel = True
  Exit Sub
End If
wcantidad = Val(DBGrid1.Columns(1))
wprecio = Val(DBGrid1.Columns(2))
DBGrid1.Columns(3) = wcantidad * wprecio
''SUMTOTAL
End Sub

Private Sub Form_Load()
Dim cade
kPASE = False
LISTA_PROCESOS
LISTA_ARTI

End Sub

Public Sub LISTA_ARTI()
Dim ca, WF
Dim loc_arti As rdoResultset
cad = "SELECT * FROM ARTI WHERE ART_KEY > 0 AND ART_CODCIA = '" & LK_CODCIA & "' ORDER BY ART_NOMBRE"
Set loc_arti = CN.OpenResultset(cad, rdOpenKeyset, rdConcurValues)
loc_arti.Requery

Do Until loc_arti.EOF
  LISARTI.AddItem loc_arti!ART_NOMBRE & String(30, " ") & loc_arti!ART_KEY
  loc_arti.MoveNext
Loop

End Sub
Public Sub LISTA_PROCESOS()
SQ_OPER = 2
PUB_CODCIA = LK_CODCIA
PUB_TIPREG = 888
LEER_TAB_LLAVE
LISPROCESOS.Clear
Do Until tab_mayor.EOF
  LISPROCESOS.AddItem tab_mayor!tab_nomlargo & String(30, " ") & tab_mayor!tab_numtab
tab_mayor.MoveNext
Loop


End Sub

Private Sub LISARTI_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
  Exit Sub
End If
Dim WSEC
Static WORDEN
Dim wCODARTI
wCODARTI = Val(Trim(Right(LISARTI.text, 8)))
If wCODARTI = 0 Or kPASE = False Then
  MsgBox "Debe Seleccionar Proceso ..", 48, Pub_Titulo
  LISPROCESOS.SetFocus
  Exit Sub
End If

SQ_OPER = 2
PUB_CODPRO = Val(Right(LISPROCESOS.text, 8))
PUB_CODCIA = LK_CODCIA
LEER_PROC_LLAVE
Do Until proc_mayor.EOF
 If Val(proc_mayor!PRO_CODART) = wCODARTI Then
   MsgBox " YA Existe en Proceso. ", 48, Pub_Titulo
   Exit Sub
 End If
proc_mayor.MoveNext
Loop

OTRA_VEZ:
WSEC = Int((1000 * Rnd) + 1)
proc_mayor.MoveFirst
Do Until proc_mayor.EOF
 If Val(proc_mayor!PRO_SEC) = WSEC Then
    WSEC = 0
 End If
proc_mayor.MoveNext
Loop
If WSEC = 0 Then GoTo OTRA_VEZ


proc_mayor.AddNew
proc_mayor!pro_codpro = PUB_CODPRO
proc_mayor!PRO_CODART = Val(wCODARTI)
proc_mayor!PRO_CODCIA = LK_CODCIA
proc_mayor!PRO_SEC = WSEC
proc_mayor!PRO_ORDEN = WORDEN

proc_mayor.Update
WORDEN = WORDEN + 10
MSRDC1.Refresh
End Sub

Private Sub LISPROCESOS_Click()
On Error GoTo SALE
Frmformula.DBGrid1.Visible = False
Dim csele As String
Dim cadetodo As String
cod = Trim(Right(LISPROCESOS.text, 8))
If cod = "" Then
  MsgBox "Tiene que Seleccionar un Proceso ..", 48, Pub_Titulo
  LISPROCESOS.SetFocus
  Exit Sub
End If
DoEvents
momento.Caption = "Un Momento .."
DoEvents
csele = "PROCESOS.PRO_CODPRO = " & cod & " AND PROCESOS.PRO_CODCIA = '" & LK_CODCIA & "' ORDER BY PROCESOS.PRO_ORDEN"
cadetodo = " SELECT ARTI.ART_NOMBRE, PROCESOS.PRO_CANTIDAD, ARTI.ART_COSPRO, PROCESOS.PRO_SUBTOTAL,ARTICULO.ARM_STOCK, PROCESOS.PRO_ORDEN, ARTI.ART_NOMBRE  From  BDATOS.dbo.PROCESOS, BDATOS.dbo.ARTICULO, BDATOS.dbo.ARTI  Where  PROCESOS.PRO_CODART = ARTICULO.ARM_CODART  AND  PROCESOS.PRO_CODCIA = ARTICULO.ARM_CODCIA   AND PROCESOS.PRO_CODART = ARTI.ART_KEY  AND PROCESOS.PRO_CODCIA = ARTI.ART_CODCIA  AND " & csele
Frmformula.MSRDC1.SQL = cadetodo
Frmformula.MSRDC1.Refresh
Frmformula.DBGrid1.Visible = True
Frmformula.momento.Caption = ""
kPASE = True
SUMTOTAL
Exit Sub
SALE:
MsgBox Err.Description, 48, Pub_Titulo
Unload Frmformula
Exit Sub
End Sub

Private Sub LISPROCESOS_GotFocus()
Frmformula.DBGrid1.Visible = False
kPASE = False
End Sub


Public Sub SUMTOTAL()
On Error GoTo SALE
Dim SUMPESO
Dim SUMTOTAL
SUMPESO = 0
SUMTOTAL = 0
Frmformula.DBGrid1.Visible = False
DoEvents
momento.Caption = "Un Momento .."
DoEvents
If Frmformula.MSRDC1.Resultset.RowCount = 0 Then
  GoTo SIGUE
End If
Frmformula.MSRDC1.Resultset.MoveFirst
Do Until Frmformula.MSRDC1.Resultset.EOF
 SUMPESO = SUMPESO + Nulo_Valor0(Frmformula.MSRDC1.Resultset!PRO_CANTIDAD)
 SUMTOTAL = SUMTOTAL + Nulo_Valor0(Frmformula.MSRDC1.Resultset!PRO_SUBTOTAL)
 Frmformula.MSRDC1.Resultset.MoveNext
Loop
SIGUE:
txtPeso.text = Format(SUMPESO, "Currency")
txttotal.text = Format(SUMTOTAL, "Currency")
Frmformula.DBGrid1.Visible = True
momento.Caption = ""
If Frmformula.MSRDC1.Resultset.RowCount > 0 Then
 DBGrid1.SetFocus
End If

Exit Sub
SALE:
If Err.Number = 40503 Then
  Frmformula.DBGrid1.Visible = True
  momento.Caption = ""
  MsgBox "Datos Incorrectos Verificar ...!!!", 48, Pub_Titulo
  DBGrid1.SetFocus
Else
  MsgBox Err.Description, 48, Pub_Titulo
  Frmformula.MSRDC1.Cancel
  Frmformula.DBGrid1.Refresh
  Unload Frmformula
  
End If

End Sub
