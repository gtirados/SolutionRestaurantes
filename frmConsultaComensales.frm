VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConsultaComensales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Comensales"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11040
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   11040
   Begin MSComctlLib.ListView lvDatos 
      Height          =   3735
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   6588
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   210
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   270
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   97189889
      CurrentDate     =   41962
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   270
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   97189889
      CurrentDate     =   41962
   End
   Begin VB.Label lblComensales 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   4680
      TabIndex        =   1
      Top             =   360
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde:"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmConsultaComensales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SP_COMENSALES_CONSULTA"

Dim orsD As ADODB.Recordset
Dim xCom As Double
xCom = 0

oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESDE", adDBTimeStamp, adParamInput, , Me.dtpDesde.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@HASTA", adDBTimeStamp, adParamInput, , Me.dtpHasta.Value)
    Me.lvDatos.ListItems.Clear
Set orsD = oCmdEjec.Execute
Dim ITEMx As Object
Do While Not orsD.EOF
    Set ITEMx = Me.lvDatos.ListItems.Add(, , orsD!fecha)
    ITEMx.SubItems(1) = orsD!PEDIDO
    ITEMx.SubItems(2) = orsD!MESA
    ITEMx.SubItems(3) = orsD!mozo
    ITEMx.SubItems(4) = orsD!comen
    xCom = orsD!comen + xCom
    orsD.MoveNext
Loop
Me.lblComensales.Caption = "Total de comensales: " & xCom
End Sub

Private Sub Form_Load()
With Me.lvDatos
    .ColumnHeaders.Add , , "Fecha"
    .ColumnHeaders.Add , , "Pedido"
    .ColumnHeaders.Add , , "Mesa", 2000
    .ColumnHeaders.Add , , "Mozo", 4000
    .ColumnHeaders.Add , , "Comensales"
    .View = lvwReport
     .Gridlines = True
    .LabelEdit = lvwManual

    .FullRowSelect = True
End With
dtpDesde.Value = LK_FECHA_DIA
dtpHasta.Value = LK_FECHA_DIA
End Sub
