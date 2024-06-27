VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmConsultaPlatosExtornados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Platos Extornados"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
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
   ScaleWidth      =   11550
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Visualizar"
      Height          =   480
      Left            =   8880
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Height          =   480
      Left            =   10200
      TabIndex        =   9
      Top             =   360
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvDatra 
      Height          =   4455
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSDataListLib.DataCombo DatMesa 
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "Mesa"
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   80019457
      CurrentDate     =   41914
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   80019457
      CurrentDate     =   41914
   End
   Begin MSDataListLib.DataCombo datUsuario 
      Height          =   315
      Left            =   4680
      TabIndex        =   7
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "Mesa"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      Height          =   195
      Left            =   3840
      TabIndex        =   6
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mesa:"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   3120
      TabIndex        =   1
      Top             =   210
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   615
   End
End
Attribute VB_Name = "frmConsultaPlatosExtornados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CargarCombos()

    Dim orsC As ADODB.Recordset

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_CONSULTA_EXTORNADOS_COMBOS"

    Set orsC = oCmdEjec.Execute(, LK_CODCIA)

    Dim ORSd As ADODB.Recordset

    Set Me.DatMesa.RowSource = orsC
    Me.DatMesa.ListField = orsC.Fields(1).Name
    Me.DatMesa.BoundColumn = orsC.Fields(0).Name
    Me.DatMesa.BoundText = -1

    Set ORSd = orsC.NextRecordset

    Set Me.datUsuario.RowSource = ORSd
    Me.datUsuario.ListField = ORSd.Fields(1).Name
    Me.datUsuario.BoundColumn = ORSd.Fields(0).Name
    Me.datUsuario.BoundText = -1
End Sub

Private Sub cmdconsultar_Click()
On Error GoTo Busqueda
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_CONSULTA_EXTORNADOS"
    Me.lvDatra.ListItems.Clear

    With oCmdEjec
        .Parameters.Append .CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        .Parameters.Append .CreateParameter("@DESDE", adDBTimeStamp, adParamInput, , Me.dtpDesde.Value)
        .Parameters.Append .CreateParameter("@HASTA", adDBTimeStamp, adParamInput, , Me.dtpHasta.Value)
        .Parameters.Append .CreateParameter("@CODUSUARIO", adVarChar, adParamInput, 10, Me.datUsuario.BoundText)

        If Me.DatMesa.BoundText <> "-1" Then
            .Parameters.Append .CreateParameter("@CODMESA", adVarChar, adParamInput, 10, Me.DatMesa.BoundText)
        End If

    End With

    Dim oRSdata As ADODB.Recordset

    Set oRSdata = oCmdEjec.Execute

    Dim itemX As Object

    Do While Not oRSdata.EOF
        Set itemX = Me.lvDatra.ListItems.Add(, , oRSdata!fecha)
        itemX.SubItems(1) = oRSdata!Comanda
        itemX.SubItems(2) = oRSdata!producto
        itemX.SubItems(3) = oRSdata!Cantidad
        itemX.SubItems(4) = oRSdata!mesa
        itemX.SubItems(5) = oRSdata!mozo
        itemX.SubItems(6) = oRSdata!USUARIO
        itemX.SubItems(7) = oRSdata!MOTIVO
        oRSdata.MoveNext
    Loop
          Exit Sub
Busqueda:
          MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub

Private Sub cmdprint_Click()

  

    Dim rsd As ADODB.Recordset

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_CONSULTA_EXTORNADOS"
    Me.lvDatra.ListItems.Clear

    With oCmdEjec
        .Parameters.Append .CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        .Parameters.Append .CreateParameter("@DESDE", adDBTimeStamp, adParamInput, , Me.dtpDesde.Value)
        .Parameters.Append .CreateParameter("@HASTA", adDBTimeStamp, adParamInput, , Me.dtpHasta.Value)
        .Parameters.Append .CreateParameter("@CODUSUARIO", adVarChar, adParamInput, 10, Me.datUsuario.BoundText)

        If Me.DatMesa.BoundText <> "-1" Then
            .Parameters.Append .CreateParameter("@CODMESA", adVarChar, adParamInput, 10, Me.DatMesa.BoundText)
        End If

    End With

    Set rsd = oCmdEjec.Execute
    
  
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal  As New CRAXDRT.APPLICATION

    Dim RutaReporte As String

    RutaReporte = "C:\Admin\Nordi\PlatosExtornados.rpt"
    

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
 
 

    VReporte.DataBase.SetDataSource rsd, , 1

 
    frmVisor.cr.ReportSource = VReporte
    frmVisor.cr.ViewReport
    frmVisor.Show
    Set objCrystal = Nothing
    Set VReporte = Nothing
End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
Me.dtpHasta.Value = LK_FECHA_DIA
Me.dtpDesde.Value = DateAdd("m", -1, LK_FECHA_DIA)
CargarCombos
ConfiguraLV
End Sub
Private Sub ConfiguraLV()
'fecha
'PRODUCTO
'Cantidad
'HORA opcional
'MESA
'mozo
'Nro COMANDA
With Me.lvDatra
 .LabelEdit = lvwManual
    .FullRowSelect = True
    .View = lvwReport
    
    .HideSelection = False
    .ColumnHeaders.Add , , "Fecha"
    .ColumnHeaders.Add , , "Nro Comanda"
    .ColumnHeaders.Add , , "Producto", 2400
    .ColumnHeaders.Add , , "Cantidad", 1000
    .ColumnHeaders.Add , , "Mesa"
    .ColumnHeaders.Add , , "Mozo", 2800
    .ColumnHeaders.Add , , "Usuario", 2800
    .ColumnHeaders.Add , , "Motivo", 2800
    .MultiSelect = False
End With
End Sub

