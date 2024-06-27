VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmResumenDescuentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Descuentos"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
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
   ScaleHeight     =   5415
   ScaleWidth      =   9060
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "Mostrar"
      Height          =   600
      Left            =   6120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "Imprimir"
      Height          =   600
      Left            =   7680
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvDatos 
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   8775
      _ExtentX        =   15478
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
   Begin MSMask.MaskEdBox mbHasta 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   375
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mbDesde 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   375
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmResumenDescuentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMostrar_Click()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_COMANDA_CONSULTADESCUENTOS"

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@desde", adDBTimeStamp, adParamInput, , Me.mbDesde.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@hasta", adDBTimeStamp, adParamInput, , Me.mbHasta.Text)
    
    Dim orsD As ADODB.Recordset
    Set orsD = oCmdEjec.Execute
    
    Me.lvDatos.ListItems.Clear
    
    Dim xItem As Object
    Do While Not orsD.EOF
        Set xItem = Me.lvDatos.ListItems.Add(, , orsD!Comanda)
        xItem.SubItems(1) = orsD!fecha
        xItem.SubItems(2) = orsD!DSto
        xItem.SubItems(3) = orsD!usuario
        orsD.MoveNext
    
    Loop
End Sub

Private Sub cmdVer_Click()

    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal  As New CRAXDRT.Application

    Dim RutaReporte As String

    RutaReporte = PUB_RUTA_REPORTE & "resumenDsctos.rpt"  'LO CAMBIA CUANDO QUIERA

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    Set crParamDefs = VReporte.ParameterFields

    For Each crParamDef In crParamDefs

        Select Case crParamDef.ParameterFieldName

            Case "pFECHAS"
                crParamDef.AddCurrentValue "DESDE: " & Me.mbDesde.Text & Space(10) & "DESDE: " & Me.mbHasta.Text ' str(vPlato)
        End Select

    Next

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.CommandText = "SP_COMANDA_CONSULTADESCUENTOS"
    'oCmdEjec.CommandText = "SpPrintComanda"

    Dim rsd As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@desde", adDBTimeStamp, adParamInput, , Me.mbDesde.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@hasta", adDBTimeStamp, adParamInput, , Me.mbHasta.Text)
     
   Set rsd = oCmdEjec.Execute

  

    VReporte.Database.SetDataSource rsd, 3, 1
    'frmprint.CRViewer1.ReportSource = VReporte
    'frmprint.CRViewer1.ViewReport
    
      
        frmVisor.cr.ReportSource = VReporte
          frmVisor.cr.ViewReport
        frmVisor.Show vbModal
    Set objCrystal = Nothing
    Set VReporte = Nothing
End Sub

Private Sub Form_Load()
ConfiguraLV
Me.mbDesde.Text = LK_FECHA_DIA
Me.mbHasta.Text = LK_FECHA_DIA
End Sub
Private Sub ConfiguraLV()
With Me.lvDatos
.Gridlines = True
    .LabelEdit = lvwManual
    .FullRowSelect = True
    .View = lvwReport
    
    .HideSelection = False
    .ColumnHeaders.Add , , "Comanda"
    .ColumnHeaders.Add , , "Fecha"
    .ColumnHeaders.Add , , "Dscto"
    .ColumnHeaders.Add , , "Usuario"
    

End With
End Sub
