VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmConsultaAtenciones2 
   Caption         =   "Reporte Detallado de Consumos"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   11340
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "C&errar"
      Height          =   735
      Left            =   10080
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo datContrata 
      Height          =   315
      Left            =   6480
      TabIndex        =   6
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin CRVIEWERLibCtl.CRViewer crvView 
      Height          =   5175
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   11175
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "C&onsultar"
      Height          =   735
      Left            =   8760
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   300
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      Format          =   94830593
      CurrentDate     =   41502
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   300
      Left            =   3480
      TabIndex        =   3
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      Format          =   94830593
      CurrentDate     =   41502
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CONTRATA:"
      Height          =   195
      Left            =   5400
      TabIndex        =   8
      Top             =   300
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "HASTA:"
      Height          =   195
      Left            =   2760
      TabIndex        =   1
      Top             =   300
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DESDE:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   675
   End
End
Attribute VB_Name = "frmConsultaAtenciones2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdconsultar_Click()

    If Me.datContrata.BoundText = "" Then
        MsgBox "Debe elegir la Contrata.", vbCritical, Pub_Titulo
        Me.datContrata.SetFocus
        Exit Sub
    End If
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_ATENCION_REPORT_DIARIO"

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@F1", adDBTimeStamp, adParamInput, , Me.dtpDesde.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@F2", adDBTimeStamp, adParamInput, , Me.dtpHasta.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATA", adInteger, adParamInput, , Me.datContrata.BoundText)
   ' If Me.cboComedor.ListIndex <> 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
   ' End If

    Dim orsD As ADODB.Recordset

    Set orsD = oCmdEjec.Execute

    Dim objCrystal  As New CRAXDRT.Application
Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

        Dim crParamDef  As CRAXDRT.ParameterFieldDefinition
    Dim RutaReporte As String

    RutaReporte = "c:\Admin\Nordi\DETALLADOATENCIONES.rpt"

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    
    Set crParamDefs = VReporte.ParameterFields

            For Each crParamDef In crParamDefs

                Select Case crParamDef.ParameterFieldName

                    Case "subtitulo"
                        crParamDef.AddCurrentValue "DESDE: " & CStr(Me.dtpDesde.Value) & Space(10) & "HASTA: " & CStr(Me.dtpHasta.Value)
'                    Case "subtitulo2"
'                    crParamDef.AddCurrentValue Me.datContrata.Text
                End Select

            Next
            
            
    VReporte.Database.SetDataSource orsD, 3, 1
    'frmprint.CRViewer1.ReportSource = VReporte
    'frmprint.CRViewer1.ViewReport
    Me.crvView.ReportSource = VReporte
    crvView.ViewReport
    
    'VReporte.PrintOut False, 1, , 1, 1
    Set objCrystal = Nothing
    Set VReporte = Nothing

End Sub

Private Sub datContrata_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdconsultar_Click
End Sub

Private Sub dtpDesde_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.dtpHasta.SetFocus
End Sub

Private Sub dtpHasta_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.datContrata.SetFocus
End Sub

Private Sub Form_Load()
Me.dtpHasta.Value = Date
Me.dtpDesde.Value = DateAdd("d", -1, Date)

Dim orsC As ADODB.Recordset

LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "sp_contrata_fill2"

Set orsC = oCmdEjec.Execute

Set Me.datContrata.RowSource = orsC
Me.datContrata.ListField = orsC.Fields(1).Name
Me.datContrata.BoundColumn = orsC.Fields(0).Name
Me.datContrata.BoundText = "-1"
End Sub

Private Sub Form_Resize()
If (Me.ScaleWidth - 6300) <= 0 Then Exit Sub
If (Me.ScaleHeight - 1800) <= 0 Then Exit Sub
Me.crvView.Width = Me.ScaleWidth
Me.crvView.Height = Me.ScaleHeight - 900
End Sub
