VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmConsultaAtenciones 
   Caption         =   "Reporte de Consumos"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12390
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
   ScaleHeight     =   6855
   ScaleWidth      =   12390
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSDataListLib.DataCombo datContrata 
      Height          =   315
      Left            =   4560
      TabIndex        =   10
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.ComboBox cboComedor 
      Height          =   315
      ItemData        =   "frmConsultaAtenciones.frx":0000
      Left            =   1200
      List            =   "frmConsultaAtenciones.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "C&errar"
      Height          =   600
      Left            =   10440
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
   Begin CRVIEWERLibCtl.CRViewer crwVisor 
      Height          =   5415
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   12135
      DisplayGroupTree=   0   'False
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
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   427
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      Format          =   100073473
      CurrentDate     =   41498
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   300
      Left            =   4560
      TabIndex        =   1
      Top             =   420
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      Format          =   100073473
      CurrentDate     =   41498
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "C&onsultar"
      Height          =   600
      Left            =   9000
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTRATA:"
      Height          =   195
      Left            =   3480
      TabIndex        =   9
      Top             =   900
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMEDOR:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   900
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HASTA:"
      Height          =   195
      Left            =   3840
      TabIndex        =   5
      Top             =   480
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESDE:"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   675
   End
End
Attribute VB_Name = "frmConsultaAtenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cboComedor_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.datContrata.SetFocus
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdConsultar_Click()

    If Me.datContrata.BoundText = "" Then
        MsgBox "Debe elegir la Contrata.", vbCritical, Pub_Titulo
        Me.datContrata.SetFocus
        Exit Sub
    End If
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_ATENCION_REPORT_FILL"

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@F1", adDBTimeStamp, adParamInput, , Me.dtpDesde.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@F2", adDBTimeStamp, adParamInput, , Me.dtpHasta.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATA", adInteger, adParamInput, , Me.datContrata.BoundText)
    If Me.cboComedor.ListIndex <> 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, IIf(Me.cboComedor.ListIndex = 1, "01", "02"))
    End If

    Dim orsD As ADODB.Recordset

    Set orsD = oCmdEjec.Execute

    Dim objCrystal  As New CRAXDRT.Application
Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

        Dim crParamDef  As CRAXDRT.ParameterFieldDefinition
    Dim RutaReporte As String

    RutaReporte = "c:\Admin\Nordi\RESUMENATENCIONES.rpt"

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    
    Set crParamDefs = VReporte.ParameterFields

            For Each crParamDef In crParamDefs

                Select Case crParamDef.ParameterFieldName

                    Case "subtitulo"
                        crParamDef.AddCurrentValue "DESDE: " & CStr(Me.dtpDesde.Value) & Space(10) & "HASTA: " & CStr(Me.dtpHasta.Value)
                    Case "subtitulo2"
                    crParamDef.AddCurrentValue Me.datContrata.Text
                End Select

            Next
            
            
    VReporte.Database.SetDataSource orsD, 3, 1
    'frmprint.CRViewer1.ReportSource = VReporte
    'frmprint.CRViewer1.ViewReport
    crwVisor.ReportSource = VReporte
    crwVisor.ViewReport
    
    'VReporte.PrintOut False, 1, , 1, 1
    Set objCrystal = Nothing
    Set VReporte = Nothing

End Sub



Private Sub datContrata_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdConsultar_Click
End Sub

Private Sub dtpDesde_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.dtpHasta.SetFocus
End Sub

Private Sub dtpHasta_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.cboComedor.SetFocus
End Sub

Private Sub Form_Load()
Me.dtpHasta.Value = Date
Me.dtpDesde.Value = DateAdd("d", -1, Date)

Me.cboComedor.ListIndex = 0
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "sp_contrata_fill"
Dim oRSc As ADODB.Recordset
Set oRSc = oCmdEjec.Execute

Set Me.datContrata.RowSource = oRSc
Me.datContrata.ListField = oRSc.Fields(1).Name
Me.datContrata.BoundColumn = oRSc.Fields(0).Name
End Sub

Private Sub Form_Resize()
If (Me.ScaleWidth - 6300) <= 0 Then Exit Sub
If (Me.ScaleHeight - 1800) <= 0 Then Exit Sub
Me.crwVisor.Width = Me.ScaleWidth
Me.crwVisor.Height = Me.ScaleHeight - 900
End Sub
