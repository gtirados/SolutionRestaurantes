VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEventoConsulta 
   Caption         =   "Consulta de Eventos"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13845
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   7065
   ScaleWidth      =   13845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin CRVIEWERLibCtl.CRViewer crvReporte 
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   13695
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
   Begin MSDataListLib.DataCombo dcboEvento 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Evento"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   585
   End
End
Attribute VB_Name = "frmEventoConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConsultar_Click()

  

    Dim rsd As ADODB.Recordset

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "USP_EVENTO_DETALLE"
    

    With oCmdEjec

        

        .Parameters.Append .CreateParameter("@idevento", adInteger, adParamInput, , Me.dcboEvento.BoundText)


    End With

    Set rsd = oCmdEjec.Execute
    
  
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal  As New CRAXDRT.APPLICATION

    Dim RutaReporte As String

    RutaReporte = "C:\Admin\Nordi\EventoDetalle.rpt"
    

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
 
 

    VReporte.DataBase.SetDataSource rsd, , 1

 
    Me.crvReporte.ReportSource = VReporte
    Me.crvReporte.ViewReport

    Set objCrystal = Nothing
    Set VReporte = Nothing
End Sub

Private Sub Form_Load()
LlenarCombo
End Sub


Private Sub LlenarCombo()
 LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "USP_EVENTO_LIST"
Dim orsData As New ADODB.Recordset
Set orsData = oCmdEjec.Execute()

If Not orsData.EOF Then

    'frmcomandamozomesa.gCodMesa = Me.lblNomMesa(Index).Tag
    Set Me.dcboEvento.RowSource = orsData
    Me.dcboEvento.ListField = "evento"
    Me.dcboEvento.BoundColumn = "ide"
    
End If
End Sub

Private Sub Form_Resize()
Me.crvReporte.Width = Me.ScaleWidth
Me.crvReporte.Height = Me.ScaleHeight
End Sub
