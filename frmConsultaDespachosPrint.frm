VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmConsultaDespachosPrint 
   Caption         =   "Visor"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
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
   ScaleHeight     =   6015
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer crvVisor 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
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
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmConsultaDespachosPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gMesa As Integer
Public gFecha As Date

Private Sub Form_Load()

    On Error GoTo printe

    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    
    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal  As New CRAXDRT.Application

    Dim vIgv        As Currency

    Dim vSubTotal   As Currency

    Dim RutaReporte As String
    
    RutaReporte = "C:\Admin\Nordi\ConsultaDespachos.rpt"

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    Set crParamDefs = VReporte.ParameterFields
    
    For Each crParamDef In crParamDefs
    
        Select Case crParamDef.ParameterFieldName
    
            Case "pFecha"
                'crParamDef.AddCurrentValue IIf(Len(Trim(Me.txtRS.Text)) = 0, "CLIENTES VARIOS", Trim(Me.txtRS.Text))
                crParamDef.AddCurrentValue CStr(gFecha)

            Case "pMesa"
                crParamDef.AddCurrentValue frmConsultaDespachos.DatMesas.Text
    
        End Select
    
    Next

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc

    'oCmdEjec.CommandText = "SpPrintComanda"

    Dim rsd As ADODB.Recordset

    oCmdEjec.CommandText = "SP_CONSULTA_DESPACHOS"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , gFecha)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MESA", adVarChar, adParamInput, 10, gMesa)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fbg", adChar, adParamInput, 1, IIf(Me.cboTipoDocto.ListIndex = 0, "F", IIf(Me.cboTipoDocto.ListIndex = 1, "B", "")))
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NroCom", adInteger, adParamInput, , vNroCom)

    Set rsd = oCmdEjec.Execute

    'COCINA
    'rsd.Filter = "PED_FAMILIA=2"

    ' For i = 0 To Printers.count - 1
    '        MsgBox Printers(i).DeviceName
    '    Next
  
    If Not rsd.EOF Then

        VReporte.Database.SetDataSource rsd, 3, 1 'lleno el objeto reporte
        'VReporte.SelectPrinter Printer.DriverName, "\\laptop\doPDF v6", Printer.Port
        '
        'VReporte.PrintOut False, 1, , 1, 1
        'frmVisor.cr.ReportSource = VReporte
        Me.crvVisor.ReportSource = VReporte
        Me.crvVisor.ViewReport
        'frmVisor.cr.ViewReport
        'frmVisor.Show vbModal
    
    End If

    Set objCrystal = Nothing
    Set VReporte = Nothing

    Exit Sub

printe:
    MostrarErrores Err
 
End Sub

Private Sub Form_Resize()
Me.crvVisor.Width = Me.ScaleWidth
Me.crvVisor.Height = Me.ScaleHeight
End Sub
