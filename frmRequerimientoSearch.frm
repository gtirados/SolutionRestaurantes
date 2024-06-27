VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRequerimientoSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda Requerimiento"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVer 
      Caption         =   "&Visualizar"
      Height          =   360
      Left            =   8265
      TabIndex        =   6
      Top             =   480
      Width           =   990
   End
   Begin MSComctlLib.ListView lvDatos 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   300
      Left            =   3960
      TabIndex        =   3
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   205455361
      CurrentDate     =   42121
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   205455361
      CurrentDate     =   42121
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   360
      Left            =   7200
      TabIndex        =   5
      Top             =   480
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   3360
      TabIndex        =   1
      Top             =   413
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   413
      Width           =   615
   End
End
Attribute VB_Name = "frmRequerimientoSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBuscar_Click()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_REQUERIMIENTO_SEARCH"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESDE", adDBTimeStamp, adParamInput, , Me.dtpDesde.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@HASTA", adDBTimeStamp, adParamInput, , Me.dtpHasta.Value)

    Dim ORSd As ADODB.Recordset

    Set ORSd = oCmdEjec.Execute
    Me.lvDatos.ListItems.Clear

    Dim Xitem As Object

    Do While Not ORSd.EOF
        Set Xitem = Me.lvDatos.ListItems.Add(, , ORSd!Nro)
        Xitem.SubItems(1) = ORSd!fecha
        Xitem.SubItems(2) = ORSd!USUARIO
        Xitem.SubItems(3) = ORSd!PENDIENTE

        ORSd.MoveNext
    Loop

End Sub

Private Sub cmdVer_Click()
VisualizarReq
End Sub

Private Sub Form_Load()

    With Me.lvDatos
        .ColumnHeaders.Add , , "Nro Req"
        .ColumnHeaders.Add , , "Fecha"
        .ColumnHeaders.Add , , "Usuario"
        .ColumnHeaders.Add , , "Pendiente"
        .LabelEdit = lvwAutomatic
        .View = lvwReport
    End With

End Sub

Private Sub VisualizarReq()
  LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_REQUERIMIENTO_VIEW"
    oCmdEjec.CommandType = adCmdStoredProc
    
    
    Dim RutaReporte As String


    RutaReporte = PUB_RUTA_REPORTE & "Req.rpt"
    

    'OBTENIENDO DATOS DEL CLIENTE
'    LimpiaParametros oCmdEjec
'    oCmdEjec.CommandText = "SP_DELIVERY_DOCTOCLIENTE"
'    oCmdEjec.CommandType = adCmdStoredProc
'
'    'Dim orsD As ADODB.Recordset
'
'    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
'    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , frmDeliveryApp.lblCliente.Caption)
'
'    Set orsD = oCmdEjec.Execute
'
'    Dim Vdocto As String
'
'    If Not orsD.EOF Then
'        Vdocto = Trim(orsD!DOCTO)
'    End If

    On Error GoTo printe

'    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
'
'    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition
'
   Dim objCrystal  As New CRAXDRT.Application

'    Dim vIgv        As Currency
'
'    Dim vSubTotal   As Currency
'
'
'vSubTotal = Round((xTotal / ((100 + LK_IGV) / 100)), 2)
'
'vIgv = xTotal - vSubTotal

  

    

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
'    Set crParamDefs = VReporte.ParameterFields
'
'    For Each crParamDef In crParamDefs
'
'        Select Case crParamDef.ParameterFieldName
'
'            Case "cliente"
'
'                crParamDef.AddCurrentValue Trim(xcliente)
'
'            Case "FechaEmi"
'                crParamDef.AddCurrentValue LK_FECHA_DIA
'
'            Case "Son"
'                crParamDef.AddCurrentValue CONVER_LETRAS(CStr(xTotal), "S")
'
'            Case "total"
'                crParamDef.AddCurrentValue CStr(FormatNumber(xTotal, 2)) ' CStr(xTotal)
'
'            Case "subtotal"
'                crParamDef.AddCurrentValue CStr(FormatNumber(vSubTotal, 2))
'
'            Case "igv"
'                crParamDef.AddCurrentValue CStr(FormatNumber(vIgv, 2))
'
'            Case "SerFac"
'                crParamDef.AddCurrentValue XsERIE
'
'            Case "NumFac"
'                crParamDef.AddCurrentValue CStr(xNumero)
'
'            Case "DirClie"
'
'                'crParamDef.AddCurrentValue frmDeliveryApp.DatDireccion.Text
'                crParamDef.AddCurrentValue xDireccion
'
'            Case "RucClie"
'
'                'crParamDef.AddCurrentValue Vdocto
'                crParamDef.AddCurrentValue xRuc
'
'            'Case "Importe" 'linea nueva
'                'crParamDef.AddCurrentValue frmDeliveryApp.lblTot.Caption 'linea nueva
'
'
'        End Select
'
'    Next

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc

    'oCmdEjec.CommandText = "SpPrintComanda"

    Dim rsd As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@nroreq", adBigInt, adParamInput, , Me.lvDatos.SelectedItem.Text)
    
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
        frmVisor.cr.ReportSource = VReporte
        frmVisor.cr.ViewReport
        frmVisor.Show vbModal
    
    End If

    'Set VReporte = Nothing
    'Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    'cr.DataSource = VReporte
    'cr.Destination = crptToWindow
    '
    'rsd.Filter = "PED_FAMILIA=3"
    'If Not rsd.EOF Then
    '    VReporte.Database.SetDataSource rsd, 3, 1 'lleno el objeto reporte
    '    VReporte.SelectPrinter Printer.DriverName, "\\SERVIDOR\Canon MP140 series Printer", Printer.Port 'doPDF v6
    '    VReporte.PrintOut ' , 1, , 1, 1
    'End If

    Set objCrystal = Nothing
    Set VReporte = Nothing

    Exit Sub

printe:
    MostrarErrores Err

End Sub
