VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContratosList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Contratos Programados"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13350
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContratosList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   13350
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   585
      Left            =   12120
      Picture         =   "frmContratosList.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9360
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratosList.frx":0714
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratosList.frx":0AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratosList.frx":0E48
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratosList.frx":11E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratosList.frx":157C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   13350
      _ExtentX        =   23548
      _ExtentY        =   1058
      ButtonWidth     =   1746
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Editar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Visualizar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Finalizar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Anular"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvData 
      Height          =   4455
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   7858
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
   Begin VB.OptionButton OptCliente 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame fraCliente 
      Height          =   615
      Left            =   5400
      TabIndex        =   10
      Top             =   720
      Width           =   6615
      Begin VB.TextBox txtCliente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.OptionButton OptFechas 
      Caption         =   "Fechas"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.Frame FraFechas 
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   4935
      Begin MSMask.MaskEdBox MasDesde 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   150
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MasHasta 
         Height          =   315
         Left            =   3360
         TabIndex        =   1
         Top             =   150
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   2640
         TabIndex        =   9
         Top             =   210
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   210
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmContratosList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vPasa As Boolean

Private Sub cmdBuscar_Click()

    If Me.OptFechas.Value Then
        If Not IsDate(Me.MasDesde.Text) Then
            MsgBox "Fecha <DESDE> incorrecta.", vbCritical, TituloSistema
            vPasa = False
            Me.MasDesde.SetFocus
        ElseIf Not IsDate(Me.MasHasta.Text) Then
            vPasa = False
            MsgBox "Fecha <HASTA> incorrecta.", vbCritical, TituloSistema
            Me.MasHasta.SetFocus
        End If

    Else

        If Len(Trim(Me.txtCliente.Text)) = 0 Then
            vPasa = False
            MsgBox "Debe ingresar al Cliente.", vbCritical, TituloSistema
            Me.txtCliente.SetFocus
        End If
    End If

    Dim orsDATA As ADODB.Recordset

    Me.lvData.ListItems.Clear

    If vPasa Then
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPCONTRATOSLIST"

        With oCmdEjec
            .Parameters.Append .CreateParameter("@XFECHAS", adBoolean, adParamInput, , IIf(Me.OptFechas.Value, True, False))
            .Parameters.Append .CreateParameter("@FECHAINI", adDBTimeStamp, adParamInput, , IIf(Me.OptFechas.Value, Me.MasDesde.Text, vbNull))
            .Parameters.Append .CreateParameter("@FECHAFIN", adDBTimeStamp, adParamInput, , IIf(Me.OptFechas.Value, Me.MasHasta.Text, vbNull))
            .Parameters.Append .CreateParameter("@SEARCH", adVarChar, adParamInput, 150, IIf(Me.OptCliente.Value, Me.txtCliente.Text, vbNull))
            .Parameters.Append .CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            Set orsDATA = .Execute
        End With
        
        Dim itemX As Object

        Do While Not orsDATA.EOF
            Set itemX = Me.lvData.ListItems.Add(, , orsDATA!NROCONTRATO)
            itemX.SubItems(1) = Trim(orsDATA!CLIENTE)
            itemX.SubItems(2) = orsDATA!INICIOEVENTO
            itemX.SubItems(3) = orsDATA!TERMINOEVENTO
            itemX.SubItems(4) = orsDATA!EXCLUSIVO
            itemX.SubItems(5) = orsDATA!Total
            itemX.SubItems(6) = orsDATA!ACUENTA
            itemX.SubItems(7) = orsDATA!SALDO
            itemX.SubItems(8) = orsDATA!ESTADO
            orsDATA.MoveNext
        Loop

    End If

End Sub

Private Sub Form_Load()
    vPasa = True
    ConfigurarLv
    Me.MasDesde.Text = "01/" & Right("00" & CStr(Month(Date)), 2) & "/" & CStr(Year(Date))
    Me.MasHasta.Text = DateAdd("m", 1, Me.MasDesde.Text) - 1

    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

End Sub

Private Sub ConfigurarLv()
    With Me.lvData
    
        .ColumnHeaders.Add , , "NRO CONTRATO"
        .ColumnHeaders.Add , , "CLIENTE"
        .ColumnHeaders.Add , , "INICIO EVENTO", 1200
        .ColumnHeaders.Add , , "TERMINO EVENTO", 1200
        .ColumnHeaders.Add , , "EXCLUSIVO"
        .ColumnHeaders.Add , , "TOTAL"
        .ColumnHeaders.Add , , "A CUENTA"
        .ColumnHeaders.Add , , "SALDO"
        .ColumnHeaders.Add , , "ESTADO"
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .View = lvwReport
        .Gridlines = True
    End With
End Sub

Private Sub MasDesde_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Me.MasHasta.SetFocus
End Sub

Private Sub MasHasta_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then cmdBuscar_Click
End Sub

Private Sub OptCliente_Click()

    If Me.OptCliente.Value Then
        Me.MasDesde.Enabled = False
        Me.MasHasta.Enabled = False
        Me.txtCliente.Enabled = True
        Me.txtCliente.SetFocus
    End If

End Sub

Private Sub OptFechas_Click()

    If Me.OptFechas.Value Then
        Me.MasDesde.Enabled = True
        Me.MasHasta.Enabled = True
        Me.txtCliente.Enabled = False
        Me.MasDesde.SetFocus
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error GoTo Menu

    Select Case Button.Index

        Case 1
            frmContratos.VNuevo = True
            frmContratos.Show vbModal

        Case 2

            If Me.lvData.ListItems.count = 0 Then
                MsgBox "Debe seleccionar un contrato para Editarlo.", vbCritical, TituloSistema
                Me.lvData.SetFocus
            ElseIf Me.lvData.SelectedItem.SubItems(8) = "ANULADO" Then
                MsgBox "No se puede Alterar un Contrato Anulado.", vbInformation, TituloSistema
                Me.lvData.SetFocus
            ElseIf Me.lvData.SelectedItem.SubItems(8) = "FINALIZADO" Then
                MsgBox "No se puede Alterar un Contrato Finalizado.", vbInformation, TituloSistema
                Me.lvData.SetFocus
            Else
                frmContratos.VNuevo = False
                frmContratos.vIDContrato = Me.lvData.SelectedItem.Text
                frmContratos.dtpInicio.Value = Me.lvData.SelectedItem.SubItems(2)
                frmContratos.dtpTermino.Value = Me.lvData.SelectedItem.SubItems(3)
                
                frmContratos.Show vbModal
                If frmContratos.vGraba Then cmdBuscar_Click
            End If

        Case 3
            VisualizarContrato Me.lvData.SelectedItem.Text

        Case 4

            'FINALIZA CONTRATO
            If Me.lvData.ListItems.count = 0 Then
                MsgBox "Debe seleccionar un contrato para Finalizarlo.", vbCritical, TituloSistema
                Me.lvData.SetFocus
            ElseIf Me.lvData.SelectedItem.SubItems(8) = "ANULADO" Then
                MsgBox "El contrato seleccionado ya fue Anulado" & vbCrLf & "no se puede Finalizar.", vbCritical, TituloSistema
                Me.lvData.SetFocus
            ElseIf Me.lvData.SelectedItem.SubItems(8) = "FINALIZADO" Then
                MsgBox "El contrato seleccionado ya se encuentra Finalizado.", vbInformation, TituloSistema
                Me.lvData.SetFocus
            ElseIf MsgBox("¿Desea continuar con la Finalización del Contrato Seleccionado.?", vbQuestion + vbYesNo, TituloSistema) = vbNo Then
                Me.lvData.SetFocus
            Else
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SPFINALIZARANULARCONTRATO"
            
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adBigInt, adParamInput, , Me.lvData.SelectedItem.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ANULANDO", adBoolean, adParamInput, , 0)
                oCmdEjec.Execute
                
                Me.lvData.SelectedItem.SubItems(8) = "FINALIZADO"
                MsgBox "Contrato Finalizado Correctamente.", vbInformation, TituloSistema
            End If

        Case 5

            'ANULA CONTRATO
            If Me.lvData.ListItems.count = 0 Then
                MsgBox "Debe seleccionar un contrato para Anularlo.", vbCritical, TituloSistema
                Me.lvData.SetFocus
            ElseIf Me.lvData.SelectedItem.SubItems(8) = "FINALIZADO" Then
                MsgBox "El Contrato seleccionado se encuentra Finalizado." + vbCrLf + "No se puede Anular.", vbCritical, TituloSistema
                Me.lvData.SetFocus
            ElseIf Me.lvData.SelectedItem.SubItems(8) = "ANULADO" Then
                MsgBox "El Contrato seleccionado ya se encuentra Anulado", vbInformation, TituloSistema
                Me.lvData.SetFocus
            ElseIf MsgBox("¿Desea continuar con la Anulación del Contrato Seleccionado.?", vbQuestion + vbYesNo, TituloSistema) = vbNo Then
                Me.lvData.SetFocus
            Else
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SPFINALIZARANULARCONTRATO"
            
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adBigInt, adParamInput, , Me.lvData.SelectedItem.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ANULANDO", adBoolean, adParamInput, , 1)
                oCmdEjec.Execute
                
                Me.lvData.SelectedItem.SubItems(8) = "ANULADO"
                MsgBox "Contrato Anulado Correctamente.", vbInformation, TituloSistema
            End If

    End Select

    Exit Sub

Menu:
    MsgBox Err.Description, vbCritical, TituloSistema
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then cmdBuscar_Click
End Sub

Private Sub VisualizarContrato(xIdContrato As Integer)

    On Error GoTo Ver

    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal  As New CRAXDRT.Application

    Dim RutaReporte As String
    Dim oUSER As String, oCLAVE As String, oLOCAL As String

    'RutaReporte = "C:\Admin\Nordi\Comanda1.rpt"
    
    
    'DATOS COMPLEMENTARIOS
    Dim orsC As ADODB.Recordset
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPDATOSCOMPLEMENTARIOSCONTRATOS"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    
    Set orsC = oCmdEjec.Execute
    
    
    'RutaReporte = "d:\VISTACONTRATO.rpt"
    RutaReporte = Trim(orsC!RutaReporte) + "VISTACONTRATO.rpt"
    oUSER = orsC!usuario
    oCLAVE = orsC!Clave
    oLOCAL = orsC!LOCAL

    If VReporte Is Nothing Then VReporte = New CRAXDRT.Report
    
    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    Set crParamDefs = VReporte.ParameterFields

    For Each crParamDef In crParamDefs

        Select Case crParamDef.ParameterFieldName

            Case "TITULO"
                crParamDef.AddCurrentValue "CONTRATO CENA " & oLOCAL & " - Nº - " & Me.lvData.SelectedItem.Text
        End Select

    Next

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.CommandText = "SPVISUALIZARCONTRATO"
    'oCmdEjec.CommandText = "SpPrintComanda"

    Dim rsd As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adBigInt, adParamInput, , xIdContrato)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    Set rsd = oCmdEjec.Execute

'    Dim RSS As ADODB.Recordset
'
'    LimpiaParametros oCmdEjec
'
'    oCmdEjec.CommandText = "SPVISUALIZARCONTRATO2"
'
'    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adBigInt, adParamInput, , xIdContrato)
'    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
'
'    Set RSS = oCmdEjec.Execute

    'SUB REPORTE
    Dim VReporteS As New CRAXDRT.Report

    VReporte.Database.SetDataSource rsd, , 1  'lleno el objeto reporte

    Set VReporteS = VReporte.OpenSubreport("DETALLE")
    
    VReporte.OpenSubreport("DETALLE").Database.LogOnServer "p2sodbc.dll", "DSN_DATOS", "bdatos", oUSER, oCLAVE
    'VReporte.OpenSubreport("DETALLE").Database.LogOnServer "p2sodbc.dll", "DSN_DATOS", "bdatos", oUSER, "accesodenegado"
    
    
    Set crParamDefs = VReporteS.ParameterFields

    For Each crParamDef In crParamDefs

        Select Case crParamDef.ParameterFieldName

            Case "Pm-ado.idcontrato"
                crParamDef.AddCurrentValue xIdContrato
                Case "Pm-ado.CODCIA"
                crParamDef.AddCurrentValue LK_CODCIA
        End Select

    Next
       
    'VReporteS.Database.SetDataSource RSS, , 1
 
    'VReporteS.ReadRecords
    frmContratosReporte.crContrato.ReportSource = VReporte

    'frmContratosReporte.crContrato.Refresh
    frmContratosReporte.crContrato.ViewReport

    frmContratosReporte.Show
    Set VReporte = Nothing
    Set VReporteS = Nothing

    Exit Sub

Ver:
    MostrarErrores Err
End Sub

