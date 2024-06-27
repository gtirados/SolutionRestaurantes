VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmContratos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elaboración de Contratos"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContratos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtpFechaImpresion 
      Height          =   290
      Left            =   7680
      TabIndex        =   0
      Top             =   350
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   164233217
      CurrentDate     =   40777
   End
   Begin MSComctlLib.StatusBar sbContratos 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   44
      Top             =   8430
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3176
            MinWidth        =   3176
            Text            =   "F3 - Agregar Zonas"
            TextSave        =   "F3 - Agregar Zonas"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "F4 - Agregar Alternativas"
            TextSave        =   "F4 - Agregar Alternativas"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3246
            MinWidth        =   3246
            Text            =   "F5 - Agregar Detalles"
            TextSave        =   "F5 - Agregar Detalles"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4657
            MinWidth        =   4657
            Text            =   "F6 - Agregar Amortizaciones"
            TextSave        =   "F6 - Agregar Amortizaciones"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10080
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratos.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratos.frx":0724
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtComplementario 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   7320
      Width           =   9015
   End
   Begin VB.TextBox txtImportante 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   6000
      Width           =   9015
   End
   Begin MSComctlLib.ListView lvData 
      Height          =   1095
      Left            =   1425
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1245
      Visible         =   0   'False
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1931
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame FraEvento 
      Caption         =   "Evento"
      Height          =   1095
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Width           =   9015
      Begin MSComCtl2.DTPicker dtpTermino 
         Height          =   255
         Left            =   5640
         TabIndex        =   5
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy hh:mm tt"
         Format          =   164233219
         CurrentDate     =   40773
      End
      Begin VB.ComboBox ComExclusivo 
         Height          =   315
         ItemData        =   "frmContratos.frx":0ABE
         Left            =   6480
         List            =   "frmContratos.frx":0AC8
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtAdultos 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtNinios 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         MaxLength       =   4
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy hh:mm tt"
         Format          =   164233219
         CurrentDate     =   40772
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TERMINO:"
         Height          =   195
         Left            =   4560
         TabIndex        =   43
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INICIO:"
         Height          =   195
         Left            =   315
         TabIndex        =   26
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADULTOS:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   645
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIÑOS:"
         Height          =   195
         Left            =   2640
         TabIndex        =   24
         Top             =   645
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EXCLUSIVO:"
         Height          =   195
         Left            =   5160
         TabIndex        =   23
         Top             =   645
         Width           =   1095
      End
   End
   Begin VB.Frame FraZonas 
      Caption         =   "Zonas"
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Width           =   9015
      Begin VB.Label lblZonas 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   8820
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   33
      Top             =   5160
      Width           =   9015
      Begin VB.TextBox txtAcuenta 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   4080
         TabIndex        =   13
         Text            =   "0"
         Top             =   160
         Width           =   1380
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1560
         TabIndex        =   38
         Top             =   165
         Width           =   1380
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6720
         TabIndex        =   37
         Top             =   165
         Width           =   1380
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
         Height          =   195
         Left            =   720
         TabIndex        =   36
         Top             =   220
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A CUENTA:"
         Height          =   195
         Left            =   3000
         TabIndex        =   35
         Top             =   220
         Width           =   960
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SALDO:"
         Height          =   195
         Left            =   6000
         TabIndex        =   34
         Top             =   225
         Width           =   675
      End
   End
   Begin VB.Frame FraCliente 
      Caption         =   "Cliente"
      Height          =   1575
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   9015
      Begin VB.TextBox txtCliente 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   7575
      End
      Begin VB.TextBox txtAtencion 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   4575
      End
      Begin VB.ComboBox ComTipoContrato 
         Height          =   315
         ItemData        =   "frmContratos.frx":0AD4
         Left            =   6480
         List            =   "frmContratos.frx":0AE4
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   705
         Width           =   2415
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE:"
         Height          =   195
         Left            =   390
         TabIndex        =   47
         Top             =   405
         Width           =   810
      End
      Begin VB.Label lblTelefonos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6480
         TabIndex        =   46
         Top             =   1080
         Width           =   2445
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO:"
         Height          =   195
         Left            =   6000
         TabIndex        =   45
         Top             =   765
         Width           =   495
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOCUMENTO:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1140
         Width           =   1200
      End
      Begin VB.Label lblDocumento 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1320
         TabIndex        =   18
         Top             =   1080
         Width           =   2700
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ATENCIÓN:"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TELEFONOS:"
         Height          =   195
         Left            =   5280
         TabIndex        =   16
         Top             =   1140
         Width           =   1080
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   635
      ButtonWidth     =   2011
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Grabar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraDetalle 
      Caption         =   "Detalle"
      Height          =   1335
      Left            =   120
      TabIndex        =   27
      Top             =   3840
      Width           =   9015
      Begin VB.Frame FraAlternativas 
         Height          =   570
         Left            =   120
         TabIndex        =   28
         Top             =   170
         Width           =   8775
         Begin MSDataListLib.DataCombo DatAlternativas 
            Height          =   315
            Left            =   1560
            TabIndex        =   9
            Top             =   180
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DatComposicion 
            Height          =   315
            Left            =   5760
            TabIndex        =   10
            Top             =   150
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ALTERNATIVAS:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COMPOSICIÓN:"
            Height          =   195
            Left            =   4320
            TabIndex        =   29
            Top             =   210
            Width           =   1380
         End
      End
      Begin VB.Frame FraExternos 
         Height          =   550
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   8775
         Begin VB.Label lblExternos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   840
            TabIndex        =   32
            Top             =   170
            Width           =   7095
         End
      End
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORTANTE:"
      Height          =   195
      Left            =   120
      TabIndex        =   42
      Top             =   5760
      Width           =   1200
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INCLUYE COMPLEMENTARIAMENTE:"
      Height          =   195
      Left            =   120
      TabIndex        =   41
      Top             =   7080
      Width           =   3060
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA IMPRESIÓN:"
      Height          =   195
      Left            =   5880
      TabIndex        =   40
      Top             =   405
      Width           =   1695
   End
End
Attribute VB_Name = "frmContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public VNuevo            As Boolean

Private orsDATA          As New ADODB.Recordset

Private vBuscar          As Boolean 'variable para la busqueda de clientes

Private loc_key          As Integer

Private vPUNTO           As Boolean

Public oRsZonas          As ADODB.Recordset 'recordset para almacenar las zonas del contrato

Public oRsAlternativas   As ADODB.Recordset

Public vGraba            As Boolean

Public oRsComposicion    As ADODB.Recordset

Public oRsExternos       As ADODB.Recordset

Public oRSAMORTIZACIONES As ADODB.Recordset

Public vIDContrato       As Integer

Public EstaAnulado       As Boolean

Private Sub ComFormatoHora_KeyPress(KeyAscii As Integer)

    If Asc(KeyAscii) = vbKeyReturn Then Me.txtADULTOS.SetFocus
End Sub

Private Sub ComExclusivo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Me.DatAlternativas.SetFocus
End Sub

Private Sub ComTipoContrato_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.dtpInicio.SetFocus
End Sub

Private Sub DatAlternativas_Click(Area As Integer)

    'MsgBox Me.DatAlternativas.BoundText
    If Me.DatAlternativas.BoundText = "" Then Exit Sub
    Me.DatComposicion.BoundText = ""

    If Not oRsComposicion.EOF And Not oRsComposicion.BOF Then
        oRsComposicion.Filter = "CODALTERNATIVA=" & Me.DatAlternativas.BoundText

        Dim ors As New ADODB.Recordset

        With ors
            .Fields.Append "CODPLATO", adBigInt
            .Fields.Append "PRODUCTO", adVarChar, 100
            .CursorLocation = adUseClient
            .LockType = adLockBatchOptimistic
            .CursorType = adOpenDynamic
            .Open
        End With

        If Not oRsComposicion.EOF Then

            Do While Not oRsComposicion.EOF
                ors.AddNew
                
                ors.Fields("CODPLATO") = oRsComposicion!CODPLATO
                ors.Fields("PRODUCTO") = oRsComposicion!PRODUCTO
                
                ors.Update
                oRsComposicion.MoveNext
            Loop

            ors.MoveFirst
            oRsComposicion.MoveFirst
        End If

        If Not ors.EOF Then
            Set Me.DatComposicion.RowSource = ors
            Me.DatComposicion.BoundColumn = "CODPLATO"
            Me.DatComposicion.ListField = "PRODUCTO"
        End If
    
    End If

End Sub

Private Sub DatAlternativas_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Me.DatComposicion.SetFocus
End Sub

Private Sub dtpFechaImpresion_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Me.txtCliente.SetFocus
End Sub

Private Sub dtpInicio_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Me.dtpTermino.SetFocus
End Sub

Private Sub dtpTermino_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Me.ComExclusivo.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF3 Then If Not EstaAnulado Then LlamarZonas
    If KeyCode = vbKeyF4 Then If Not EstaAnulado Then LlamarAlternativas
    If KeyCode = vbKeyF5 Then If Not EstaAnulado Then LlamarExternos
    If KeyCode = vbKeyEscape Then Unload Me
    If KeyCode = vbKeyF6 Then If Not EstaAnulado Then LlamarAmortizaciones
End Sub

Private Sub LlamarAmortizaciones()
    
    frmContratosAmortizaciones.Show vbModal
End Sub

Private Sub Form_Load()

    vGraba = False
    Me.txtAcuenta.Text = "0.00"
    Me.dtpFechaImpresion.Value = LK_FECHA_DIA
    Me.dtpInicio.Value = LK_FECHA_DIA
    Me.dtpTermino.Value = LK_FECHA_DIA
    ConfiguraLV
    Set oRsZonas = New ADODB.Recordset
    oRsZonas.Fields.Append "CODZONA", adInteger
    oRsZonas.Fields.Append "ZONA", adVarChar, 20
    oRsZonas.Fields.Append "ADULTOS", adDouble
    oRsZonas.Fields.Append "NINIOS", adDouble
    oRsZonas.CursorLocation = adUseClient
    oRsZonas.LockType = adLockBatchOptimistic
    oRsZonas.CursorType = adOpenDynamic
    oRsZonas.Open

    Set oRsAlternativas = New ADODB.Recordset

    With oRsAlternativas
        .Fields.Append "CODALTERNATIVA", adBigInt
        .Fields.Append "ALTERNATIVA", adVarChar, 90
        .Fields.Append "IDDSCTO", adInteger, , adFldIsNullable
        
        .Fields.Append "DESCUENTO", adDouble
        .Fields.Append "CANTIDAD", adDouble
        .Fields.Append "PRECIO", adDouble
        .Fields.Append "BRUTO", adDouble
        .Fields.Append "NETO", adDouble
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenDynamic
        .Open
    End With

    Set oRsComposicion = New ADODB.Recordset

    With oRsComposicion
        .Fields.Append "CODALTERNATIVA", adBigInt
        .Fields.Append "CODPLATO", adBigInt
        .Fields.Append "PRODUCTO", adVarChar, 100
        .Fields.Append "PRECIO", adDouble
        .Fields.Append "CANTIDAD", adInteger
        .Fields.Append "IMPORTE", adDouble
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenDynamic
        .Open
    End With

    Set oRsExternos = New ADODB.Recordset

    With oRsExternos
        .Fields.Append "CODEXTERNO", adBigInt
        .Fields.Append "DESCRIPCION", adVarChar, 150
        .Fields.Append "CANTIDAD", adInteger
        .Fields.Append "PRECIO", adDouble
        .Fields.Append "IMPORTE", adDouble
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenDynamic
        .Open
    
    End With
        
    Set oRSAMORTIZACIONES = New ADODB.Recordset

    With oRSAMORTIZACIONES
        .Fields.Append "MONEDA", adChar, 1
        .Fields.Append "MONTO", adDouble
        .Fields.Append "IDFORMAPAGO", adDouble
        .Fields.Append "FORMAPAGO", adVarChar, 80
        .Fields.Append "FECHAPAGO", adDBTimeStamp
        .Fields.Append "IDAMORTIZACION", adDouble
        .Open

    End With

    If Not VNuevo Then
        Me.txtAcuenta.Enabled = False

        'carga la info del contrato seleccionado
        Dim orsGen  As ADODB.Recordset

        Dim oRsTemp As ADODB.Recordset

        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPCONTRATObyCORRELATIVO"

        With oCmdEjec
            .Parameters.Append .CreateParameter("@IDCONTRATO", adBigInt, adParamInput, , frmContratosList2.contextEvent.ScheduleID)
            .Parameters.Append .CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            Set orsGen = .Execute
        End With
        
        Me.lblTotal.Caption = orsGen!Total ' frmContratosList.lvData.SelectedItem.SubItems(5)
        Me.txtAcuenta.Text = orsGen!ACUENTA ' frmContratosList.lvData.SelectedItem.SubItems(6)
        Me.lblSaldo.Caption = orsGen!SALDO ' frmContratosList.lvData.SelectedItem.SubItems(7)
        
        Me.txtADULTOS.Text = orsGen!adultos
        Me.txtNINIOS.Text = orsGen!ninios
        Me.ComExclusivo.ListIndex = orsGen!EXCLUSIVO
        Me.txtImportante.Text = IIf(IsNull(orsGen!IMPORTANTE), "", Trim(orsGen!IMPORTANTE))
        Me.txtComplementario.Text = IIf(IsNull(orsGen!COMPLEMENTARIO), "", Trim(orsGen!COMPLEMENTARIO))
        Me.txtCliente.Text = Trim(orsGen!CLIENTE)
        Me.txtCliente.Tag = orsGen!IDCLIENTE
        Me.dtpFechaImpresion.Value = orsGen!FECHACONTRATO

        If orsGen!EXCLUSIVO = 0 Then
            Me.ComExclusivo.ListIndex = 0
        Else
            Me.ComExclusivo.ListIndex = 1
        End If

        Me.lblDocumento.Caption = orsGen!DOCTO
        Me.lblTelefonos.Caption = orsGen!TELEFONOS
        'Me.lblAtencion.Caption = orsGen!ATENCION
        Me.txtAtencion.Text = orsGen!ATENCION
        
        Me.dtpInicio.Value = orsGen!INICIOEVENTO
        Me.dtpTermino.Value = orsGen!TERMINOEVENTO
        If Not IsNull(orsGen!tipo) Then
        Me.ComTipoContrato.ListIndex = orsGen!tipo
        
        End If
        
        'CARGANDO LAS ZONAS ASOCIADAS AL CONTRATO
        Set oRsTemp = orsGen.NextRecordset

        Do While Not oRsTemp.EOF
            oRsZonas.AddNew
            oRsZonas!codzona = oRsTemp!codzona
            oRsZonas!ZONA = oRsTemp!denomina
            oRsZonas!adultos = oRsTemp!adultos
            oRsZonas!ninios = oRsTemp!ninios
            oRsZonas.Update
            oRsTemp.MoveNext
        Loop

        Dim VZONA As String

        If oRsZonas.Fields.count > 0 Then
            oRsZonas.MoveFirst

            Do While Not oRsZonas.EOF
                VZONA = VZONA + oRsZonas.Fields("ZONA").Value + " - "
                oRsZonas.MoveNext
            Loop

            Me.lblzonas.Caption = Left(VZONA, Len(Trim(VZONA)) - 2)
            oRsZonas.MoveFirst
        End If
    
        'Set oRsZonas = orsTemp
        
        'CARGANDO LAS ALTERNATIVAS DEL CONTRATO
        Set oRsTemp = orsGen.NextRecordset
      
        Do While Not oRsTemp.EOF

            With oRsAlternativas
                .AddNew
                .Fields!CODALTERNATIVA = oRsTemp!CODALTERNATIVA
                .Fields!ALTERNATIVA = oRsTemp!ALTERNATIVA
                
                .Fields!IDDSCTO = oRsTemp!IDEDSCTO
                .Fields!DESCUENTO = IIf(IsNull(oRsTemp!DESCUENTO), 0, oRsTemp!DESCUENTO)
                .Fields!Cantidad = oRsTemp!Cantidad
                .Fields!PRECIO = oRsTemp!PRECIO
                .Fields!NETO = oRsTemp!NETO
                .Fields!BRUTO = oRsTemp!BRUTO
                .Update
            End With

            oRsTemp.MoveNext
        Loop

        If oRsAlternativas.RecordCount <> 0 Then
            oRsAlternativas.Filter = ""
            oRsAlternativas.MoveFirst
        End If

        'CARGANDO LAS COMPOSICIONES DEL CONTRATO
        Set oRsTemp = orsGen.NextRecordset
 
        Do While Not oRsTemp.EOF

            With oRsComposicion
                .AddNew
                .Fields!CODALTERNATIVA = oRsTemp!CODALTERNATIVA
                .Fields!CODPLATO = oRsTemp!CODPLATO
                .Fields!PRODUCTO = oRsTemp!PLATO
                .Fields!Cantidad = oRsTemp!Cantidad
                .Fields!PRECIO = oRsTemp!PRECIO
                .Fields!Importe = oRsTemp!Importe
                .Update
            End With

            oRsTemp.MoveNext
        Loop
        
        Set oRsTemp = orsGen.NextRecordset

        Do While Not oRsTemp.EOF

            With oRsExternos
                .AddNew
                .Fields!CODEXTERNO = oRsTemp!CODEXTERNO
                .Fields!DESCRIPCION = oRsTemp!EXTERNO
                .Fields!Cantidad = oRsTemp!Cantidad
                .Fields!PRECIO = oRsTemp!PRECIO
                .Fields!Importe = oRsTemp!Importe
                .Update
            End With

            oRsTemp.MoveNext
        
        Loop
        
''''        'CARGANDO LAS AMORTIZACIONES DEL CONTRATO
''''        Set oRsTemp = orsGen.NextRecordset
''''
''''        If Not oRsTemp.EOF Then
''''
''''            Do While Not oRsTemp.EOF
''''
''''                With oRSAMORTIZACIONES
''''                    .AddNew
''''                    .Fields!moneda = oRsTemp!moneda
''''                    .Fields!MONTO = oRsTemp!MONTO
''''                    .Fields!IDFORMAPAGO = oRsTemp!IDFORMAPAGO
''''                    .Fields!FORMAPAGO = oRsTemp!FORMAPAGO
''''                    .Fields!FECHAPAGO = oRsTemp!FECHAPAGO
''''                    .Fields!IDAMORTIZACION = oRsTemp!IDAMORTIZACION
''''                    .Update
''''                End With
''''
''''                oRsTemp.MoveNext
''''            Loop
''''
''''            oRSAMORTIZACIONES.MoveFirst
''''        End If

        'SEPARANDO LA INFORMACION
        If oRsExternos.RecordCount > 0 Then
            Me.lblExternos.Caption = oRsExternos.RecordCount & " ITEMS AGREGADOS"

            If oRsExternos.RecordCount <> 0 Then
                oRsExternos.Filter = ""
                oRsExternos.MoveFirst
            End If

        Else
            Me.lblExternos.Caption = oRsExternos.RecordCount & " ITEMS AGREGADOS"
        End If
        
        
         If frmContratosList2.contextEvent.label = 1 Or frmContratosList2.contextEvent.label = 2 Then
        Me.FraAlternativas.Enabled = False
        Me.FraCliente.Enabled = False
        Me.dtpFechaImpresion.Enabled = False
        Me.FraEvento.Enabled = False
        Me.FraZonas.Enabled = False
        Me.txtImportante.Enabled = False
        Me.txtComplementario.Enabled = False
        Me.Toolbar1.Enabled = False
    End If
    End If

   

End Sub

'Private Sub FraDetalle_DragDrop(Source As Control, x As Single, y As Single)
'Alternativas:
'End Sub

Private Sub MasFecha_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Me.dtpInicio.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error GoTo Graba

    Select Case Button.Index

        Case 1

            If Not IsDate(Me.dtpFechaImpresion.Value) Then
                MsgBox "Fecha incorrecta para el Contrato.", vbCritical, Pub_Titulo
                Me.dtpFechaImpresion.SetFocus

                Exit Sub

            End If

            If Len(Trim(Me.txtCliente.Tag)) = 0 Then
                MsgBox "Debe ingresar el cliente respectivo.", vbCritical, Pub_Titulo
                Me.txtCliente.SelStart = 0
                Me.txtCliente.SelLength = Len(Me.txtCliente.Text)
                Me.txtCliente.SetFocus

                Exit Sub

            End If

            If Not IsDate(Me.dtpInicio.Value) Then
                MsgBox "Fecha de Evento incorrecta para el Contrato.", vbCritical, Pub_Titulo
                Me.dtpInicio.SetFocus

                Exit Sub

            End If

            If Len(Trim(Me.txtADULTOS.Text)) = 0 And Len(Trim(Me.txtNINIOS.Text)) = 0 Then
                MsgBox "Debe ingresar la Cantida de Adultos o Niños para el Contrato.", vbCritical, Pub_Titulo

                If Len(Trim(Me.txtADULTOS.Text)) = 0 Then
                    Me.txtADULTOS.SelStart = 0
                    Me.txtADULTOS.SelLength = Len(Me.txtADULTOS.Text)
                    Me.txtADULTOS.SetFocus
                ElseIf Len(Trim(Me.txtNINIOS.Text)) = 0 Then
                    Me.txtNINIOS.SelStart = 0
                    Me.txtNINIOS.SelLength = Len(Me.txtNINIOS.Text)
                    Me.txtNINIOS.SetFocus
                End If

                Exit Sub

            End If

            If Me.ComExclusivo.ListIndex = -1 Then
                MsgBox "Debe seleccionar la Exclusividad del Contrato.", vbCritical, Pub_Titulo
                Me.ComExclusivo.SetFocus

                Exit Sub

            End If

            If oRsZonas.RecordCount = 0 Then
                MsgBox "Debe agregar zonas para el Contrato.", vbCritical, Pub_Titulo
                LlamarZonas

                Exit Sub

            End If

            If oRsAlternativas.RecordCount = 0 Then
                MsgBox "Debe agregar información de Alternativas para el contrato.", vbCritical, Pub_Titulo
                LlamarAlternativas

                Exit Sub

            End If
        
            If oRsComposicion.RecordCount = 0 Then
                MsgBox "Debe agregar la composición de las Alternativas.", vbCritical, Pub_Titulo

                Exit Sub

            End If

            If val(Me.lblSaldo.Caption) < 0 Then
                MsgBox "El Importe Acuenta debe ser menor al total del Contrato.", vbCritical, Pub_Titulo
                Me.txtAcuenta.SelStart = 0
                Me.txtAcuenta.SelLength = Len(Me.txtAcuenta.Text)
                Me.txtAcuenta.SetFocus

                Exit Sub

            End If

            'LLENANDO LAS AMORTIZACIONES DEL CONTRATO
            If Not oRSAMORTIZACIONES.BOF Or Not oRSAMORTIZACIONES.EOF Then
                oRSAMORTIZACIONES.MoveFirst
            End If

            Dim xPAGOS As String

            If Not oRSAMORTIZACIONES.EOF Then
                xPAGOS = "<r>"

                Do While Not oRSAMORTIZACIONES.EOF
                    xPAGOS = xPAGOS & "<d "
                    xPAGOS = xPAGOS & "ida= """ & oRSAMORTIZACIONES.Fields!IDAMORTIZACION & """ "
                    xPAGOS = xPAGOS & "mn = """ & oRSAMORTIZACIONES.Fields!moneda & """ "
                    xPAGOS = xPAGOS & "mt = """ & oRSAMORTIZACIONES.Fields!MONTO & """ "
                    xPAGOS = xPAGOS & "ifp = """ & oRSAMORTIZACIONES.Fields!IDFORMAPAGO & """ "
                    xPAGOS = xPAGOS & "fp = """ & Year(oRSAMORTIZACIONES.Fields!FECHAPAGO) & Right("00" & Month(oRSAMORTIZACIONES.Fields!FECHAPAGO), 2) & Right("00" & Day(oRSAMORTIZACIONES.Fields!FECHAPAGO), 2) & """"
                    xPAGOS = xPAGOS & " />"
                    oRSAMORTIZACIONES.MoveNext
                Loop

                xPAGOS = xPAGOS & "</r>"
            End If

            LimpiaParametros oCmdEjec

            Dim xZONAS   As String

            Dim ORSDATO  As ADODB.Recordset

            Dim vENC     As Boolean

            Dim vmensaje As String


            If VNuevo Then 'NUEVO CONTRATO
                oCmdEjec.CommandText = "SPVERIFICACONTRATOREGISTRO"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHAINICIO", adDBTimeStamp, adParamInput, , Me.dtpInicio.Value)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHATERMINO", adDBTimeStamp, adParamInput, , Me.dtpTermino.Value)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 0)
                oCmdEjec.Execute

                vENC = oCmdEjec.Parameters("@PASA").Value

                If vENC Then
                    MsgBox "Hay Contratos que se cruzan con las fechas proporcionadas.", vbCritical, TituloSistema
                    Me.dtpInicio.SetFocus

                    Exit Sub

                End If
                
                'VALIDACION DE CAPACIDAD EN ZONAS POR COMPAÑIA

                If Not oRsZonas.EOF Or Not oRsZonas.BOF Then
                    oRsZonas.MoveFirst
                End If
                
                If Not oRsZonas.EOF Then
                    xZONAS = "<r>"

                    Do While Not oRsZonas.EOF
                        xZONAS = xZONAS + "<d "
                        xZONAS = xZONAS + "idz=""" & oRsZonas!codzona & """ "
                        xZONAS = xZONAS + "zon=""" & oRsZonas!ZONA & """ "
                        xZONAS = xZONAS + "cn=""" & oRsZonas!adultos + oRsZonas!ninios & """ "
                        xZONAS = xZONAS + " />"

                        oRsZonas.MoveNext
                    Loop

                    xZONAS = xZONAS + "</r>"
                End If
            
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SP_VALIDAREGISTROCONTRATO"
                oCmdEjec.CommandType = adCmdStoredProc
                
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA1", adDBTimeStamp, adParamInput, , Me.dtpInicio.Value)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA2", adDBTimeStamp, adParamInput, , Me.dtpTermino.Value)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@xZONAS", adBSTR, adParamInput, 8000, xZONAS)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MENSAJE", adBSTR, adParamOutput, 4000, vmensaje)

                Set ORSDATO = oCmdEjec.Execute
                vmensaje = Trim(oCmdEjec.Parameters(4).Value)
                
                If ORSDATO!INFO = True Then
                    If Len(Trim(vmensaje)) = 0 Then
                
                        MsgBox "No se puede registrar el contrato." & vbCrLf & "Debido a que alguna zona esta en el tope de su capacidad.", vbInformation, Pub_Titulo

                        Exit Sub

                    Else
                        MsgBox "No se puede registrar el contrato." & vbCrLf & "Debido a que las siguientes zonas estan en su capacidad maxima." + vbCrLf + vmensaje, vbInformation, Pub_Titulo

                        Exit Sub

                    End If
                End If

                Pub_ConnAdo.BeginTrans
                LimpiaParametros oCmdEjec

                'GRABANDO NUEVO EN CONTRATO
                With oCmdEjec
                    .CommandText = "SPREGISTRARCONTRATO"
                    .Parameters.Append .CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                    .Parameters.Append .CreateParameter("@IDCLIENTE", adInteger, adParamInput, , Me.txtCliente.Tag)
                    .Parameters.Append .CreateParameter("@INICIOEVENTO", adDBTimeStamp, adParamInput, , Me.dtpInicio.Value)
                    .Parameters.Append .CreateParameter("@TERMINOEVENTO", adDBTimeStamp, adParamInput, , Me.dtpTermino.Value)
                    .Parameters.Append .CreateParameter("@FECHACONTRATO", adDBTimeStamp, adParamInput, , Me.dtpFechaImpresion.Value)
                    .Parameters.Append .CreateParameter("@ADULTOS", adInteger, adParamInput, , Me.txtADULTOS.Text)
                    .Parameters.Append .CreateParameter("@NINIOS", adInteger, adParamInput, , IIf(Len(Trim(Me.txtNINIOS.Text)) = 0, 0, Me.txtNINIOS.Text))
                    .Parameters.Append .CreateParameter("@EXCLUSIVO", adBoolean, adParamInput, , Me.ComExclusivo.ListIndex)
                    .Parameters.Append .CreateParameter("@TOTAL", adDouble, adParamInput, , Me.lblTotal.Caption)
                    .Parameters.Append .CreateParameter("@ACUENTA", adDouble, adParamInput, , Me.txtAcuenta.Text)
                    .Parameters.Append .CreateParameter("@SALDO", adDouble, adParamInput, , Me.lblSaldo.Caption)
                    .Parameters.Append .CreateParameter("@IMPORTANTE", adVarChar, adParamInput, 1000, Me.txtImportante.Text)
                    .Parameters.Append .CreateParameter("@COMPLEMENTARIO", adVarChar, adParamInput, 1000, Me.txtComplementario.Text)
                    .Parameters.Append .CreateParameter("@xPAGOS", adBSTR, adParamInput, 8000, xPAGOS)
                    .Parameters.Append .CreateParameter("@ESTADO", adChar, adParamInput, 1, "V")
                    .Parameters.Append .CreateParameter("@ATENCION", adVarChar, adParamInput, 60, Me.txtAtencion.Text)
                    .Parameters.Append .CreateParameter("@TIPO", adTinyInt, adParamInput, 1, Me.ComExclusivo.ListIndex)
                    .Parameters.Append .CreateParameter("@IDCONTRATO", adBigInt, adParamOutput, , 0)
                    .Execute
                    vIDContrato = .Parameters("@IDCONTRATO").Value
                End With

            Else
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SPVERIFICACONTRATOMODIFICO"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adInteger, adParamInput, , vIDContrato)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHAINICIO", adDBTimeStamp, adParamInput, , Me.dtpInicio.Value)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHATERMINO", adDBTimeStamp, adParamInput, , Me.dtpTermino.Value)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 0)
                oCmdEjec.Execute

                vENC = oCmdEjec.Parameters("@PASA").Value

                If vENC Then
                    MsgBox "Hay Contratos que se cruzan con las fechas proporcionadas.", vbCritical, Pub_Titulo
                    Me.dtpInicio.SetFocus

                    Exit Sub

                End If
                
                'VALIDACION DE CAPACIDAD EN ZONAS POR COMPAÑIA

                If Not oRsZonas.EOF Or Not oRsZonas.BOF Then
                    oRsZonas.MoveFirst
                End If
                
                If Not oRsZonas.EOF Then
                    xZONAS = "<r>"

                    Do While Not oRsZonas.EOF
                        xZONAS = xZONAS + "<d "
                        xZONAS = xZONAS + "idz=""" & oRsZonas!codzona & """ "
                        xZONAS = xZONAS + "zon=""" & oRsZonas!ZONA & """ "
                        xZONAS = xZONAS + "cn=""" & oRsZonas!adultos + oRsZonas!ninios & """ "
                        xZONAS = xZONAS + " />"

                        oRsZonas.MoveNext
                    Loop

                    xZONAS = xZONAS + "</r>"
                End If
            
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SP_VALIDAREGISTROCONTRATO_UPDATE"
                oCmdEjec.CommandType = adCmdStoredProc
                
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adDouble, adParamInput, , vIDContrato)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA1", adDBTimeStamp, adParamInput, , Me.dtpInicio.Value)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA2", adDBTimeStamp, adParamInput, , Me.dtpTermino.Value)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@xZONAS", adBSTR, adParamInput, 8000, xZONAS)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MENSAJE", adBSTR, adParamOutput, 4000, vmensaje)

                Set ORSDATO = oCmdEjec.Execute
                vmensaje = Trim(oCmdEjec.Parameters(5).Value)
                
                If ORSDATO!INFO = True Then
                    If Len(Trim(vmensaje)) <> 0 Then
                        MsgBox "No se puede registrar el contrato." & vbCrLf & "Las siguientes zonas estan al tope:" & vbclrf & vmensaje, vbInformation, Pub_Titulo

                        Exit Sub

                    Else
                        MsgBox "No se puede registrar el contrato." & vbCrLf & "Debido a que alguna zona esta en el tope de su capacidad.", vbInformation, Pub_Titulo

                        Exit Sub

                    End If

                End If

                LimpiaParametros oCmdEjec
                Pub_ConnAdo.BeginTrans

                'MODIFICANDO EN LA TABLA CONTRATO
                With oCmdEjec
                    .CommandText = "SPMODIFICARCONTRATO"
                    .Parameters.Append .CreateParameter("@IDCONTRATO", adBigInt, adParamInput, , vIDContrato)
                    .Parameters.Append .CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                    .Parameters.Append .CreateParameter("@IDCLIENTE", adInteger, adParamInput, , Me.txtCliente.Tag)
                    .Parameters.Append .CreateParameter("@INICIOEVENTO", adDBTimeStamp, adParamInput, , Me.dtpInicio.Value)
                    .Parameters.Append .CreateParameter("@TERMINOEVENTO", adDBTimeStamp, adParamInput, , Me.dtpTermino.Value)
                    .Parameters.Append .CreateParameter("@FECHACONTRATO", adDBTimeStamp, adParamInput, , Me.dtpFechaImpresion.Value)
                    .Parameters.Append .CreateParameter("@ADULTOS", adInteger, adParamInput, , Me.txtADULTOS.Text)
                    .Parameters.Append .CreateParameter("@NINIOS", adInteger, adParamInput, , IIf(Len(Trim(Me.txtNINIOS.Text)) = 0, 0, Me.txtNINIOS.Text))
                    .Parameters.Append .CreateParameter("@EXCLUSIVO", adBoolean, adParamInput, , Me.ComExclusivo.ListIndex)
                    .Parameters.Append .CreateParameter("@TOTAL", adDouble, adParamInput, , Me.lblTotal.Caption)
                    .Parameters.Append .CreateParameter("@ACUENTA", adDouble, adParamInput, , Me.txtAcuenta.Text)
                    .Parameters.Append .CreateParameter("@SALDO", adDouble, adParamInput, , Me.lblSaldo.Caption)
                    .Parameters.Append .CreateParameter("@ATENCION", adVarChar, adParamInput, 60, Me.txtAtencion.Text)
                    .Parameters.Append .CreateParameter("@TIPO", adTinyInt, adParamInput, , Me.ComTipoContrato.ListIndex)
                    '.Parameters.Append .CreateParameter("@xPAGOS", adBSTR, adParamInput, 8000, xPAGOS)

                    If Len(Trim(Me.txtImportante.Text)) <> 0 Then .Parameters.Append .CreateParameter("@IMPORTANTE", adVarChar, adParamInput, 1000, Me.txtImportante.Text)
                    If Len(Trim(Me.txtComplementario.Text)) <> 0 Then .Parameters.Append .CreateParameter("@COMPLEMENTARIO", adVarChar, adParamInput, 1000, Me.txtComplementario.Text)
                    
                    .Execute
                End With

            End If

            LimpiaParametros oCmdEjec

            If Not VNuevo Then
                oCmdEjec.CommandText = "SPELIMINARCONTRATO"

                With oCmdEjec
                    .Parameters.Append .CreateParameter("@IDCONTRATO", adBigInt, adParamInput, , vIDContrato)
                    .Parameters.Append .CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                    .Execute
                End With

                LimpiaParametros oCmdEjec
            End If
            
            'PROCESANDO TABLA CONTRATOS_ZONAS
            
            If oRsZonas.RecordCount <> 0 Then
                oRsZonas.Filter = ""
                oRsZonas.MoveFirst
            End If

            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SPREGISTRARCONTRATOZONAS"
            
            Do While Not oRsZonas.EOF

                With oCmdEjec
                    .Parameters.Append .CreateParameter("@IDCONTRATO", adBigInt, adParamInput, , vIDContrato)
                    .Parameters.Append .CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                    .Parameters.Append .CreateParameter("@CODZONA", adInteger, adParamInput, , oRsZonas!codzona)
                    .Parameters.Append .CreateParameter("@ADULTOS", adDouble, adParamInput, , oRsZonas!adultos)
                    .Parameters.Append .CreateParameter("@NINIOS", adDouble, adParamInput, , oRsZonas!ninios)
                    .Execute
                    LimpiaParametros oCmdEjec
                End With

                oRsZonas.MoveNext
            Loop
            
            LimpiaParametros oCmdEjec

            'PROCESANDO TABLA CONTRATOS_ALTERANATIVAS
            If oRsComposicion.RecordCount <> 0 Then
                oRsAlternativas.Filter = ""
                oRsComposicion.Filter = ""
                oRsAlternativas.MoveFirst
                oRsComposicion.MoveFirst
            End If
            
            LimpiaParametros oCmdEjec
            'PROCESANDO TABLA CONTRATOS_ALTERNATIVAS
            oCmdEjec.CommandText = "SPREGISTRARCONTRATOALTERNATIVA"
                          
            Do While Not oRsAlternativas.EOF

                With oCmdEjec
                    .Parameters.Append .CreateParameter("IDCONTRATO", adBigInt, adParamInput, , vIDContrato)
                    .Parameters.Append .CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                    .Parameters.Append .CreateParameter("@IDALTERNATIVA", adInteger, adParamInput, , oRsAlternativas!CODALTERNATIVA)
                    .Parameters.Append .CreateParameter("@ALTERNATIVA", adVarChar, adParamInput, 90, oRsAlternativas!ALTERNATIVA)
                    .Parameters.Append .CreateParameter("@IDDSCTO", adInteger, adParamInput, , oRsAlternativas!IDDSCTO)
                    .Parameters.Append .CreateParameter("@DESCUENTO", adDouble, adParamInput, , oRsAlternativas!DESCUENTO)
                    .Parameters.Append .CreateParameter("@CANTIDAD", adInteger, adParamInput, , oRsAlternativas!Cantidad)
                    .Parameters.Append .CreateParameter("@PRECIO", adDouble, adParamInput, , oRsAlternativas!PRECIO)
                    .Parameters.Append .CreateParameter("@NETO", adDouble, adParamInput, , oRsAlternativas!NETO)
                    .Parameters.Append .CreateParameter("@BRUTO", adDouble, adParamInput, , oRsAlternativas!BRUTO)
                    .Execute
                End With

                LimpiaParametros oCmdEjec
                oRsAlternativas.MoveNext
            Loop
            
            LimpiaParametros oCmdEjec
            
            oCmdEjec.CommandText = "SPREGISTRARCONTRATOEXTERNO"
                          
            Do While Not oRsExternos.EOF

                With oCmdEjec
                    .Parameters.Append .CreateParameter("@IDCONTRATO", adBigInt, adParamInput, , vIDContrato)
                    .Parameters.Append .CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                    .Parameters.Append .CreateParameter("@IDEXTERNO", adInteger, adParamInput, , oRsExternos!CODEXTERNO)
                    .Parameters.Append .CreateParameter("@EXTERNO", adVarChar, adParamInput, 120, oRsExternos!DESCRIPCION)
                    .Parameters.Append .CreateParameter("@CANTIDAD", adInteger, adParamInput, , oRsExternos!Cantidad)
                    .Parameters.Append .CreateParameter("@PRECIO", adDouble, adParamInput, , oRsExternos!PRECIO)
                    .Parameters.Append .CreateParameter("@IMPORTE", adDouble, adParamInput, , oRsExternos!Importe)
                    .Execute
                End With

                LimpiaParametros oCmdEjec
                oRsExternos.MoveNext
            Loop

            oCmdEjec.CommandText = "SPREGISTRARCONTRATOCOMPOSICION"
            
            Do While Not oRsComposicion.EOF

                With oCmdEjec
                    .Parameters.Append .CreateParameter("@IDCONTRATO", adBigInt, adParamInput, , vIDContrato)
                    .Parameters.Append .CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                    .Parameters.Append .CreateParameter("@IDALTERNATIVA", adInteger, adParamInput, , oRsComposicion.Fields!CODALTERNATIVA)
                    .Parameters.Append .CreateParameter("@CODPLATO", adBigInt, adParamInput, , oRsComposicion.Fields!CODPLATO)
                    .Parameters.Append .CreateParameter("@CANTIDAD", adInteger, adParamInput, , oRsComposicion.Fields!Cantidad)
                    .Parameters.Append .CreateParameter("@PRECIO", adDouble, adParamInput, , oRsComposicion.Fields!PRECIO)
                    .Parameters.Append .CreateParameter("@IMPORTE", adDouble, adParamInput, , oRsComposicion.Fields!Importe)
                    .Execute
                End With

                LimpiaParametros oCmdEjec
                oRsComposicion.MoveNext
            Loop
          
            LimpiaParametros oCmdEjec
            
            If oRsZonas.RecordCount > 0 Then oRsAlternativas.MoveFirst
            If oRsAlternativas.RecordCount > 0 Then oRsZonas.MoveFirst
            If oRsComposicion.RecordCount > 0 Then oRsComposicion.MoveFirst
            ' If oRsExternos.RecordCount > 0 Then oRsExternos.MoveFirst
            Pub_ConnAdo.CommitTrans
            vGraba = True
            MsgBox "Datos Almacenados correctamente", vbInformation, Pub_Titulo
            Unload Me

        Case 2
            Unload Me
    End Select

    Exit Sub

Graba:

    Pub_ConnAdo.RollbackTrans
    
    MsgBox Err.Description
End Sub

Private Sub txtAcuenta_Change()

    If InStr(txtAcuenta.Text, ".") = 0 Then
        vPUNTO = False
    Else
        vPUNTO = True
    End If
    
    If IsNumeric(Me.txtAcuenta.Text) Then
        Me.lblSaldo.Caption = val(Me.lblTotal.Caption) - val(Me.txtAcuenta.Text)
    End If

End Sub

Private Sub txtAcuenta_GotFocus()

    If InStr(txtAcuenta.Text, ".") = 0 Then
        vPUNTO = False
    Else
        vPUNTO = True
    End If

End Sub

Private Sub txtAcuenta_KeyPress(KeyAscii As Integer)

    If NumerosyPunto(KeyAscii) Then KeyAscii = 0
    If Chr(KeyAscii) = "." Then
        If vPUNTO = True Or Len(Trim(txtAcuenta.Text)) = 0 Then
         
            KeyAscii = 0
         
        End If
    End If

    If KeyAscii = vbKeyReturn Then RestaAcuenta
End Sub

Private Sub txtADULTOS_KeyPress(KeyAscii As Integer)

    If SoloNumeros(KeyAscii) Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Me.txtNINIOS.SetFocus
End Sub

Private Sub txtAtencion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.ComTipoContrato.SetFocus
End Sub

Private Sub txtCliente_Change()
    vBuscar = True
    Me.txtCliente.Tag = ""
    Me.lblDocumento.Caption = ""
    Me.lblTelefonos.Caption = ""
    
    Me.txtAtencion.Text = ""
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key = loc_key + 1

        If loc_key > lvData.ListItems.count Then loc_key = lvData.ListItems.count
        GoTo POSICION
    End If

    If KeyCode = 38 Then
        loc_key = loc_key - 1

        If loc_key < 1 Then loc_key = 1
        GoTo POSICION
    End If

    If KeyCode = 34 Then
        loc_key = loc_key + 17

        If loc_key > lvData.ListItems.count Then loc_key = lvData.ListItems.count
        GoTo POSICION
    End If

    If KeyCode = 33 Then
        loc_key = loc_key - 17

        If loc_key < 1 Then loc_key = 1
        GoTo POSICION
    End If

    If KeyCode = 27 Then
        Me.lvData.Visible = False
        Me.txtCliente.Text = ""
        Me.lblDocumento.Caption = ""
        Me.lblTelefonos.Caption = ""
    End If

    GoTo fin
POSICION:
    lvData.ListItems.Item(loc_key).Selected = True
    lvData.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    Me.txtCliente.SelStart = Len(Me.txtCliente.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub ConfiguraLV()

    With Me.lvData

        .ColumnHeaders.Add , , "CLIENTE", 5000
        .ColumnHeaders.Add , , "DOCTO", 0
        .ColumnHeaders.Add , , "ATENCION", 0
        .ColumnHeaders.Add , , "TELEFONOS", 0
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual

    End With

End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = 13 Then
        If Len(Trim(Me.txtCliente.Text)) = 0 Then Exit Sub
        If vBuscar Then
            Me.lvData.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SPDATOSCLIENTESCONTRATO"
            Set orsDATA = oCmdEjec.Execute(, Array(Me.txtCliente.Text, LK_CODCIA))

            Dim Item As Object
        
            If Not orsDATA.EOF Then

                Do While Not orsDATA.EOF
                    Set Item = Me.lvData.ListItems.Add(, , orsDATA!CLIENTE)
                    Item.Tag = Trim(orsDATA!Codigo)
                    Item.SubItems(1) = orsDATA!DOCTO
                    Item.SubItems(2) = Trim(orsDATA!ATENCION)
                    Item.SubItems(3) = Trim(orsDATA!TELEFONOS)
                    orsDATA.MoveNext
                Loop

                Me.lvData.Visible = True
                Me.lvData.ListItems(1).Selected = True
                loc_key = 1
                Me.lvData.ListItems(1).EnsureVisible
                vBuscar = False
            
                '         If MsgBox("Cliente no existe." + vbCrLf + "¿Desea Crearlo.?", vbQuestion + vbYesNo, "Restaurantes") = vbYes Then
                '         frmCLI.Show vbModal
                '         End If
            Else
                Me.lvData.Visible = False
            End If

        Else 'ASIGNAR VALORES A LOS CONTROLES
            Me.lvData.Visible = False
            Me.txtCliente.Text = Me.lvData.SelectedItem.Text
            Me.txtCliente.Tag = Me.lvData.SelectedItem.Tag
            Me.lblDocumento.Caption = Me.lvData.SelectedItem.SubItems(1)
            Me.txtAtencion.Text = Me.lvData.SelectedItem.SubItems(2)
            Me.lblTelefonos.Caption = Me.lvData.SelectedItem.SubItems(3)
            Me.txtAtencion.SetFocus
        End If
    End If

End Sub

Private Sub LlamarZonas()

    If Not oRsZonas.BOF Or Not oRsZonas.EOF Then oRsZonas.MoveFirst
    If Not oRsZonas.EOF And Not oRsZonas.BOF Then
        Set frmContratosZonas.orsMarcados = oRsZonas.Clone
    Else
        Set frmContratosZonas.orsMarcados = Nothing
    End If
    
    ' If Not oRsZonas.BOF Or Not oRsZonas.EOF Then oRsZonas.MoveFirst

    frmContratosZonas.Show vbModal

    'frmContratosZonas.Show vbModal

End Sub

Private Sub LlamarAlternativas()

    'If Not oRsAlternativas.EOF And Not oRsAlternativas.BOF Then
    
    If frmContratosAlternativas.oRSAlt Is Nothing Then
        Set frmContratosAlternativas.oRSAlt = New ADODB.Recordset
        frmContratosAlternativas.oRSAlt.Fields.Append "CODALTERNATIVA", adBigInt
        frmContratosAlternativas.oRSAlt.Fields.Append "ALTERNATIVA", adVarChar, 90
        frmContratosAlternativas.oRSAlt.Fields.Append "IDDSCTO", adInteger, , adFldIsNullable
        frmContratosAlternativas.oRSAlt.Fields.Append "DESCUENTO", adDouble
        frmContratosAlternativas.oRSAlt.Fields.Append "CANTIDAD", adDouble
        frmContratosAlternativas.oRSAlt.Fields.Append "PRECIO", adDouble
        frmContratosAlternativas.oRSAlt.Fields.Append "NETO", adDouble
        frmContratosAlternativas.oRSAlt.Fields.Append "BRUTO", adDouble
        frmContratosAlternativas.oRSAlt.CursorLocation = adUseClient
        frmContratosAlternativas.oRSAlt.LockType = adLockBatchOptimistic
        frmContratosAlternativas.oRSAlt.CursorType = adOpenDynamic
        frmContratosAlternativas.oRSAlt.Open
    Else

        If frmContratosAlternativas.oRSAlt.RecordCount > 0 Then
            frmContratosAlternativas.oRSAlt.Filter = ""
            frmContratosAlternativas.oRSAlt.MoveFirst
        End If

        Do While Not frmContratosAlternativas.oRSAlt.EOF
            frmContratosAlternativas.oRSAlt.Delete
            frmContratosAlternativas.oRSAlt.MoveNext
        Loop

    End If

    Set frmContratosAlternativas.oRSComp = oRsComposicion.Clone

    'End If
    If frmContratosAlternativas.oRSAlt.RecordCount > 0 Then frmContratosAlternativas.oRSAlt.MoveFirst

    If oRsAlternativas.RecordCount > 0 Then

        Do While Not oRsAlternativas.EOF
            frmContratosAlternativas.oRSAlt.AddNew
            frmContratosAlternativas.oRSAlt!CODALTERNATIVA = oRsAlternativas!CODALTERNATIVA
            frmContratosAlternativas.oRSAlt!ALTERNATIVA = oRsAlternativas!ALTERNATIVA
            frmContratosAlternativas.oRSAlt!IDDSCTO = oRsAlternativas!IDDSCTO
            frmContratosAlternativas.oRSAlt!DESCUENTO = oRsAlternativas!DESCUENTO
            frmContratosAlternativas.oRSAlt!Cantidad = oRsAlternativas!Cantidad
            frmContratosAlternativas.oRSAlt!PRECIO = oRsAlternativas!PRECIO
            frmContratosAlternativas.oRSAlt!NETO = oRsAlternativas.Fields!NETO
            frmContratosAlternativas.oRSAlt!BRUTO = oRsAlternativas!BRUTO
            frmContratosAlternativas.oRSAlt.Update
            oRsAlternativas.MoveNext
        Loop

    End If

    oRsAlternativas.Filter = ""

    If Not oRsAlternativas.BOF And oRsAlternativas.EOF Then
        oRsAlternativas.MoveFirst
    End If

    frmContratosAlternativas.Show vbModal

    If frmContratosAlternativas.vAcepta Then
        'Set oRsAlternativas = frmContratosAlternativas.oRSAlt.Clone

        Do While Not oRsAlternativas.EOF
            oRsAlternativas.Delete
            oRsAlternativas.MoveNext
        Loop
        
        If frmContratosAlternativas.oRSAlt.RecordCount <> 0 Then
            frmContratosAlternativas.oRSAlt.Filter = ""
            frmContratosAlternativas.oRSAlt.MoveFirst
        End If

        Do While Not frmContratosAlternativas.oRSAlt.EOF
            oRsAlternativas.AddNew
            oRsAlternativas!CODALTERNATIVA = frmContratosAlternativas.oRSAlt!CODALTERNATIVA
            oRsAlternativas!ALTERNATIVA = frmContratosAlternativas.oRSAlt!ALTERNATIVA
            oRsAlternativas!IDDSCTO = frmContratosAlternativas.oRSAlt!IDDSCTO
            oRsAlternativas!DESCUENTO = frmContratosAlternativas.oRSAlt!DESCUENTO
            oRsAlternativas!Cantidad = frmContratosAlternativas.oRSAlt!Cantidad
            oRsAlternativas!PRECIO = frmContratosAlternativas.oRSAlt!PRECIO
            oRsAlternativas!NETO = frmContratosAlternativas.oRSAlt!NETO
            oRsAlternativas!BRUTO = frmContratosAlternativas.oRSAlt!BRUTO
            oRsAlternativas.Update
            frmContratosAlternativas.oRSAlt.MoveNext
        Loop

        oRsAlternativas.Filter = ""

        If oRsAlternativas.RecordCount <> 0 Then oRsAlternativas.MoveFirst
        Set oRsComposicion = frmContratosAlternativas.oRSComp.Clone
       
    End If
             
    Dim ors As ADODB.Recordset

    If oRsAlternativas.RecordCount > 0 Then
        'oRsAlternativas.Filter = "DETALLE=0"
       
        Set ors = New ADODB.Recordset
        ors.Fields.Append "CODALTERNATIVA", adBigInt
        ors.Fields.Append "ALTERNATIVA", adVarChar, 90
        ors.Fields.Append "IDDSCTO", adInteger, , adFldIsNullable
        ors.Fields.Append "DESCUENTO", adDouble
        ors.Fields.Append "CANTIDAD", adDouble
        ors.Fields.Append "PRECIO", adDouble
        ors.Fields.Append "NETO", adDouble
        ors.Fields.Append "BRUTO", adDouble
        ors.CursorLocation = adUseClient
        ors.LockType = adLockBatchOptimistic
        ors.CursorType = adOpenDynamic
        ors.Open

        If oRsAlternativas.RecordCount > 0 Then oRsAlternativas.MoveFirst

        Do While Not oRsAlternativas.EOF
            ors.AddNew
            ors.Fields!CODALTERNATIVA = oRsAlternativas!CODALTERNATIVA
            ors.Fields!ALTERNATIVA = oRsAlternativas!ALTERNATIVA
            ors.Fields!IDDSCTO = oRsAlternativas!IDDSCTO
            ors.Fields!DESCUENTO = oRsAlternativas!DESCUENTO
            ors.Fields!Cantidad = oRsAlternativas!Cantidad
            ors.Fields!PRECIO = oRsAlternativas!PRECIO
            ors.Fields!NETO = oRsAlternativas!NETO
            ors.Fields!BRUTO = oRsAlternativas!BRUTO
            ors.Update
            oRsAlternativas.MoveNext
        Loop
        
    End If

    oRsAlternativas.Filter = ""

    If Not oRsAlternativas.BOF And Not oRsAlternativas.EOF Then oRsAlternativas.MoveFirst

    Set Me.DatAlternativas.RowSource = ors
    'Me.DatAlternativas.BoundText = "alternativa"
    Me.DatAlternativas.ListField = "alternativa"
    Me.DatAlternativas.BoundColumn = "codalternativa"
    CalcularTotales
    
    Set Me.DatComposicion.RowSource = Nothing
    Me.DatAlternativas.BoundText = ""
    
End Sub

Private Sub txtNINIOS_KeyPress(KeyAscii As Integer)

    If SoloNumeros(KeyAscii) Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Me.ComExclusivo.SetFocus
End Sub

Private Sub CalcularTotales()

    If Not oRsAlternativas.BOF And Not oRsComposicion.EOF Then
        oRsAlternativas.Filter = ""

        If oRsAlternativas.RecordCount <> 0 Then oRsAlternativas.MoveFirst
    End If

    Dim vTOTAL As Double

    vTOTAL = 0

    Do While Not oRsAlternativas.EOF
        vTOTAL = vTOTAL + oRsAlternativas!BRUTO
        oRsAlternativas.MoveNext
    Loop
    
    'If Not oRsExternos.BOF And Not oRsExternos.EOF Then
    If oRsAlternativas.RecordCount <> 0 Then
        oRsAlternativas.Filter = ""
        oRsAlternativas.MoveFirst
    End If
    
    If oRsExternos.RecordCount <> 0 Then
        oRsExternos.Filter = ""
        oRsExternos.MoveFirst
    End If

    Do While Not oRsExternos.EOF
        vTOTAL = vTOTAL + oRsExternos!Importe
        oRsExternos.MoveNext
    Loop
    
    If oRsExternos.RecordCount <> 0 Then
        oRsExternos.Filter = ""
        oRsExternos.MoveFirst
    End If

    '    oRsAlternativas.Filter = "DETALLE=0"
    '
    '    If oRsAlternativas.RecordCount <> 0 Then oRsAlternativas.MoveFirst
    '
    '    Do While Not oRsAlternativas.EOF
    '        vTotal = vTotal + oRsAlternativas!Importe
    '        oRsAlternativas.MoveNext
    '    Loop
    
    Me.lblTotal.Caption = Format(vTOTAL, "##0.#0")
    Me.lblSaldo.Caption = Format(val(Me.lblTotal.Caption) - IIf(IsNumeric(Me.txtAcuenta.Text), val(Me.txtAcuenta.Text), 0), "##0.#0")
End Sub

Private Sub LlamarExternos()

    If frmContratosExternos.oRsExt Is Nothing Then
        Set frmContratosExternos.oRsExt = New ADODB.Recordset
        frmContratosExternos.oRsExt.Fields.Append "CODEXTERNO", adBigInt
        frmContratosExternos.oRsExt.Fields.Append "DESCRIPCION", adVarChar, 120
        frmContratosExternos.oRsExt.Fields.Append "CANTIDAD", adInteger
        frmContratosExternos.oRsExt.Fields.Append "PRECIO", adDouble
        frmContratosExternos.oRsExt.Fields.Append "IMPORTE", adDouble
        frmContratosExternos.oRsExt.CursorLocation = adUseClient
        frmContratosExternos.oRsExt.LockType = adLockBatchOptimistic
        frmContratosExternos.oRsExt.CursorType = adOpenDynamic
        frmContratosExternos.oRsExt.Open
    Else

        If frmContratosExternos.oRsExt.RecordCount > 0 Then
            frmContratosExternos.oRsExt.Filter = ""
            frmContratosExternos.oRsExt.MoveFirst
        End If

        Do While Not frmContratosExternos.oRsExt.EOF
            frmContratosExternos.oRsExt.Delete
            frmContratosExternos.oRsExt.MoveNext
        Loop

    End If
 
    Do While Not oRsExternos.EOF

        With frmContratosExternos.oRsExt
            .AddNew
            .Fields!CODEXTERNO = oRsExternos!CODEXTERNO ' oRsAlternativas!CODALTERNATIVA
            .Fields!DESCRIPCION = oRsExternos!DESCRIPCION
            .Fields!Cantidad = oRsExternos!Cantidad
            .Fields!PRECIO = oRsExternos!PRECIO
            .Fields!Importe = oRsExternos!Importe
            .Update
        End With

        oRsExternos.MoveNext
    Loop

    frmContratosExternos.Show vbModal
    CalcularTotales
End Sub

Private Sub RestaAcuenta()
    Me.lblSaldo.Caption = FormatNumber(val(Me.lblTotal.Caption) - IIf(IsNumeric(Me.txtAcuenta.Text), val(Me.txtAcuenta.Text), 0), 2)
End Sub

