VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmContratosAmortizaciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registrar Amortizaciones"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6735
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   480
      Left            =   5640
      Picture         =   "frmContratosAmortizaciones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4850
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   50
      TabIndex        =   7
      Top             =   15
      Width           =   6615
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4560
         TabIndex        =   13
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Cuenta:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2160
         TabIndex        =   11
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lblACuenta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblTotalContrato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4215
      Left            =   50
      TabIndex        =   14
      Top             =   585
      Width           =   6615
      Begin VB.ComboBox ComMoneda 
         Height          =   315
         ItemData        =   "frmContratosAmortizaciones.frx":038A
         Left            =   1530
         List            =   "frmContratosAmortizaciones.frx":0394
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtMonto 
         Height          =   285
         Left            =   3810
         MaxLength       =   8
         TabIndex        =   1
         Top             =   255
         Width           =   1455
      End
      Begin VB.CommandButton cmdAgregar 
         Height          =   360
         Left            =   2040
         Picture         =   "frmContratosAmortizaciones.frx":03C2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   990
      End
      Begin VB.CommandButton cmdQuitar 
         Height          =   360
         Left            =   3360
         Picture         =   "frmContratosAmortizaciones.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Width           =   990
      End
      Begin MSComctlLib.ListView lvDatos 
         Height          =   2295
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4048
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
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   3810
         TabIndex        =   3
         Top             =   615
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   159776769
         CurrentDate     =   40994
      End
      Begin MSDataListLib.DataCombo DatTipoPago 
         Height          =   315
         Left            =   1530
         TabIndex        =   2
         Top             =   615
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblTotal 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5040
         TabIndex        =   20
         Top             =   3750
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   4440
         TabIndex        =   19
         Top             =   3795
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda"
         Height          =   195
         Left            =   570
         TabIndex        =   18
         Top             =   300
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
         Height          =   195
         Left            =   5370
         TabIndex        =   17
         Top             =   300
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   195
         Left            =   5370
         TabIndex        =   16
         Top             =   675
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Pago"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   675
         Width           =   1110
      End
   End
End
Attribute VB_Name = "frmContratosAmortizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vPunto         As Boolean

Private Sub cmdAgregar_Click()

    If Me.ComMoneda.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda.", vbInformation, Pub_Titulo
        Me.ComMoneda.SetFocus

        Exit Sub

    End If

    If Len(Trim(Me.txtMonto.Text)) = 0 Then
        MsgBox "Debe ingresar el Monto.", vbInformation, Pub_Titulo
        Me.txtMonto.SetFocus

        Exit Sub

    End If

    If Me.DatTipoPago.BoundText = "" Then
        MsgBox "Debe seleccionar el Tipo de Pago.", vbInformation, Pub_Titulo
        Me.DatTipoPago.SetFocus

        Exit Sub

    End If

    If CDbl(Me.txtMonto.Text) > CDbl(Me.lblSaldo.Caption) Then
        MsgBox "El Monto proporcionado no debe exceder al saldo", vbCritical, Pub_Titulo
        Me.txtMonto.SetFocus

        Exit Sub

    End If

    Dim vTOTAL As Double

    vTOTAL = 0
    
    On Error GoTo Amortiza

    Pub_ConnAdo.BeginTrans

    If Not frmContratos.VNuevo Then
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPCONTRATOAMORTIZACION_REGISTRAR"
        oCmdEjec.CommandType = adCmdStoredProc

        Dim vIDAMORTIZACION As Integer

        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adDouble, adParamInput, , frmContratos.vIDContrato)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDFORMAPAGO", adInteger, adParamInput, , Me.DatTipoPago.BoundText)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MONTO", adDouble, adParamInput, , Me.txtMonto.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MONEDA", adChar, adParamInput, 1, IIf(Me.ComMoneda.ListIndex = 0, "S", "D"))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHAPAGO", adDBTimeStamp, adParamInput, , Me.dtpFecha.Value)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDAMORTIZACION", adInteger, adParamOutput, , vIDAMORTIZACION)

        oCmdEjec.Execute

        vIDAMORTIZACION = oCmdEjec.Parameters(6).Value

        Dim ITEMx As Object

        For Each ITEMx In Me.lvDatos.ListItems
                
            frmContratos.oRSAMORTIZACIONES.AddNew
            frmContratos.oRSAMORTIZACIONES.Fields!moneda = Left(ITEMx.Text, 1)
            frmContratos.oRSAMORTIZACIONES.Fields!MONTO = ITEMx.SubItems(1)
            frmContratos.oRSAMORTIZACIONES.Fields!IDFORMAPAGO = ITEMx.SubItems(2)
            frmContratos.oRSAMORTIZACIONES.Fields!FORMAPAGO = ITEMx.SubItems(3)
            frmContratos.oRSAMORTIZACIONES.Fields!FECHAPAGO = ITEMx.SubItems(4)
            frmContratos.oRSAMORTIZACIONES.Fields!IDAMORTIZACION = ITEMx.Tag
            frmContratos.oRSAMORTIZACIONES.Update
            vTOTAL = vTOTAL + ITEMx.SubItems(1)

        Next
    
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPCONTRATO_UPDATEMONTO"
        oCmdEjec.CommandType = adCmdStoredProc
    
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adDouble, adParamInput, , frmContratos.vIDContrato)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MONTO", adDouble, adParamInput, , vTOTAL + txtMonto.Text)
    
        oCmdEjec.Execute

        Pub_ConnAdo.CommitTrans
    Else

        Dim ITEMxs As Object

        'For Each ITEMxs In Me.lvDatos.ListItems
                
            frmContratos.oRSAMORTIZACIONES.AddNew
            frmContratos.oRSAMORTIZACIONES.Fields!moneda = IIf(Me.ComMoneda.ListIndex = 0, "S", "D")
            frmContratos.oRSAMORTIZACIONES.Fields!MONTO = Me.txtMonto.Text
            frmContratos.oRSAMORTIZACIONES.Fields!IDFORMAPAGO = Me.DatTipoPago.BoundText
            frmContratos.oRSAMORTIZACIONES.Fields!FORMAPAGO = Me.DatTipoPago.Text
            frmContratos.oRSAMORTIZACIONES.Fields!FECHAPAGO = Me.dtpFecha.Value
            frmContratos.oRSAMORTIZACIONES.Fields!IDAMORTIZACION = -1
            frmContratos.oRSAMORTIZACIONES.Update
            vTOTAL = vTOTAL + Me.txtMonto.Text

       ' Next

    End If

    Dim ITEMa As Object

    Set ITEMa = Me.lvDatos.ListItems.Add(, , IIf(Me.ComMoneda.ListIndex = 0, "SOLES", "DOLARES"))
    ITEMa.SubItems(1) = Me.txtMonto.Text
    ITEMa.SubItems(2) = Me.DatTipoPago.BoundText
    ITEMa.SubItems(3) = Me.DatTipoPago.Text
    ITEMa.SubItems(4) = Me.dtpFecha.Value
    ITEMa.Tag = vIDAMORTIZACION
    
    Me.lblTotal.Caption = val(Me.lblTotal.Caption) + val(Me.txtMonto.Text)

    Me.lblSaldo.Caption = val(Me.lblTotalContrato.Caption) - val(Me.lblTotal.Caption)
    Me.lblACuenta.Caption = Me.lblTotal.Caption

    If Not frmContratos.oRSAMORTIZACIONES.BOF Or Not frmContratos.oRSAMORTIZACIONES.EOF Then

        frmContratos.oRSAMORTIZACIONES.MoveFirst

    End If

'    Do While Not frmContratos.oRSAMORTIZACIONES.EOF
'        frmContratos.oRSAMORTIZACIONES.Delete
'        frmContratos.oRSAMORTIZACIONES.MoveNext
'    Loop

    frmContratos.txtAcuenta.Text = Me.lblACuenta.Caption
    frmContratos.lblSaldo.Caption = Me.lblSaldo.Caption
    
    Me.ComMoneda.ListIndex = -1
    Me.DatTipoPago.BoundText = ""
    Me.txtMonto.Text = ""
    Me.dtpFecha.Value = Date
    Me.ComMoneda.SetFocus
    
    Exit Sub
    
Amortiza:
    Pub_ConnAdo.RollbackTrans
    MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Function VerificaPass(vUSUARIO As String, vClave As String, ByRef vMSN As String) As Boolean
Dim orsPass As ADODB.Recordset
Dim vtpass As String, vPasa As Boolean
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SPDEVUELVECLAVEAMORTIZACIONES"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, vUSUARIO)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CLAVE", adVarChar, adParamInput, 10, vClave)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MSN", adVarChar, adParamOutput, 200, 1)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)
oCmdEjec.Execute

'If Not orsPass.EOF Then vtpass = Trim(orsPass!Clave)
vtpass = oCmdEjec.Parameters("@MSN").Value
vPasa = oCmdEjec.Parameters("@PASA").Value
vMSN = vtpass

    VerificaPass = vPasa
End Function

Private Sub EliminaAmortizacion()
On Error GoTo Eliminar
    
    Pub_ConnAdo.BeginTrans

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPCONTRATOAMORTIZACION_ANULAR"
    oCmdEjec.CommandType = adCmdStoredProc
  
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adDouble, adParamInput, , frmContratos.vIDContrato)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDAMORTIZACION", adInteger, adParamInput, , Me.lvDatos.SelectedItem.Tag)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MONTO", adDouble, adParamInput, , Me.lblTotal.Caption)
  
    oCmdEjec.Execute
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPCONTRATO_UPDATEMONTO"
    oCmdEjec.CommandType = adCmdStoredProc
    
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adDouble, adParamInput, , frmContratos.vIDContrato)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MONTO", adDouble, adParamInput, , Me.lblTotal.Caption - Me.lvDatos.SelectedItem.SubItems(1))
    
    oCmdEjec.Execute
  
  
    Me.lblACuenta.Caption = Me.lblTotal.Caption - Me.lvDatos.SelectedItem.SubItems(1)
    Me.lblTotal.Caption = Me.lblACuenta.Caption
    Me.lblSaldo.Caption = val(Me.lblTotalContrato.Caption) - val(Me.lblACuenta.Caption)
   
    frmContratos.txtAcuenta.Text = Me.lblACuenta.Caption
    frmContratos.lblSaldo.Caption = val(frmContratos.lblTotal.Caption) - val(Me.lblACuenta.Caption)
  
    Me.lvDatos.ListItems.Remove Me.lvDatos.SelectedItem.Index
  
  Pub_ConnAdo.CommitTrans
    Exit Sub

Eliminar:
Pub_ConnAdo.RollbackTrans
    MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub

Private Sub cmdQuitar_Click()

    If Me.lvDatos.SelectedItem Is Nothing Then Exit Sub
    
    
     If MsgBox("No se puede Anular una Amortización" & vbCrLf & "¿Desea ingresar la clave?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
            frmClaveCaja.Show vbModal
            If frmClaveCaja.vAceptar Then
                Dim vS As String
                If VerificaPass(frmClaveCaja.vUSUARIO, frmClaveCaja.vClave, vS) Then
                    EliminaAmortizacion
                Else
                    MsgBox vS, vbCritical, Pub_Titulo
                End If
            End If
        End If
        
End Sub

Private Sub ComMoneda_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Me.txtMonto.SetFocus
End Sub

Private Sub DatTipoPago_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.dtpFecha.SetFocus
End Sub

Private Sub dtpFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.cmdAgregar.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPCARGARFORMASPAGO"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim ors As ADODB.Recordset

    Set ors = oCmdEjec.Execute(, 2401)

    Set Me.DatTipoPago.RowSource = ors
    Me.DatTipoPago.ListField = ors.Fields(1).Name
    Me.DatTipoPago.BoundColumn = ors.Fields(0).Name
        
    ConfigurarLV
    CargarAmortizaciones
End Sub

Private Sub ConfigurarLV()

    With Me.lvDatos
        .ColumnHeaders.Add , , "Moneda"
        .ColumnHeaders.Add , , "Monto"
        .ColumnHeaders.Add , , "IDFORMAPAGO", 0
        .ColumnHeaders.Add , , "Tipo Pago"
        .ColumnHeaders.Add , , "Fecha Pago"
        .Gridlines = True
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .View = lvwReport
    End With

End Sub

Private Sub CargarAmortizaciones()

If frmContratos.VNuevo Then 'NUEVO
Dim ITEMx As Object

If Not frmContratos.oRSAMORTIZACIONES.EOF Or Not frmContratos.oRSAMORTIZACIONES.BOF Then
    frmContratos.oRSAMORTIZACIONES.MoveFirst
    End If

    Do While Not frmContratos.oRSAMORTIZACIONES.EOF
        Set ITEMx = Me.lvDatos.ListItems.Add(, , IIf(frmContratos.oRSAMORTIZACIONES!moneda = "S", "SOLES", "DOLARES"))
        ITEMx.Tag = frmContratos.oRSAMORTIZACIONES!IDAMORTIZACION
        ITEMx.SubItems(1) = frmContratos.oRSAMORTIZACIONES!MONTO
        ITEMx.SubItems(2) = frmContratos.oRSAMORTIZACIONES!IDFORMAPAGO
        ITEMx.SubItems(3) = Trim(frmContratos.oRSAMORTIZACIONES!FORMAPAGO)
        ITEMx.SubItems(4) = frmContratos.oRSAMORTIZACIONES!FECHAPAGO
        frmContratos.oRSAMORTIZACIONES.MoveNext
    Loop
Else 'MODIFICA
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SPCONTRATOAMORTIZACION_LISTAR"
oCmdEjec.CommandType = adCmdStoredProc

oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adDouble, adParamInput, , frmContratos.vIDContrato)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

Dim orsA As ADODB.Recordset

Set orsA = oCmdEjec.Execute

   Do While Not orsA.EOF
        Set ITEMx = Me.lvDatos.ListItems.Add(, , IIf(orsA!moneda = "S", "SOLES", "DOLARES"))
        ITEMx.Tag = orsA!IDAMORTIZACION
        ITEMx.SubItems(1) = orsA!MONTO
        ITEMx.SubItems(2) = orsA!IDFORMAPAGO
        ITEMx.SubItems(3) = Trim(orsA!FORMAPAGO)
        ITEMx.SubItems(4) = orsA!FECHAPAGO
        orsA.MoveNext
    Loop


End If


    

    Me.lblTotalContrato.Caption = frmContratos.lblTotal.Caption
    Me.lblACuenta.Caption = frmContratos.txtAcuenta.Text
    Me.lblSaldo.Caption = frmContratos.lblSaldo.Caption

    Me.lblTotal.Caption = frmContratos.txtAcuenta.Text
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1 'ACEPTAR

'            Dim vTOTAL As Double
'
'            vTOTAL = 0
'
'            If Not frmContratos.oRSAMORTIZACIONES.BOF Or Not frmContratos.oRSAMORTIZACIONES.EOF Then
'
'                frmContratos.oRSAMORTIZACIONES.MoveFirst
'
'            End If
'
'            Do While Not frmContratos.oRSAMORTIZACIONES.EOF
'                frmContratos.oRSAMORTIZACIONES.Delete
'                frmContratos.oRSAMORTIZACIONES.MoveNext
'            Loop
'
'            Dim ITEMx As Object
'
'            For Each ITEMx In Me.lvDatos.ListItems
'
'                frmContratos.oRSAMORTIZACIONES.AddNew
'                frmContratos.oRSAMORTIZACIONES.Fields!moneda = Left(ITEMx.Text, 1)
'                frmContratos.oRSAMORTIZACIONES.Fields!MONTO = ITEMx.SubItems(1)
'                frmContratos.oRSAMORTIZACIONES.Fields!IDFORMAPAGO = ITEMx.SubItems(2)
'                frmContratos.oRSAMORTIZACIONES.Fields!FORMAPAGO = ITEMx.SubItems(3)
'                frmContratos.oRSAMORTIZACIONES.Fields!FECHAPAGO = ITEMx.SubItems(4)
'                frmContratos.oRSAMORTIZACIONES.Fields!IDAMORTIZACION = ITEMx.Tag
'                frmContratos.oRSAMORTIZACIONES.Update
'                vTOTAL = vTOTAL + ITEMx.SubItems(1)
'
'            Next
'
'            frmContratos.txtAcuenta.Text = vTOTAL
'            frmContratos.lblSaldo.Caption = val(frmContratos.lblTotal.Caption) - vTOTAL
'            Unload Me

        Case 2 'CANCELAR
            Unload Me
    End Select

End Sub

Private Sub txtMonto_Change()
 If InStr(txtMonto.Text, ".") = 0 Then
        vPunto = False
    Else
        vPunto = True
    End If
End Sub

Private Sub txtMonto_GotFocus()
  If InStr(txtMonto.Text, ".") = 0 Then
        vPunto = False
    Else
        vPunto = True
    End If
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.DatTipoPago.SetFocus
 If NumerosyPunto(KeyAscii) Then KeyAscii = 0
    If Chr(KeyAscii) = "." Then
        If vPunto = True Or Len(Trim(txtMonto.Text)) = 0 Then
         
            KeyAscii = 0
         
        End If
    End If
End Sub
