VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmFacComanda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturar Comanda"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10110
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
   ScaleHeight     =   8565
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   1560
      TabIndex        =   18
      Top             =   8040
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1695
      Left            =   1800
      TabIndex        =   15
      Top             =   2040
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame fraTipDocto 
      Caption         =   "Empresa que Factura:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin VB.CheckBox chkprom 
         Caption         =   "Imprime Promocion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   6480
         TabIndex        =   41
         Top             =   240
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo DatEmpresas 
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   280
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.CheckBox chkGratuito 
         Caption         =   "Transferencia Gratuita"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   6480
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CheckBox chkEdit 
         Caption         =   "Editar Número"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   27
         Top             =   960
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo DatTiposDoctos 
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "TiposDoctos"
      End
      Begin VB.CheckBox chkConsumo 
         Caption         =   "Facturar x Consumo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6480
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtNro 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblCiaA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento a Emitir:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   1980
      End
      Begin VB.Label lblSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame fraCliente 
      Height          =   1815
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   9735
      Begin VB.TextBox txtDni 
         Height          =   285
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   34
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton cmdSunat 
         Caption         =   "Sunat"
         Height          =   480
         Left            =   9000
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtDireccion 
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   840
         Width           =   7935
      End
      Begin VB.TextBox txtRuc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7320
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtRS 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "DNI:"
         Height          =   195
         Left            =   1080
         TabIndex        =   21
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "DIRECCIÓN:"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nº RUC"
         Height          =   195
         Left            =   6600
         TabIndex        =   10
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "RAZÓN SOCIAL"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2520
      Left            =   240
      TabIndex        =   17
      Top             =   3240
      Width           =   9735
      Begin MSComctlLib.ListView lvDetalle 
         Height          =   2175
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   3836
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame fraImporte 
      Height          =   2175
      Left            =   240
      TabIndex        =   3
      Top             =   5760
      Width           =   9735
      Begin VB.CommandButton cmdFormasPago 
         Caption         =   "Forma de Pago"
         Height          =   720
         Left            =   1920
         TabIndex        =   24
         Top             =   480
         Width           =   3135
      End
      Begin VB.CommandButton cmdCobrar 
         Caption         =   "&Cobrar"
         Height          =   360
         Left            =   2880
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdDscto 
         Caption         =   "Descuento"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblporcigv 
         Caption         =   "Label3"
         Height          =   255
         Left            =   6240
         TabIndex        =   40
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblpICBPER 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label13"
         Height          =   195
         Left            =   5160
         TabIndex        =   39
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ICBPER:"
         Height          =   195
         Left            =   6720
         TabIndex        =   38
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblICBPER 
         BackColor       =   &H8000000C&
         Height          =   315
         Left            =   7680
         TabIndex        =   37
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IGV:"
         Height          =   195
         Left            =   6240
         TabIndex        =   32
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALOR VENTA:"
         Height          =   195
         Left            =   6240
         TabIndex        =   31
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label lblvvta 
         BackColor       =   &H8000000C&
         Height          =   315
         Left            =   7680
         TabIndex        =   30
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblIGV 
         BackColor       =   &H8000000C&
         Height          =   315
         Left            =   7680
         TabIndex        =   29
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblDscto 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   360
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblImporte 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7680
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblvuelto 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "VUELTO:"
         Height          =   195
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL A PAGAR:"
         Height          =   195
         Left            =   6120
         TabIndex        =   4
         Top             =   1800
         Width           =   1470
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   8040
      Width           =   1335
   End
End
Attribute VB_Name = "frmFacComanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vSerCom As String
Private buscars As Boolean
Public vNroCom, vCodMoz As Integer
Public vMesa As String
Public vTOTAL As Currency
Dim loc_key  As Integer
Private oRsPago As ADODB.Recordset
Private oRsTipPag As ADODB.Recordset
Public vAcepta As Boolean
Public vOper As Integer 'variable para capturar el allnumoper de allog para imprimir del facart
Private vBuscar As Boolean 'variable para la busqueda de clientes
Public xMostrador As Boolean
Public gDESCUENTO As Double 'VARIABLE PARA ALMACENAR EL DESCUENTO PARA LAS COMANDAS
Public gPAGO As Double 'VARIABLE PARA ALMACENAR EL PAGO PARA LAS COMANDAS
Private ORStd As ADODB.Recordset 'VARIABLE PARA SABER SI EL TIPO DE DOCUMENTO ES EDITABLE


Private Sub ConfigurarLVDetalle()
With Me.lvDetalle
    .ColumnHeaders.Add , , "NumFac", 0
    .ColumnHeaders.Add , , "Codigo", 800
    .ColumnHeaders.Add , , "Plato", 4000
    .ColumnHeaders.Add , , "Cta", 500
    .ColumnHeaders.Add , , "Precio", 800
    .ColumnHeaders.Add , , "Can. Tot.", 1000
    .ColumnHeaders.Add , , "Faltante", 1000
    .ColumnHeaders.Add , , "Importe", 1200
    .ColumnHeaders.Add , , "Sec", 0
    .ColumnHeaders.Add , , "apro", 0
    .ColumnHeaders.Add , , "aten", 0
    .ColumnHeaders.Add , , "unidad", 0
    .ColumnHeaders.Add , , "pednumsec", 0
    .ColumnHeaders.Add , , "cambio", 800
    .ColumnHeaders.Add , , "ICBPER", 300
    .ColumnHeaders.Add , , "comboICBPER", 300
    .View = lvwReport
    .FullRowSelect = True
    .LabelEdit = lvwManual
    .Gridlines = True
    .HideSelection = False
End With
End Sub

Private Function VerificaPass(vUSUARIO As String, vClave As String, ByRef vMSN As String) As Boolean
Dim orsPass As ADODB.Recordset
Dim vtpass As String, vPasa As Boolean
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SpDevuelveClaveCaja"
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



Private Sub ConfiguraLV()
With Me.ListView1
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Codigo", 1000
    .ColumnHeaders.Add , , "Cliente", 5000
    .ColumnHeaders.Add , , "Ruc", 0
    .ColumnHeaders.Add , , "Direcion", 0
    .MultiSelect = False
End With
End Sub



Private Sub cboMoneda_KeyPress(index As Integer, KeyAscii As Integer)
If LK_TIPO_CAMBIO = 0 Then
    MsgBox "ingresar tipo de cambio"
End If
End Sub


Private Sub chkEdit_Click()
Me.txtNro.Enabled = Me.chkEdit.Value
If Me.chkEdit.Value Then
Me.txtNro.SetFocus
Me.txtNro.SelStart = 0
Me.txtNro.SelLength = Len(Me.txtNro.Text)
End If
End Sub

Private Sub chkGratuito_Click()

    If Me.chkGratuito.Value Then

        frmClaveCaja.Show vbModal
    
        If frmClaveCaja.vAceptar Then
    
            Dim vS As String
    
            If VerificaPass(frmClaveCaja.vUSUARIO, frmClaveCaja.vClave, vS) Then
               Me.lblImporte.Caption = "0.00"
                Me.cmdformaspago.Enabled = False
            Else
                MsgBox "Clave incorrecta", vbCritical, NombreProyecto
            End If
        End If
Else
Me.cmdformaspago.Enabled = True
CalcularImporte
    End If

End Sub

Private Sub cmdAceptar_Click()
'ImprimirDocumentoVenta "01", "Factura", True, "1", 27427495, "126.5", "caleta", "123", "tirado" ' Me.DatTiposDoctos.BoundText, Me.DatTiposDoctos.Text, Me.chkConsumo.Value, Me.lblserie.Caption, Me.txtNro.Text, Me.lblImporte.Caption, Me.txtDireccion.Text, Me.txtRuc.Text, Me.txtcli.Text
'Exit Sub

If Me.DatTiposDoctos.BoundText = "" Then ' .ListIndex = -1 Then
    MsgBox "Debe elegir el Tipo de documento.", vbCritical, Pub_Titulo
    Me.DatTiposDoctos.SetFocus

    Exit Sub

End If

If Me.cmdformaspago.Enabled And oRSfp.RecordCount = 0 Then
    MsgBox "Debe ingresar pagos", vbCritical, Pub_Titulo

    Exit Sub

End If
    
'    If oRSfp.RecordCount <> 0 Then
'        oRSfp.MoveFirst
'        oRSfp.Filter = "IDFORMAPAGO=4"
'        If oRSfp.RecordCount <> 0 And Len(Trim(Me.txtRuc.Tag)) = 0 Then
'            MsgBox "Debe elegir el cliente.", vbCritical, Pub_Titulo
'            oRSfp.Filter = ""
'        oRSfp.MoveFirst
'        Exit Sub
'        End If
'    End If

Dim xPAGOS         As Double

Dim xCONTINUA      As Boolean

Dim xARCENCONTRADO As Boolean

xCONTINUA = False
xARCENCONTRADO = False
xPAGOS = 0

oRSfp.MoveFirst

Do While Not oRSfp.EOF
    xPAGOS = xPAGOS + oRSfp!monto
    oRSfp.MoveNext
Loop
    
If Me.cmdformaspago.Enabled And xPAGOS < val(Me.lblImporte.Caption) Then
    MsgBox "Falta importe por pagar", vbCritical, Pub_Titulo

    Exit Sub

End If

On Error GoTo Graba

'If Me.cboTipoDocto.ListIndex = 0 Then 'F
If Me.DatTiposDoctos.BoundText = "01" Then 'F
    If Me.lvDetalle.ListItems.count > par_llave!par_fac_lines And Me.chkConsumo.Value = 0 Then
        MsgBox "Numero Máximo de Lineas alcanzado"

        Exit Sub

    End If

'ElseIf Me.cboTipoDocto.ListIndex = 1 Then 'B
ElseIf Me.DatTiposDoctos.BoundText = "03" Then 'B

    If Me.lvDetalle.ListItems.count > par_llave!par_BOL_lines And Me.chkConsumo.Value = 0 Then
        MsgBox "Numero Máximo de Lineas alcanzado"

        Exit Sub

    End If
End If

Dim f As Integer

Dim sOri, sMod As Integer

If Me.lvDetalle.ListItems.count = 0 Then
    MsgBox "No hay ningun plato para procesar"

    Exit Sub

End If

If Me.DatTiposDoctos.BoundText = "01" Then
    If Len(Trim(Me.txtRuc.Text)) = 0 Then
        MsgBox "Debe ingresar el Ruc para poder generar la Factura", vbInformation, "Error"

        Exit Sub

    End If
End If

'valida la uit
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SP_VALIDA_UIT"

Dim ORSuit As ADODB.Recordset

Set ORSuit = oCmdEjec.Execute(, Me.lblImporte.Caption)

If Not ORSuit.EOF Then
    If ORSuit!Dato = 1 Then
        If Len(Trim(Me.txtRS.Text)) = 0 Or (Len(Trim(Me.txtDni.Text)) = 0 And Len(Trim(Me.txtRuc.Text)) = 0) Then
            MsgBox "El Importe sobrepasa media UIT, debe ingresar el cliente.", vbCritical, Pub_Titulo

            Exit Sub

        End If
    End If
End If
    
If Len(Trim(Me.txtDni.Text)) <> 0 Then
    If Len(Trim(Me.txtDni.Text)) < 8 Then
        MsgBox "El DNI debe tener 8 dígitos.", vbInformation, Pub_Titulo

        Exit Sub

    End If
End If

'Armando Xml
Dim CP As Double

Dim pr, st As Currency

Dim vXml, un, cam As String

vXml = "<r>"

For f = 1 To Me.lvDetalle.ListItems.count

    CP = Me.lvDetalle.ListItems(f).SubItems(1)
'    Me.dgrdData.Row = f
'    'Codigo de Plato
'    Me.dgrdData.COL = 1
'    CP = CInt(Trim(Me.dgrdData.Text))
'Cantidad a facturar
    st = Me.lvDetalle.ListItems(f).SubItems(6)
'    Me.dgrdData.COL = 5
'    st = CInt(Trim(Me.dgrdData.Text))
'Precio
'Me.dgrdData.COL = 6
'pr = CDec(Trim(Me.dgrdData.Text))
    pr = Me.lvDetalle.ListItems(f).SubItems(4)
'Unidad de medida
'Me.dgrdData.COL = 10
'un = Trim(Me.dgrdData.Text)
    un = Trim(Me.lvDetalle.ListItems(f).SubItems(11))
'SECUENCIA
'Me.dgrdData.COL = 7
'    sc = Trim(Me.dgrdData.Text)
    sc = Me.lvDetalle.ListItems(f).SubItems(8)
    cam = Me.lvDetalle.ListItems(f).SubItems(13)
    
    vXml = vXml & "<d "
    vXml = vXml & "cp=""" & Trim(str(CP)) & """ "
    vXml = vXml & "st=""" & Trim(str(st)) & """ "
    vXml = vXml & "pr=""" & Trim(str(pr)) & """ "
    vXml = vXml & "un=""" & Trim(un) & """ "
    vXml = vXml & "sc=""" & Trim(sc) & """ "
    vXml = vXml & "cam=""" & Trim(cam) & """ "
    vXml = vXml & "/>"
Next

vXml = vXml & "</r>"
 
'obteniendo datos de tipo de pago tabla sub_transa
'Dim alltipdoc As String
'Dim allcp As String
'oRsTipPag.Filter = "sut_secuencia=" & Me.dcboPago.BoundText
'alltipdoc = oRsTipPag!sut_tipdoc
'allcp = oRsTipPag!sut_cp
    
'recorriendo las formas de pago
Dim xFP      As String

Dim xPAGACON As Double, xVUELTO As Double
    
If oRSfp.RecordCount <> 0 Then
    oRSfp.MoveFirst
    xPAGACON = oRSfp!pagacon
    xVUELTO = oRSfp!VUELTO
    xFP = "<r>"

    Do While Not oRSfp.EOF
        xFP = xFP & "<d "
        xFP = xFP & "idfp=""" & Trim(str(oRSfp!idformapago)) & """ "
        xFP = xFP & "fp=""" & Trim(oRSfp!formaPAGO) & """ "
        xFP = xFP & "mon=""" & "S" & """ "
        xFP = xFP & "monto=""" & Trim(str(oRSfp!monto)) & """ "
        xFP = xFP & "ref=""" & Trim(oRSfp!referencia) & """ "
        xFP = xFP & "dcre=""" & Trim(oRSfp!diascredito) & """ "
        xFP = xFP & "/>"
        oRSfp.MoveNext
    Loop

    xFP = xFP & "</r>"
End If

With oCmdEjec
        
'PARCHE - SE TENDRIA QUE HACER MEJOR EN EL SP PARA UNA MEJOR CONSISTENCIA DE DATOS
    If vAcepta Then
        MsgBox "Ya se facturo.", vbInformation, Pub_Titulo

        Exit Sub

    Else
'VALIDANDO SI EL ARCHIVO DEL REPORTE EXISTE
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_ARCHIVO_PRINT"
        oCmdEjec.CommandType = adCmdStoredProc
    
        Dim ORSd        As ADODB.Recordset

        Dim RutaReporte As String

        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGO", adChar, adParamInput, 2, Me.DatTiposDoctos.BoundText)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@COMSUMO", adBoolean, adParamInput, , Me.chkConsumo.Value)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, Me.DatEmpresas.BoundText)
    
        Set ORSd = oCmdEjec.Execute
        RutaReporte = PUB_RUTA_REPORTE & ORSd!ReportE
    
        If ORSd!ReportE = "" Then
            If MsgBox("El Archivo no existe." & vbCrLf & "¿Desea continuar sin imprimir?.", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
                xCONTINUA = True
            End If

        Else
            FileName = dir(RutaReporte)

            If FileName = "" Then
                If MsgBox("El Archivo no existe." & vbCrLf & "¿Desea continuar sin imprimir?.", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
                    xCONTINUA = True
                Else

                    Exit Sub

                End If

            Else
                xCONTINUA = True
                xARCENCONTRADO = True
'ImprimirDocumentoVenta Me.DatTiposDoctos.BoundText, Me.DatTiposDoctos.Text, Me.chkConsumo.Value, "009", 2, 100, "33", "43", "343"
            End If
        End If

        If xCONTINUA Then
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpFacturarComanda"
            .Parameters.Append .CreateParameter("@codcia", adChar, adParamInput, 2, Me.DatEmpresas.BoundText)
            .Parameters.Append .CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
            .Parameters.Append .CreateParameter("@usuario", adVarChar, adParamInput, 20, LK_CODUSU)
            .Parameters.Append .CreateParameter("@SerCom", adChar, adParamInput, 3, vSerCom)
            .Parameters.Append .CreateParameter("@nroCom", adInteger, adParamInput, , vNroCom)
            .Parameters.Append .CreateParameter("@SerDoc", adChar, adParamInput, 3, Me.lblSerie.Caption)
            .Parameters.Append .CreateParameter("@NroDoc", adDouble, adParamInput, , CDbl(Me.txtNro.Text))
            .Parameters.Append .CreateParameter("@Fbg", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1)) ' Me.cboTipoDocto.Text)
            .Parameters.Append .CreateParameter("@XmlDet", adVarChar, adParamInput, 4000, Trim(vXml))

            If Me.DatTiposDoctos.BoundText = "01" Then
                .Parameters.Append .CreateParameter("@codcli", adVarChar, adParamInput, 10, Me.txtRuc.Tag)
            Else
                .Parameters.Append .CreateParameter("@codcli", adVarChar, adParamInput, 10, IIf(Len(Trim(Me.txtRuc.Tag)) = 0, 1, Me.txtRuc.Tag))
            End If

            .Parameters.Append .CreateParameter("@codMozo", adInteger, adParamInput, , vCodMoz)
            .Parameters.Append .CreateParameter("@totalfac", adDouble, adParamInput, , Me.lblImporte.Caption)
            If oRSfp.RecordCount <> 0 Then oRSfp.MoveFirst
            '.Parameters.Append .CreateParameter("@sec", adInteger, adParamInput, , Me.dcboPago.BoundText)
            .Parameters.Append .CreateParameter("@sec", adInteger, adParamInput, , oRSfp!idformapago)
            .Parameters.Append .CreateParameter("@moneda", adChar, adParamInput, 1, "S")
            .Parameters.Append .CreateParameter("@diascre", adInteger, adParamInput, , 0)
            .Parameters.Append .CreateParameter("@farjabas", adTinyInt, adParamInput, , IIf(Me.chkConsumo.Value = 1, 1, 0))
'.Parameters.Append .CreateParameter("@dscto", adDouble, adParamInput, , IIf(Len(Trim(Me.lblDscto.Caption)) = 0, 0, Me.lblDscto.Caption))
            .Parameters.Append .CreateParameter("@dscto", adDouble, adParamInput, , gDESCUENTO)
            .Parameters.Append .CreateParameter("@CODIGODOCTO", adChar, adParamInput, 2, Me.DatTiposDoctos.BoundText)
            .Parameters.Append .CreateParameter("@Xmlpag", adVarChar, adParamInput, 4000, xFP)
            .Parameters.Append .CreateParameter("@PAGACON", adDouble, adParamInput, , xPAGACON)
            .Parameters.Append .CreateParameter("@VUELTO", adDouble, adParamInput, , xVUELTO)
                
            .Parameters.Append .CreateParameter("@VALORVTA", adDouble, adParamInput, , Me.lblvvta.Caption)
            .Parameters.Append .CreateParameter("@VIGV", adDouble, adParamInput, , Me.lblIgv.Caption)
            .Parameters.Append .CreateParameter("@GRATUITO", adBoolean, adParamInput, , Me.chkGratuito.Value)
            .Parameters.Append .CreateParameter("@CIAPEDIDO", adChar, adParamInput, 2, LK_CODCIA)
            .Parameters.Append .CreateParameter("@ALL_ICBPER", adDouble, adParamInput, , IIf(Len(Trim(Me.lblicbper.Caption)) = 0, 0, Me.lblicbper.Caption))
            .Parameters.Append .CreateParameter("@MaxNumOper", adInteger, adParamOutput, , 0)
            .Parameters.Append .CreateParameter("@AUTONUMFAC", adInteger, adParamOutput, , 0)
        
            .Execute

            If Not IsNull(oCmdEjec.Parameters("@MaxNumOper").Value) Then
                vOper = oCmdEjec.Parameters("@MaxNumOper").Value
            End If

            Me.txtNro.Text = oCmdEjec.Parameters("@AUTONUMFAC").Value
                
            If vAcepta = False Then

                vAcepta = True
'MsgBox "Datos Almacenados correctamente", vbInformation, Pub_Titulo

'Imprimir Left(Me.DatTiposDoctos.Text, 1), Me.chkConsumo.Value
                CreaCodigoQR "6", Me.DatTiposDoctos.BoundText, Me.lblSerie.Caption, Me.txtNro.Text, LK_FECHA_DIA, CStr(Me.lblIgv.Caption), Me.lblImporte.Caption, Me.txtRuc.Text, Me.txtDni.Text
                If xARCENCONTRADO Then
                    ImprimirDocumentoVenta Me.DatTiposDoctos.BoundText, Me.DatTiposDoctos.Text, Me.chkConsumo.Value, Me.lblSerie.Caption, Me.txtNro.Text, Me.lblImporte.Caption, Me.lblvvta.Caption, Me.lblIgv.Caption, Me.txtDireccion.Text, Me.txtRuc.Text, Me.txtRS.Text, Me.txtDni.Text, Me.DatEmpresas.BoundText, IIf(Len(Trim(Me.lblicbper.Caption)) = 0, 0, Me.lblicbper.Caption), Me.chkprom.Value
                End If
                    
'If Me.DatTiposDoctos.BoundText = "01" Or Me.DatTiposDoctos.BoundText = "03" Then
               ' If Me.DatTiposDoctos.BoundText = "01" Then
                    If LK_PASA_BOLETAS = "A" And (Me.DatTiposDoctos.BoundText = "01" Or Me.DatTiposDoctos.BoundText = "03") Then
                    CrearArchivoPlano Left(Me.DatTiposDoctos.Text, 1), Me.lblSerie.Caption, Me.txtNro.Text
                    ElseIf Me.DatTiposDoctos.BoundText = "01" Then
                    CrearArchivoPlano Left(Me.DatTiposDoctos.Text, 1), Me.lblSerie.Caption, Me.txtNro.Text
                    End If
               ' End If

                Unload Me
            Else
                MsgBox "Ya se facturo"
            End If

        End If
    End If

'FIN DEL PARCHE
End With
   
Exit Sub

Graba:
MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdCancelar_Click()
    Unload Me

  

End Sub

Private Sub cmdCobrar_Click()
If oRSfp.RecordCount = 0 Then
        MsgBox "Debe ingresar pagos", vbCritical, Pub_Titulo
        Exit Sub
    End If

  Dim xPAGOS As Double

    xPAGOS = 0

    'If Not oRSfp.EOF Then oRSfp.MoveFirst
    Do While Not oRSfp.EOF
        xPAGOS = xPAGOS + oRSfp!monto
        oRSfp.MoveNext
    Loop
    
    If xPAGOS <= 0 Then
        MsgBox "Falta importe por pagar", vbCritical, Pub_Titulo

        Exit Sub

    End If
    

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_DOCUMENTO_PAGAR"

    Dim xFP As String
    
    If oRSfp.RecordCount <> 0 Then
        oRSfp.MoveFirst
        xFP = "<r>"

        Do While Not oRSfp.EOF
            xFP = xFP & "<d "
            xFP = xFP & "idfp=""" & Trim(str(oRSfp!idformapago)) & """ "
            xFP = xFP & "fp=""" & Trim(oRSfp!formaPAGO) & """ "
            xFP = xFP & "mon=""" & Trim(oRSfp!moneda) & """ "
            xFP = xFP & "monto=""" & Trim(str(oRSfp!monto)) & """ "
            xFP = xFP & "/>"
            oRSfp.MoveNext
        Loop

        xFP = xFP & "</r>"
    End If
    
    Dim oMSN As String

    oMSN = ""

    With oCmdEjec
        .Parameters.Append .CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        .Parameters.Append .CreateParameter("@SERIECOMANDA", adChar, adParamInput, 3, frmComanda.lblSerie.Caption)
        .Parameters.Append .CreateParameter("@NUMEROCOMANDA", adBigInt, adParamInput, , frmComanda.lblNumero.Caption)
        .Parameters.Append .CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        .Parameters.Append .CreateParameter("@Xmlpag", adVarChar, adParamInput, 4000, xFP)
        .Parameters.Append .CreateParameter("@totalfac", adDouble, adParamInput, , val(frmComanda.lblTot.Caption))
        .Parameters.Append .CreateParameter("@usuario", adVarChar, adParamInput, 10, LK_CODUSU)
        
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@exito", adVarChar, adParamOutput, 300, oMSN)
        oCmdEjec.Execute

        oMSN = oCmdEjec.Parameters("@exito").Value
        
    End With
    
     If Len(Trim(oMSN)) <> 0 Then
        MsgBox oMSN, vbCritical, Pub_Titulo
    Else
        MsgBox "Datos Almacenados Correctamente.", vbInformation, Pub_Titulo
        vAcepta = True
        Unload Me
    End If
    
End Sub

Private Sub cmdDscto_Click()

    frmClaveCaja.Show vbModal
    If frmClaveCaja.vAceptar Then
     Dim vS As String
     If VerificaPass(frmClaveCaja.vUSUARIO, frmClaveCaja.vClave, vS) Then
        frmAsigCantFac.txtCantidad.Text = Me.lblDscto.Caption
        frmAsigCantFac.Show vbModal
        If frmAsigCantFac.vAcepta Then
            Me.lblDscto.Caption = frmAsigCantFac.vCANTIDAD
            CalcularImporte
        End If
      Else
       MsgBox "Clave incorrecta", vbCritical, NombreProyecto
      End If
    End If
    
  

   
    
    

   

End Sub

Private Sub cmdFormasPago_Click()
'frmFacComandaFP.gMostrador = xMostrador
frmFacComandaFP2.gDELIVERY = False
'frmFacComandaFP.Show vbModal
frmFacComandaFP2.lblTotalPagar.Caption = FormatCurrency(Me.lblImporte.Caption, 2)
frmFacComandaFP2.Show vbModal

End Sub

Private Sub cmdSunat_Click()

    On Error GoTo cCruc

    Dim p          As Object

    Dim Texto      As String, xTOk As String

    Dim cadena     As String, xvRUC As String

    Dim sInputJson As String, xEsRuc As Boolean

    xEsRuc = True

    MousePointer = vbHourglass
    Set httpURL = New WinHttp.WinHttpRequest
    
    If IsNumeric(Me.txtRS.Text) Then
        If Len(Trim(Me.txtRS.Text)) = 8 Then
            xEsRuc = False
        End If

        xvRUC = Me.txtRS.Text
    Else

        If Len(Trim(Me.txtRS.Text)) = 8 Then
            xEsRuc = False
        End If

        xvRUC = Me.txtRuc.Text
    End If

    xTOk = Leer_Ini(App.Path & "\config.ini", "TOKEN", "")
    
    If xEsRuc Then
        cadena = "http://dniruc.apisperu.com/api/v1/ruc/" & xvRUC & "?token=" & xTOk
    Else
        cadena = "http://dniruc.apisperu.com/api/v1/dni/" & xvRUC & "?token=" & xTOk
    End If
    
    httpURL.Open "GET", cadena
    httpURL.Send
    
    Texto = httpURL.ResponseText

    'sInputJson = "{items:" & Texto & "}"

    Set p = JSON.parse(Texto)

    '    Me.lblRUC.Caption = p.Item("ruc")
    '    Me.lblRazonSocial.Caption = p.Item("razonSocial")
    '    Me.lblDireccion.Caption = p.Item("direccion")
    '    Me.lblTipo.Caption = p.Item("tipo")
    '    Me.lblEstado.Caption = p.Item("estado")
    '    Me.lblcondicion.Caption = p.Item("condicion")
    
    If Len(Trim(Me.txtRuc.Text)) = 0 Then
        If IsNumeric(Me.txtRS.Text) Then
            If Len(Trim(Me.txtRS.Text)) = 11 Or Len(Trim(Me.txtRS.Text)) = 8 Then
                If Texto = "[]" Then
                    MousePointer = vbDefault
                    MsgBox ("No se obtuvo resultados")
                    Me.txtRuc.Text = ""
                    Me.txtRS.Text = ""
                    Me.txtDireccion.Text = ""

                    Exit Sub

                End If

                If Len(Trim(Texto)) = 0 Then
                    MousePointer = vbDefault
                    MsgBox ("No se obtuvo resultados")
                    Me.txtRuc.Text = ""
                    Me.txtRS.Text = ""
                    Me.txtDireccion.Text = ""

                    Exit Sub

                End If

                If xEsRuc Then
                    Me.txtDireccion.Text = IIf(IsNull(p.Item("direccion")), "", p.Item("direccion"))
                    Me.txtRS.Text = p.Item("razonSocial")
                    Me.txtRuc.Text = p.Item("ruc")
                    Me.txtDni.Text = ""
                Else
                    Me.txtRuc.Text = ""
                    Me.txtDireccion.Text = ""
                    Me.txtDni.Text = p.Item("dni")
                    Me.txtRS.Text = p.Item("nombres") & " " & p.Item("apellidoPaterno") & " " & p.Item("apellidoMaterno")
                End If
    
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SP_CLIENTES_UPDATE_DATOS_SUNAT"
                oCmdEjec.CommandType = adCmdStoredProc
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RAZONSOCIAL", adVarChar, adParamInput, 200, Left(Trim(Me.txtRS.Text), 200))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIRECCION", adVarChar, adParamInput, 200, Left(Trim(Me.txtDireccion.Text), 200))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RUC", adVarChar, adParamInput, 11, Me.txtRuc.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DNI", adChar, adParamInput, 8, Me.txtDni.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , 0)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@salida", adBigInt, adParamOutput, , 0)
                oCmdEjec.Execute
                Me.txtRuc.Tag = oCmdEjec.Parameters("@salida").Value
            
            Else
                MsgBox "El ruc debe tener 11 caracteres", vbCritical, Pub_Titulo
            End If

        Else
            MsgBox "El ruc debe ser Numeros", vbCritical, Pub_Titulo
        End If

    Else

        If Texto = "[]" Then
            MousePointer = vbDefault
            MsgBox ("No se obtuvo resultados")
            Me.txtRuc.Text = ""
            Me.txtRS.Text = ""
            Me.txtDireccion.Text = ""

            Exit Sub

        End If

        If Len(Trim(Texto)) = 0 Then
            MousePointer = vbDefault
            MsgBox ("No se obtuvo resultados")
            Me.txtRuc.Text = ""
            Me.txtRS.Text = ""
            Me.txtDireccion.Text = ""

            Exit Sub

        End If
        
        If xEsRuc Then
            Me.txtDni.Text = ""
            'Me.txtDireccion.Text = p.Item("direccion")
            Me.txtDireccion.Text = IIf(IsNull(p.Item("direccion")), "", p.Item("direccion"))
            Me.txtRS.Text = p.Item("razonSocial")
            Me.txtRuc.Text = p.Item("ruc")
        Else
            Me.txtRuc.Text = ""
            Me.txtDireccion.Text = ""
            Me.txtDni.Text = p.Item("dni")
            Me.txtRS.Text = p.Item("nombres") & " " & p.Item("apellidoPaterno") & " " & p.Item("apellidoMaterno")
        End If
    
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_CLIENTES_UPDATE_DATOS_SUNAT"
        oCmdEjec.CommandType = adCmdStoredProc
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    
        'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RAZONSOCIAL", adVarChar, adParamInput, 60, Trim(Me.txtRS.Text))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RAZONSOCIAL", adVarChar, adParamInput, 200, Left(Trim(Me.txtRS.Text), 200))
        'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIRECCION", adVarChar, adParamInput, 50, Left(Trim(Me.txtDireccion.Text), 50))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIRECCION", adVarChar, adParamInput, 200, Left(Trim(Me.txtDireccion.Text), 200))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RUC", adVarChar, adParamInput, 11, Me.txtRuc.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DNI", adChar, adParamInput, 8, Me.txtDni.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, 10, Me.txtRuc.Tag)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@salida", adBigInt, adParamOutput, , 0)
        oCmdEjec.Execute
    End If
       
    MousePointer = vbDefault

    Exit Sub

cCruc:
    MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, Pub_Titulo


End Sub

Sub CargarDocumentos()
Me.DatTiposDoctos.BoundText = ""

 LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_TIPOS_DOCTOS_LIST"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, Me.DatEmpresas.BoundText)
    Set ORStd = oCmdEjec.Execute
    Set Me.DatTiposDoctos.RowSource = ORStd
    Me.DatTiposDoctos.ListField = ORStd.Fields(1).Name
    Me.DatTiposDoctos.BoundColumn = ORStd.Fields(0).Name
    
    

    
    If ORStd.RecordCount <> 0 Then Me.DatTiposDoctos.BoundText = ORStd.Fields(0).Value
    
    'DatTiposDoctos_Click Area
End Sub

Sub cargarSeries()

      If Me.DatEmpresas.BoundText = "" Then Exit Sub

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_SERIES_CARGAR"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim xSerie As String

    Dim xNro   As Double

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, Me.DatEmpresas.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1)) ' Me.cboTipoDocto.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamOutput, 3, 1)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MAXIMO", adBigInt, adParamOutput, , 1)
    oCmdEjec.Execute

    xSerie = oCmdEjec.Parameters("@SERIE").Value
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)

    Me.lblSerie.Caption = xSerie
    Me.txtNro.Text = oCmdEjec.Parameters("@MAXIMO").Value

    ORStd.Filter = "CODIGO='" & Me.DatTiposDoctos.BoundText & "'"

    If ORStd.RecordCount <> 0 Then
        Me.chkEdit.Enabled = ORStd!Editable
    End If

    ORStd.Filter = ""
    Me.txtNro.Enabled = False
    Me.chkEdit.Value = False
End Sub

Private Sub DatEmpresas_Change()
sumatoria
End Sub

Private Sub DatEmpresas_Click(Area As Integer)

CargarDocumentos

    cargarSeries
End Sub


Private Sub DatTiposDoctos_Click(Area As Integer)


    cargarSeries
End Sub

Private Sub DatTiposDoctos_KeyDown(KeyCode As Integer, Shift As Integer)
LimpiaParametros oCmdEjec
    
    oCmdEjec.CommandText = "SP_SERIES_CARGAR"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim xSerie As String

    Dim xNro   As Double

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, Me.DatEmpresas.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1)) ' Me.cboTipoDocto.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamOutput, 3, 1)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MAXIMO", adBigInt, adParamOutput, , 1)
    oCmdEjec.Execute

    xSerie = oCmdEjec.Parameters("@SERIE").Value
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)

    Me.lblSerie.Caption = xSerie
    Me.txtNro.Text = oCmdEjec.Parameters("@MAXIMO").Value

    ORStd.Filter = "CODIGO='" & Me.DatTiposDoctos.BoundText & "'"

    If ORStd.RecordCount <> 0 Then
        Me.chkEdit.Enabled = ORStd!Editable
    End If

    ORStd.Filter = ""
    Me.txtNro.Enabled = False
    Me.chkEdit.Value = False
End Sub

Private Sub DatTiposDoctos_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then If Me.txtNro.Enabled Then Me.txtNro.SetFocus
End Sub

Private Sub DatTiposDoctos_KeyUp(KeyCode As Integer, Shift As Integer)
LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_SERIES_CARGAR"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim xSerie As String

    Dim xNro   As Double

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, Me.DatEmpresas.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1)) ' Me.cboTipoDocto.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamOutput, 3, 1)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MAXIMO", adBigInt, adParamOutput, , 1)
    oCmdEjec.Execute

    xSerie = oCmdEjec.Parameters("@SERIE").Value
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)

    Me.lblSerie.Caption = xSerie
    Me.txtNro.Text = oCmdEjec.Parameters("@MAXIMO").Value

    ORStd.Filter = "CODIGO='" & Me.DatTiposDoctos.BoundText & "'"

    If ORStd.RecordCount <> 0 Then
        Me.chkEdit.Enabled = ORStd!Editable
    End If

    ORStd.Filter = ""
    Me.txtNro.Enabled = False
    Me.chkEdit.Value = False
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Set oRSfp = Nothing
    vAcepta = False
    vBuscar = False
    Me.ListView1.Visible = False
    buscars = True
    ConfiguraLV
    ConfigurarLVDetalle


     LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_CIAS_FACTURACION"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    Set ORStd = oCmdEjec.Execute
    
    Set Me.DatEmpresas.RowSource = ORStd
    Me.DatEmpresas.ListField = ORStd.Fields(1).Name
    Me.DatEmpresas.BoundColumn = ORStd.Fields(0).Name
    
    If ORStd.RecordCount <> 0 Then Me.DatEmpresas.BoundText = ORStd.Fields(0).Value
    'DatEmpresas_Click 1
    
    CargarDocumentos
    cargarSeries
    
   

    LimpiaParametros oCmdEjec
   
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodMesa", adVarChar, adParamInput, 10, vMesa)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)

    If xMostrador Then
        oCmdEjec.CommandText = "SpCargarComanda2"
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 20, LK_CODUSU)
    Else
    
        oCmdEjec.CommandText = "SpCargarComanda"
    End If

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Fac", adBoolean, adParamInput, , 1)

    Set oRsPago = oCmdEjec.Execute '(, Array(LK_CODCIA, vMesa, LK_FECHA_DIA))


    Do While Not oRsPago.EOF

        Set itemX = Me.lvDetalle.ListItems.Add(, , oRsPago!NumFac)
        itemX.SubItems(1) = oRsPago!CODPLATO
        itemX.SubItems(2) = Trim(oRsPago!plato)
        'itemX.SubItems(3) = Trim(oRsPago!cuenta)
        itemX.SubItems(3) = IIf(IsNull(oRsPago!cuenta), "", Trim(oRsPago!cuenta))
        itemX.SubItems(4) = oRsPago!PRECIO
        itemX.SubItems(5) = oRsPago!CantTotal
        itemX.SubItems(6) = oRsPago!faltante
        itemX.SubItems(7) = FormatNumber(oRsPago!Importe, 2)
        itemX.SubItems(8) = oRsPago!SEC
        itemX.SubItems(9) = oRsPago!aPRO
        itemX.SubItems(10) = oRsPago!aten
        itemX.SubItems(11) = oRsPago!uni
        itemX.SubItems(12) = oRsPago!PED_numsec
        itemX.SubItems(14) = oRsPago!icbper
        itemX.SubItems(15) = oRsPago!combo_icbper
                Me.lblpICBPER.Caption = oRsPago!gen_icbper
       
        oRsPago.MoveNext
   
    Loop

  

    Me.lblImporte.Caption = Format(vTOTAL, "########0.#0") 'FormatNumber(vTOTAL, 2)
   
    

   
    Me.lblVuelto.Caption = "0.00"
    Me.chkConsumo.Value = 0

    If LK_CODUSU = "MOZOB" Then
        Me.DatTiposDoctos.BoundText = "01"
        Me.DatTiposDoctos.Locked = True
        'cboTipoDocto.ListIndex = 1
        'cboTipoDocto.Locked = True
    End If

    vBuscar = True
    sumatoria
    'RECORDSET PARA LAS FORMAS DE PAGO

    Set oRSfp = New ADODB.Recordset
    oRSfp.CursorType = adOpenDynamic ' setting cursor type
    oRSfp.Fields.Append "idformapago", adBigInt
    oRSfp.Fields.Append "formapago", adVarChar, 120
    oRSfp.Fields.Append "referencia", adVarChar, 100
    oRSfp.Fields.Append "monto", adDouble
    oRSfp.Fields.Append "tipo", adChar, 1
    oRSfp.Fields.Append "pagacon", adDouble
    oRSfp.Fields.Append "vuelto", adDouble
    oRSfp.Fields.Append "diascredito", adInteger
    
    oRSfp.Fields.Refresh
    oRSfp.Open
    
    oRSfp.AddNew
    oRSfp!idformapago = 1
    oRSfp!formaPAGO = "CONTADO"
    oRSfp!referencia = ""
    oRSfp!monto = Me.lblImporte.Caption
    oRSfp!tipo = "E"
    oRSfp!pagacon = 0
    oRSfp!VUELTO = 0
    oRSfp!diascredito = 0
    oRSfp.Update
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_USUARIO_VERIFICACOBRO"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)

    Dim ORSd As ADODB.Recordset

    Set ORSd = oCmdEjec.Execute

    If Not ORSd.EOF Then
      
        ' Me.cmdFormasPago.Enabled = IIf(ORSd!Fact = "A", True, False)
        Me.cmdCobrar.Enabled = CBool(ORSd!cobra)
       
    End If

Me.DatEmpresas.BoundText = LK_CODCIA
    '    If xMostrador Then Me.DatTiposDoctos.BoundText = "01" ' Me.cboTipoDocto.ListIndex = 1

End Sub

Private Sub CrearArchivoPlano(cTipoDocto As String, cSerie As String, cNumero As Double)
    Dim oRS As ADODB.Recordset

    LimpiaParametros oCmdEjec

    If cTipoDocto = "F" Then
           oCmdEjec.CommandText = "SP_VENTA_FACTURA_SFS"
    ElseIf cTipoDocto = "B" Then
           oCmdEjec.CommandText = "SP_VENTA_BOLETA_SFS"
    ElseIf LK_CODTRA = 1111 Then
    
    End If
    
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, Me.DatEmpresas.BoundText)

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@serie", adVarChar, adParamInput, 3, IIf(LK_CODTRA = 1111, PUB_NUMSER_C, cSerie))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numero", adDouble, adParamInput, , IIf(LK_CODTRA = 1111, PUB_NUMFAC_C, cNumero))
    If LK_CODTRA = 1111 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adChar, adParamInput, 1, PUB_FBG)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TRANSACCION", adBigInt, adParamInput, , LK_CODTRA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    End If
    
    Set oRS = oCmdEjec.Execute
    
    Dim sCadena As String

    sCadena = ""
    
    Dim obj_FSO     As Object

    Dim ArchivoCab  As Object
    Dim ArchivoTri As Object
    Dim ArchivoDet  As Object
    Dim ArchivoLey As Object
    Dim ArchivoAca As Object
    
    Dim sARCHIVOcab As String
    Dim sARCHIVOdet As String
    Dim sARCHIVOtri As String
    Dim sARCHIVOley As String
    Dim sARCHIVOaca As String
    
    Dim sRUC        As String
    
    If Me.DatEmpresas.BoundText = "01" Then
    sRUC = Leer_Ini(App.Path & "\config.ini", "RUC", "C:\")
    ElseIf Me.DatEmpresas.BoundText = "02" Then
    sRUC = Leer_Ini(App.Path & "\config2.ini", "RUC", "C:\")
    ElseIf Me.DatEmpresas.BoundText = "03" Then
    sRUC = Leer_Ini(App.Path & "\config3.ini", "RUC", "C:\")
    Else
    sRUC = Leer_Ini(App.Path & "\config4.ini", "RUC", "C:\")
    End If
     
    sARCHIVOcab = sRUC & "-" & oRS!Nombre + IIf(LK_CODTRA = 2412, ".not", IIf(LK_CODTRA = 1111, ".cba", ".cab"))
        
    If LK_CODTRA <> 1111 Then
        sARCHIVOdet = sRUC & "-" & oRS!Nombre + ".det"
        sARCHIVOtri = sRUC & "-" & oRS!Nombre + ".tri" 'IIf(LK_CODTRA = 2412, ".not", IIf(LK_CODTRA = 1111, ".tri", ".tri"))
        sARCHIVOley = sRUC & "-" & oRS!Nombre + ".ley" 'IIf(LK_CODTRA = 2412, ".not", IIf(LK_CODTRA = 1111, ".ley", ".ley"))
        sARCHIVOaca = sRUC & "-" & oRS!Nombre + ".aca"
        If cTipoDocto = "F" Then 'es factura
            sARCHIVOpag = sRUC & "-" & oRS!Nombre + ".pag"
            sARCHIVOdpa = sRUC & "-" & oRS!Nombre + ".dpa"
            sARCHIVOrtn = sRUC & "-" & oRS!Nombre + ".rtn"
        End If
    End If
    
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")

    'Creamos un archivo con el método CreateTextFile
    If Me.DatEmpresas.BoundText = "01" Then
        Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
        Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
        Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOley, True)
        Set ArchivoAca = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOaca, True)
        If cTipoDocto = "F" Then
            Set ArchivoPAG = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOpag, True)
        End If
    ElseIf Me.DatEmpresas.BoundText = "02" Then
        Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
        Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
        Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOley, True)
        Set ArchivoAca = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOaca, True)
        If cTipoDocto = "F" Then
           Set ArchivoPAG = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOpag, True)
        End If
    ElseIf Me.DatEmpresas.BoundText = "03" Then
        Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
        Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
        Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOley, True)
        Set ArchivoAca = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOaca, True)
        If cTipoDocto = "F" Then
           Set ArchivoPAG = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOpag, True)
        End If
    ElseIf Me.DatEmpresas.BoundText = "04" Then
        Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
        Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
        Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOley, True)
        Set ArchivoAca = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOaca, True)
        If cTipoDocto = "F" Then
           Set ArchivoPAG = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOpag, True)
        End If
    End If
    If LK_CODTRA <> 1111 Then
   If Me.DatEmpresas.BoundText = "01" Then
    Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
    ElseIf Me.DatEmpresas.BoundText = "02" Then
    Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
    ElseIf Me.DatEmpresas.BoundText = "03" Then
    Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
    Else
    Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
    End If
    End If
    
    
    If LK_CODTRA = 2412 Then

        Do While Not oRS.EOF
            sCadena = sCadena & oRS!fecemision & "|" & oRS!CODMOTIVO & "|" & oRS!DESCMOTIVO & "|" & oRS!TIPODOCAFECTADO & "|" & oRS!NUMDOCAFECTADO & "|" & oRS!TIPDOCUSUARIO & "|" & oRS!NUMDOCUSUARIO & "|" & oRS!CLI1 & "|" & oRS!TIPMONEDA & "|" & oRS!SUMOTROSCARGOS & "|" & oRS!MTOOPERGRAVADAS & "|" & oRS!MTOOPERINAFECTAS & "|" & oRS!MTOOPEREXONERADAS & "|" & oRS!MTOIGV & "|" & oRS!MTOISC & "|" & oRS!MTOOTROSTRIBUTOS & "|" & oRS!MTOIMPVENTA & "|"
            oRS.MoveNext
        Loop
    
    ElseIf LK_CODTRA = 1111 Then
         Do While Not oRS.EOF
            sCadena = sCadena & oRS!FEC_GENERACcION & "|" & oRS!FEC_COMUNICACION & "|" & oRS!TIPDOCBAJA & "|" & oRS!NUMDOCBAJA & "|" & oRS!DESMOTIVOBAJA & "|"
            oRS.MoveNext
        Loop
    Else

        Do While Not oRS.EOF
            sCadena = sCadena & oRS!TIPOPERACION & "|" & oRS!fecemision & "|" & oRS!hORA & "|" & oRS!FECHAVENC & "|" & oRS!codlocalemisor & "|" & oRS!TIPDOCUSUARIO & "|" & oRS!NUMDOCUSUARIO & "|" & oRS!rznsocialusuario & "|" & oRS!TIPMONEDA & "|" & oRS!MTOIGV & "|" & oRS!MTOOPERGRAVADAS & "|" & oRS!MTOIMPVENTA & "|" & oRS!SUMDSCTOGLOBAL & "|" & oRS!SUMOTROSCARGOS & "|" & oRS!TOTANTICIPOS & "|" & oRS!IMPTOTALVENTA & "|" & oRS!UBL & "|" & oRS!CUSTOMDOC & "|"
         oRS.MoveNext
        Loop

    End If
   
    'Escribimos lineas
    ArchivoCab.WriteLine sCadena
    
    'Cerramos el fichero
    ArchivoCab.Close
    Set ArchivoCab = Nothing
    
    If LK_CODTRA <> 1111 Then
         'DIRECCION
    oRS.MoveFirst
    sCadena = ""
    Do While Not oRS.EOF
        sCadena = sCadena & oRS!ACA1 & "|" & oRS!ACA2 & "|" & oRS!ACA3 & "|" & oRS!ACA4 & "|" & oRS!ACA5 & "|" & oRS!PAIS & "|" & oRS!UBIGEO & "|" & oRS!dir & "|" & oRS!PAIS1 & "|" & oRS!UBIGEO1 & "|" & oRS!dir1 & "|"
        oRS.MoveNext
    Loop
    
    'Escribimos LINEAS
    ArchivoAca.WriteLine sCadena
    
    'Cerramos el fichero
    ArchivoAca.Close
    Set ArchivoAca = Nothing
    Else
    End If
    
   
    Dim oRSdet As ADODB.Recordset

    Set oRSdet = oRS.NextRecordset
   
    sCadena = ""
    Dim c As Integer
    c = 1

    If LK_CODTRA = 2412 Then

        Do While Not oRSdet.EOF
         
            sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & oRSdet!DESITEM & "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!CODTIPTRIBUTOIGV & "|" & oRSdet!MTOIGVITEM1 & "|" & oRSdet!NOMTRIBITEM & "|" & oRSdet!CODTIPTRIBUTOITEM & "|" & oRSdet!TIPAFEIGV & "|" & FormatNumber(oRSdet!PORCIGV, 2) & "|" & oRSdet!CODISC & "|" & oRSdet!CODOTROITEM & "|" & oRSdet!GRATUITO & "|"
            
            If c < oRSdet.RecordCount Then
                sCadena = sCadena + vbCrLf
            End If
             c = c + 1
            oRSdet.MoveNext
            
        Loop

    ElseIf LK_CODTRA <> 1111 Then
    

        Do While Not oRSdet.EOF
       
           ' sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & oRSdet!DESITEM & "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTODSCTOITEM & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!TIPAFEIGV & "|" & oRSdet!MTOISCITEM & "|" & oRSdet!TIPSISISC & "|" & oRSdet!MTOPRECIOVENTAITEM & "|" & oRSdet!MTOVALORVENTAITEM & "|"
           sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & Trim(oRSdet!DESITEM) & _
           "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!CODTIPTRIBUTOIGV & "|" & oRSdet!MTOIGVITEM1 & "|" & oRSdet!BASEIMPIGV & "|" & _
           oRSdet!NOMTRIBITEM & "|" & oRSdet!CODTIPTRIBUTOITEM & "|" & oRSdet!TIPAFEIGV & "|" & FormatNumber(oRSdet!PORCIGV, 2) & "|" & oRSdet!CODISC & "|" & oRSdet!MONTOISC & _
           "|" & oRSdet!BASEIMPONIBLEISC & "|" & oRSdet!NOMBRETRIBITEM & "|" & oRSdet!CODTRIBITEM & "|" & oRSdet!CODSISISC & "|" & oRSdet!PORCISC & "|" & oRSdet!CODTRIBOTO & _
           "|" & oRSdet!MONTOTRIBOTO & "|" & oRSdet!BASEIMPONIBLEOTO & "|" & oRSdet!NOMBRETRIBOTO & "|" & oRSdet!TIPSISISC & "|" & oRSdet!PORCOTO & "|" & oRSdet!CODIGOICBPER & _
           "|" & oRSdet!IMPORTEICBPER & "|" & oRSdet!CANTIDADICBPER & "|" & oRSdet!TITULOICBPER & "|" & oRSdet!IDEICBPER & "|" & oRSdet!MONTOICBPER & "|" & _
           oRSdet!PRECIOVTAUNITARIO & "|" & oRSdet!VALORVTAXITEM & "|" & oRSdet!GRATUITO & "|"
            If c < oRSdet.RecordCount Then
                sCadena = sCadena + vbCrLf
            End If
             c = c + 1
            oRSdet.MoveNext
             
        Loop

    End If

    'Escribimos lineas
    If LK_CODTRA <> 1111 Then
    ArchivoDet.WriteLine sCadena
    
     'Cerramos el fichero
    ArchivoDet.Close
    Set ArchivoDet = Nothing
    
    Dim orsTri As ADODB.Recordset
    Set orsTri = oRS.NextRecordset
    
    sCadena = ""
    c = 1
    'ARCIVO .TRI
    Do While Not orsTri.EOF
    sCadena = sCadena & orsTri!Codigo & "|" & orsTri!Nombre & "|" & orsTri!cod & "|" & orsTri!BASEIMPONIBLE & "|" & orsTri!TRIBUTO & "|"
    If c < orsTri.RecordCount Then
        sCadena = sCadena & vbCrLf
    End If
    c = c + 1
        orsTri.MoveNext
    Loop
    
    
     ArchivoTri.WriteLine sCadena
    
     'Cerramos el fichero
    ArchivoTri.Close
    Set ArchivoTri = Nothing
    
    Dim orsLey As ADODB.Recordset
    Set orsLey = oRS.NextRecordset
    
    c = 1
    sCadena = ""
    Do While Not orsLey.EOF
        sCadena = sCadena & orsLey!cod & "|" & Trim(CONVER_LETRAS(Me.lblImporte.Caption, "S")) & "|"
        If c < orsLey.RecordCount Then
            sCadena = sCadena & vbCrLf
        End If
        c = c + 1
        orsLey.MoveNext
    Loop
    
    ArchivoLey.WriteLine sCadena
    ArchivoLey.Close
    Set ArchivoLey = Nothing
    
    Dim xFormaPago As String
    If cTipoDocto = "F" Then
            'PAG
            Dim orsPAG As ADODB.Recordset
            Set orsPAG = oRS.NextRecordset
            
            c = 1
            sCadena = ""
            Do While Not orsPAG.EOF
                xFormaPago = orsPAG!formaPAGO
                sCadena = sCadena & orsPAG!formaPAGO & "|" & orsPAG!pendientepago & "|" & orsPAG!TIPMONEDA & "|"
                If c < orsPAG.RecordCount Then
                    sCadena = sCadena & vbCrLf
                End If
                c = c + 1
                orsPAG.MoveNext
            Loop
            
            ArchivoPAG.WriteLine sCadena
            ArchivoPAG.Close
            Set ArchivoPAG = Nothing
            
            'DPA
            Dim orsDPA As ADODB.Recordset
            Set orsDPA = oRS.NextRecordset
            If UCase(xFormaPago) = "CREDITO" Or UCase(xFormaPago) = "CRÉDITO" Then
                Set ArchivoDPA = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOdpa, True)
               
                
                c = 1
                sCadena = ""
                Do While Not orsDPA.EOF
                    sCadena = sCadena & orsDPA!cuotapago & "|" & orsDPA!fechavcto & "|" & orsDPA!TIPMONEDA & "|"
                    If c < orsDPA.RecordCount Then
                        sCadena = sCadena & vbCrLf
                    End If
                    c = c + 1
                    orsDPA.MoveNext
                Loop
                
                ArchivoDPA.WriteLine sCadena
                ArchivoDPA.Close
                Set ArchivoDPA = Nothing
            End If
             'RTN
'            Dim orsRTN As ADODB.Recordset
'            Set orsRTN = oRS.NextRecordset
'
'            c = 1
'            sCadena = ""
'            Do While Not orsRTN.EOF
'                sCadena = sCadena & orsRTN!impoperacion & "|" & orsRTN!porretencion & "|" & orsRTN!impretencion & "|"
'                If c < orsRTN.RecordCount Then
'                    sCadena = sCadena & vbCrLf
'                End If
'                c = c + 1
'                orsRTN.MoveNext
'            Loop
'
'            ArchivoRTN.WriteLine sCadena
'            ArchivoRTN.Close
'            Set ArchivoRTN = Nothing
        End If
    
    End If
    
   
    
    Set obj_FSO = Nothing
    
End Sub

Private Sub CalcularImporte()
'1combo
'2 combo
'3 combo

Dim vp1  As Currency, vp2 As Currency, vp3 As Currency

Dim vDol As Currency

'If Me.cboMoneda1.ListIndex = 0 Then 'soles
'    vp1 = val(Me.txtMoney1.Text)
'ElseIf Me.cboMoneda1.ListIndex = 1 Then 'dolares
'    vp1 = LK_TIPO_CAMBIO * val(Me.txtMoney1.Text)
'    'vp1 = vDol - val(Me.lblImporte.Caption)
'End If

Dim Item As Object
Dim icbper As Double
vp1 = 0
icbper = 0

For Each Item In Me.lvDetalle.ListItems

    vp1 = vp1 + Item.SubItems(7)
    If Item.SubItems(14) = 1 Then
        icbper = icbper + (Item.SubItems(6) * Me.lblpICBPER.Caption)
    End If
    
    If Item.SubItems(15) > 0 Then
         icbper = icbper + Item.SubItems(15)
    End If
    
Next

'If Me.cboMoneda1.ListIndex = 1 Then 'dolares
If LK_MONEDA = "D" Then
    vp1 = LK_TIPO_CAMBIO * vp1
'vp1 = vDol - val(Me.lblImporte.Caption)
End If

vp1 = vp1 - val(Me.lblDscto.Caption)
Me.lblicbper.Caption = FormatNumber(icbper, 2)

'If Me.txtMoney1.Text <> 0 Then
    Me.lblImporte.Caption = vp1
    Me.lblVuelto.Caption = val(Me.lblImporte.Caption) - vp1
'End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
gDESCUENTO = 0
gPAGO = 0

End Sub



Private Sub lvDetalle_DblClick()
frmFaccomandaOtroPlato.Show vbModal
End Sub

Private Sub lvDetalle_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        If Not Me.lvDetalle.SelectedItem Is Nothing Then
            Me.lvDetalle.ListItems.Remove Me.lvDetalle.SelectedItem.index
            CalcularImporte
            sumatoria
      
           ' Me.txtMoney1.Text = Me.lblImporte.Caption
            
            If oRSfp.RecordCount <> 0 Then
                oRSfp.MoveFirst
                
                Do While Not oRSfp.EOF
                    oRSfp.Delete
                    oRSfp.MoveNext
                Loop
                
            End If
            
            oRSfp.AddNew
            oRSfp!idformapago = 1
            oRSfp!formaPAGO = "CONTADO"
            oRSfp!referencia = ""
            oRSfp!tipo = "E"
            oRSfp!monto = Me.lblImporte.Caption
            oRSfp!diascredito = 0
            oRSfp.Update
      
        End If
    End If

End Sub

Private Sub sumatoria()
Dim vIgv As Integer
 LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "USP_EMPRESA_IGV"
            Dim orsIGV As ADODB.Recordset
            Set orsIGV = oCmdEjec.Execute(, Me.DatEmpresas.BoundText)

            Dim Item As Object
        
            If Not orsIGV.EOF Then
            vIgv = orsIGV.Fields(0).Value
            Me.lblporcigv.Caption = vIgv & "%"
            End If

    Dim vimp As Double

    For i = 1 To Me.lvDetalle.ListItems.count
        vimp = vimp + Me.lvDetalle.ListItems(i).SubItems(7)
       
    Next

    'Me.lblImporte.Caption = Format(vimp, "########0.#0") 'FormatNumber(vimp, 2)
     Me.lblvvta.Caption = Round(vimp / ((vIgv / 100) + 1), 2)
        Me.lblIgv.Caption = vimp - Me.lblvvta.Caption
        Me.lblImporte.Caption = Format(val(vimp) + val(Me.lblicbper.Caption), "########0.#0")
        

    
End Sub

Private Sub lvDetalle_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        frmAsigCantFac.txtCantidad.Text = Me.lvDetalle.SelectedItem.SubItems(5)
        frmAsigCantFac.Show vbModal

        If frmAsigCantFac.vAcepta Then
            If frmAsigCantFac.vCANTIDAD > Me.lvDetalle.SelectedItem.SubItems(5) Then
                MsgBox "La cantidad supera la actual", vbInformation, "Error"

                Exit Sub

            End If

            Me.lvDetalle.SelectedItem.SubItems(6) = frmAsigCantFac.vCANTIDAD
            Me.lvDetalle.SelectedItem.SubItems(7) = Me.lvDetalle.SelectedItem.SubItems(6) * Me.lvDetalle.SelectedItem.SubItems(4)
            CalcularImporte
            sumatoria
      
            If oRSfp.RecordCount <> 0 Then
                oRSfp.MoveFirst
                
                Do While Not oRSfp.EOF
                    oRSfp.Delete
                    oRSfp.MoveNext
                Loop
                
            End If
            
    oRSfp.AddNew
    oRSfp!idformapago = 1
    oRSfp!formaPAGO = "CONTADO"
    oRSfp!referencia = ""
    oRSfp!monto = Me.lblImporte.Caption
    oRSfp!tipo = "E"
    oRSfp.Update
            
        End If

    End If

End Sub

Private Sub txtDni_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtMoney1_Change()
CalcularImporte
End Sub

Private Sub txtMoney2_Change()
CalcularImporte
End Sub

Private Sub txtMoney3_Change()
CalcularImporte
End Sub

Private Sub txtnro_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.txtRS.SetFocus
'If KeyAscii = vbKeyReturn And Me.DatTiposDoctos.BoundText = "01" Then Me.txtRS.SetFocus
'If KeyAscii = vbKeyReturn And Me.DatTiposDoctos.BoundText = "03" Then Me.txtcli.SetFocus

End Sub

Private Sub txtRS_Change()
vBuscar = True
'txtRuc.Tag = ""
'If buscars Then
'Me.ListView1.Visible = True
'If Len(Trim(Me.txtRS.Text)) <> 0 Then
'For i = 1 To Me.ListView1.ListItems.count
'    If Me.txtRS.Text = Left(Me.ListView1.ListItems(i).SubItems(1), Len(Me.txtRS.Text)) Then
'        Me.ListView1.ListItems(i).Selected = True
'        loc_key = i
'        Me.ListView1.ListItems(i).EnsureVisible
'        Exit For
'    Else
'        Me.ListView1.ListItems(i).Selected = False
'        loc_key = -1
'
'    End If
'Next
'Else
'Me.ListView1.Visible = False
'Me.txtRuc.Text = ""
'Me.txtDireccion.Text = ""
'Me.txtRS.Text = ""
'End If
'End If
End Sub

Private Sub txtRS_GotFocus()
buscars = True
End Sub

Private Sub txtRS_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key = loc_key + 1

        If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 38 Then
        loc_key = loc_key - 1

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 34 Then
        loc_key = loc_key + 17

        If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 33 Then
        loc_key = loc_key - 17

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 27 Then
        Me.ListView1.Visible = False
        Me.txtRS.Text = ""
        Me.txtRuc.Text = ""
        Me.txtDireccion.Text = ""
    End If

    GoTo fin
posicion:
    ListView1.ListItems.Item(loc_key).Selected = True
    ListView1.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    txtRS.SelStart = Len(txtRS.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtRS_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscar Then
            Me.ListView1.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpListarCliProv"
            Set oRsPago = oCmdEjec.Execute(, Array(Me.DatEmpresas.BoundText, "C", Me.txtRS.Text))

            Dim Item As Object
        
            If Not oRsPago.EOF Then

                Do While Not oRsPago.EOF
                    Set Item = Me.ListView1.ListItems.Add(, , oRsPago!CodClie)
                    Item.SubItems(1) = Trim(oRsPago!Nombre)
                    Item.SubItems(2) = IIf(IsNull(oRsPago!RUC), "", oRsPago!RUC)
                    Item.SubItems(3) = Trim(oRsPago!dir)
                    Item.Tag = oRsPago!DNI
                    oRsPago.MoveNext
                Loop

                Me.ListView1.Visible = True
                Me.ListView1.ListItems(1).Selected = True
                loc_key = 1
                Me.ListView1.ListItems(1).EnsureVisible
                vBuscar = False
            Else

                If MsgBox("Cliente no existe." + vbCrLf + "¿Desea Crearlo.?", vbQuestion + vbYesNo, "Restaurantes") = vbYes Then
                    frmCLI.Show vbModal
                End If
            End If
        
        Else
            
            Me.txtRuc.Text = Me.ListView1.ListItems(loc_key).SubItems(2)
            Me.txtDireccion.Text = Me.ListView1.ListItems(loc_key).SubItems(3)
            Me.txtRS.Text = Me.ListView1.ListItems(loc_key).SubItems(1)
            Me.ListView1.Visible = False
            Me.txtDni.Text = Me.ListView1.ListItems(loc_key).Tag
            Me.txtRuc.Tag = Me.ListView1.ListItems(loc_key)
            Me.lvDetalle.SetFocus
        End If
    End If

End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
'Me.ListView1.Visible = True
If Len(Trim(Me.txtRuc.Text)) <> 0 Then
For i = 1 To Me.ListView1.ListItems.count
    If Trim(Me.txtRuc.Text) = Trim(Me.ListView1.ListItems(i).SubItems(2)) Then
        'Me.ListView1.ListItems(i).Selected = True
        buscars = False
        loc_key = i
        Me.ListView1.ListItems(i).EnsureVisible
        Me.txtRS.Text = Me.ListView1.ListItems(i).SubItems(1)
        Me.txtDireccion.Text = Me.ListView1.ListItems(i).SubItems(3)
        Me.txtRuc.Tag = Me.ListView1.ListItems(i)
        Exit For
    Else
       ' Me.ListView1.ListItems(i).Selected = False
        loc_key = -1
        
    End If
Next
Else
Me.ListView1.Visible = False
Me.txtRuc.Text = ""
Me.txtDireccion.Text = ""
Me.txtRS.Text = ""
End If
End If
End Sub
