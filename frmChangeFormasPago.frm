VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmChangeFormasPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar formas de pago"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8730
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
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   8730
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChange 
      Caption         =   "Cambiar FP"
      Height          =   600
      Left            =   7440
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   360
      Left            =   6240
      TabIndex        =   7
      Top             =   480
      Width           =   990
   End
   Begin MSComctlLib.ListView lvListado 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4260
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
   Begin VB.TextBox txtNumero 
      Height          =   285
      Left            =   4560
      TabIndex        =   2
      Top             =   495
      Width           =   1215
   End
   Begin VB.TextBox txtSerie 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   495
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo DatTipoDocto 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "TipoDocto"
   End
   Begin VB.Label lblTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5640
      TabIndex        =   10
      Top             =   3600
      Width           =   1515
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL:"
      Height          =   195
      Left            =   4920
      TabIndex        =   9
      Top             =   3600
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÚMERO:"
      Height          =   195
      Left            =   4560
      TabIndex        =   5
      Top             =   240
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SERIE:"
      Height          =   195
      Left            =   2880
      TabIndex        =   4
      Top             =   240
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO DOCTO:"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   1200
   End
End
Attribute VB_Name = "frmChangeFormasPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdBuscar_Click()
If Me.DatTipoDocto.BoundText = "" Then
    MsgBox "Debe elegir el Tipo de Documento.", vbInformation, Pub_Titulo
    Me.DatTipoDocto.SetFocus
    Exit Sub
End If
If Len(Trim(Me.txtSerie.Text)) = 0 Then
    MsgBox "Debe ingresar la serie del Documento.", vbCritical, Pub_Titulo
    Me.txtSerie.SetFocus
    Exit Sub
End If
If Len(Trim(Me.txtNumero.Text)) = 0 Then
    MsgBox "Debe ingresar el número del Documento.", vbCritical, Pub_Titulo
    Me.txtNumero.SetFocus
    Exit Sub
End If

LimpiaParametros oCmdEjec
Me.lvListado.ListItems.Clear
oCmdEjec.CommandText = "[dbo].[USP_DOCTOVENTAS_FORMASPAGO]"
Set oRSmain = oCmdEjec.Execute(, Array(LK_CODCIA, Left(Me.DatTipoDocto.Text, 1), Me.txtSerie.Text, Me.txtNumero.Text))
Dim itemX As Object
If Not oRSmain.EOF Then
Me.lblTotal.Caption = oRSmain!Total
    Do While Not oRSmain.EOF
        Set itemX = Me.lvListado.ListItems.Add(, , oRSmain!serie)
        itemX.Tag = oRSmain!idformapago
        itemX.SubItems(1) = oRSmain!NUMERO
        itemX.SubItems(2) = oRSmain!fecha
        itemX.SubItems(3) = oRSmain!formapago
        itemX.SubItems(4) = oRSmain!Importe
        itemX.SubItems(5) = oRSmain!fbg
        itemX.SubItems(6) = oRSmain!correlativo
        oRSmain.MoveNext
    Loop
Else
    MsgBox "No se encontraron registros", vbInformation, Pub_Titulo
End If
End Sub

Private Sub cmdChange_Click()
If Me.lvListado.ListItems.count = 0 Then Exit Sub

'LimpiaParametros oCmdEjec
'oCmdEjec.CommandText = "[dbo].[USP_FORMASPAGO_VALIDA]"
'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Tipodocto", adChar, adParamInput, 1, Me.lvListado.SelectedItem.SubItems(5))
'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@serie", adVarChar, adParamInput, 3, Trim(Me.lvListado.SelectedItem.Text))
'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numero", adBigInt, adParamInput, , Me.lvListado.SelectedItem.SubItems(1))
'
'Dim orsmdu As ADODB.Recordset
'
'Set orsmdu = oCmdEjec.Execute
'
'If Not orsmdu.EOF Then
'    If orsmdu!masdeuno Then
'        MsgBox "No procede si tiene mas de 1 forma de pago", vbCritical, Pub_Titulo
'    Else
    frmChangeFormasPagoEDIT.lblSerie.Caption = Me.lvListado.SelectedItem.Text
frmChangeFormasPagoEDIT.lblNumero.Caption = Me.lvListado.SelectedItem.SubItems(1)
frmChangeFormasPagoEDIT.lblFecha.Caption = Me.lvListado.SelectedItem.SubItems(2)
frmChangeFormasPagoEDIT.lblFormaPago.Caption = Me.lvListado.SelectedItem.SubItems(3)
frmChangeFormasPagoEDIT.lblNumOper.Caption = Me.lvListado.SelectedItem.Tag
frmChangeFormasPagoEDIT.lblFBG.Caption = Me.lvListado.SelectedItem.SubItems(5)
frmChangeFormasPagoEDIT.lblImporte.Caption = Me.lvListado.SelectedItem.SubItems(4)
frmChangeFormasPagoEDIT.lblCorrelativo.Caption = Me.lvListado.SelectedItem.SubItems(6)
frmChangeFormasPagoEDIT.Show vbModal
If frmChangeFormasPagoEDIT.gAcepta Then cmdBuscar_Click




End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SP_TIPOS_DOCTOS_LIST"
Set oRSmain = oCmdEjec.Execute(, LK_CODCIA)

Set Me.DatTipoDocto.RowSource = oRSmain
Me.DatTipoDocto.ListField = oRSmain.Fields(1).Name
Me.DatTipoDocto.BoundColumn = oRSmain.Fields(0).Name
ConfigurarLV

End Sub

Private Sub ConfigurarLV()
With Me.lvListado
    .Gridlines = True
    .LabelEdit = lvwManual
    .FullRowSelect = True
    .View = lvwReport
    
    .HideSelection = False
    .ColumnHeaders.Add , , "Serie", 800
    .ColumnHeaders.Add , , "Número"
    .ColumnHeaders.Add , , "Fecha", 1300
    .ColumnHeaders.Add , , "Forma Pago", 1800
    .ColumnHeaders.Add , , "Importe", 1000, 1
    .ColumnHeaders.Add , , "FBG", 0
    .ColumnHeaders.Add , , "correlativo", 0

End With
End Sub

Private Sub txtnumero_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
End Sub
