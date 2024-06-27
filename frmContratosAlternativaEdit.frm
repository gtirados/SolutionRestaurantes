VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmContratosAlternativaEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editando Alternativa de Contrato"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8205
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
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
            Caption         =   "&Aceptar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9240
      Top             =   2280
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
            Picture         =   "frmContratosAlternativaEdit.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratosAlternativaEdit.frx":039A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1440
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   7935
      Begin MSDataListLib.DataCombo DatDsctos 
         Height          =   315
         Left            =   1620
         TabIndex        =   1
         Top             =   600
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1620
         TabIndex        =   0
         Top             =   240
         Width           =   6135
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   1620
         TabIndex        =   2
         Top             =   990
         Width           =   1215
      End
      Begin VB.TextBox txtPrecio 
         Height          =   285
         Left            =   4440
         TabIndex        =   3
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPCIÓN:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCUENTO:"
         Height          =   195
         Left            =   420
         TabIndex        =   8
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANTIDAD:"
         Height          =   195
         Left            =   420
         TabIndex        =   7
         Top             =   1035
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIO:"
         Height          =   195
         Left            =   3600
         TabIndex        =   6
         Top             =   1035
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmContratosAlternativaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vPunto As Boolean
Private ORS As ADODB.Recordset

Private Sub DatDsctos_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.txtCantidad.SetFocus
End Sub

Private Sub Form_Load()
vAcepta = False
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SPMAESTRODESCUENTOS"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
Set ORS = oCmdEjec.Execute
With Me.DatDsctos
Set .RowSource = ORS
    .ListField = "DESCUENTO"
    .BoundColumn = "IDE"
End With


With frmContratosAlternativas.lvAlternativas.SelectedItem
    Me.txtDescripcion.Text = .Text
    Me.DatDsctos.BoundText = .SubItems(1)
    Me.txtCantidad.Text = .SubItems(3)
    Me.txtPrecio.Text = .SubItems(4)
End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index

    Case 1: AceptarInformacion
    Case 2: Unload Me

End Select
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
  If SoloNumeros(KeyAscii) Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then
    Me.txtPrecio.SelStart = 0
    Me.txtPrecio.SelLength = Len(Me.txtPrecio.Text)
    Me.txtPrecio.SetFocus
    End If
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then Me.DatDsctos.SetFocus
        

End Sub



Private Sub txtPrecio_Change()
 If InStr(txtPrecio.Text, ".") = 0 Then
        vPunto = False
    Else
        vPunto = True
    End If
End Sub

Private Sub txtPrecio_GotFocus()
 If InStr(txtPrecio.Text, ".") = 0 Then
        vPunto = False
    Else
        vPunto = True
    End If
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
If NumerosyPunto(KeyAscii) Then KeyAscii = 0
  If Chr(KeyAscii) = "." Then
        If vPunto = True Or Len(Trim(txtPrecio.Text)) = 0 Then
            KeyAscii = 0
        End If
    End If
If KeyAscii = vbKeyReturn Then AceptarInformacion
End Sub

Private Sub AceptarInformacion()

    If Len(Trim(Me.txtDescripcion.Text)) = 0 Then
        MsgBox "Debe ingresar la Descripcón de la Alternativa.", vbCritical, TituloSistema
        Me.txtDescripcion.SetFocus
    ElseIf Len(Trim(Me.txtCantidad.Text)) = 0 Then
        MsgBox "Debe ingresar la Cantidad de la Alternativa.", vbCritical, TituloSistema
        Me.txtCantidad.SetFocus
    ElseIf val(Me.txtCantidad.Text) = 0 Or Not IsNumeric(Me.txtCantidad.Text) Then
        MsgBox "La Cantidad ingresada no es correcta.", vbCritical, TituloSistema
        Me.txtCantidad.SetFocus
        Me.txtCantidad.SelStart = 0
        Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
    ElseIf Len(Trim(Me.txtPrecio.Text)) = 0 Then
        MsgBox "Debe ingresar el precio para la Alternativa.", vbCritical, TituloSistema
        Me.txtPrecio.SetFocus
    ElseIf val(Me.txtPrecio.Text) = 0 Or Not IsNumeric(Me.txtPrecio.Text) Then
        MsgBox "El precio ingresado es incorrecto.", vbCritical, TituloSistema
        Me.txtPrecio.SetFocus
        Me.txtPrecio.SelStart = 0
        Me.txtPrecio.SelLength = Len(Me.txtPrecio.Text)
    Else
    Dim vVALOR As Double
    Dim vdscto As Double, vVALORDSCTO As Double
    vdscto = 0
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPOBTENERDESCUENTO"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDDSCTO", adDouble, adParamInput, , IIf(Me.DatDsctos.BoundText = "", -1, Me.DatDsctos.BoundText))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@VALOR", adDouble, adParamOutput, , 0)
    oCmdEjec.Execute
    
    vdscto = oCmdEjec.Parameters("@VALOR").Value
    
    'vdscto = IIf(Len(Trim(Me.txtDescuento.Text)) <> 0, Me.txtDescuento.Text, 0)
    
        With frmContratosAlternativas.lvAlternativas.SelectedItem
            .Text = Me.txtDescripcion.Text
            .SubItems(1) = Me.DatDsctos.BoundText
            .SubItems(2) = vdscto
            vVALOR = val(Me.txtCantidad.Text) * val(Me.txtPrecio.Text)
    vVALORDSCTO = (vVALOR * vdscto) / 100
            .SubItems(3) = Me.txtCantidad.Text
            .SubItems(4) = Me.txtPrecio.Text
            .SubItems(5) = vVALOR
            .SubItems(6) = vVALOR - vVALORDSCTO
        End With
        frmContratosAlternativas.oRSAlt.Filter = "CODALTERNATIVA=" & frmContratosAlternativas.lvAlternativas.SelectedItem.Tag
        If Not frmContratosAlternativas.oRSAlt.EOF Then
            frmContratosAlternativas.oRSAlt.Fields!ALTERNATIVA = Me.txtDescripcion.Text
            frmContratosAlternativas.oRSAlt!IDDSCTO = IIf(Me.DatDsctos.BoundText = "", -1, Me.DatDsctos.BoundText)
            frmContratosAlternativas.oRSAlt.Fields!DESCUENTO = vdscto
            'vdscto
            frmContratosAlternativas.oRSAlt.Fields!Cantidad = Me.txtCantidad.Text
            frmContratosAlternativas.oRSAlt.Fields!PRECIO = Me.txtPrecio.Text
            vVALOR = val(Me.txtCantidad.Text) * val(Me.txtPrecio.Text)
            'vdscto = (vVALOR * IIf(Len(Trim(Me.txtDescuento.Text)) <> 0, Me.txtDescuento.Text, 0)) / 100
            frmContratosAlternativas.oRSAlt!NETO = vVALOR
            frmContratosAlternativas.oRSAlt!BRUTO = vVALOR - vVALORDSCTO
        End If
        
        Unload Me
        
    End If
    
End Sub
