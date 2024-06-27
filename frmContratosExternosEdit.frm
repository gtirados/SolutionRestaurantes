VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContratosExternosEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editando Item de Contrato"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6825
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
   ScaleHeight     =   1410
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   120
      TabIndex        =   4
      Top             =   380
      Width           =   6615
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtPrecio 
         Height          =   285
         Left            =   5280
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtDescExterno 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIO:"
         Height          =   195
         Left            =   4440
         TabIndex        =   7
         Top             =   645
         Width           =   750
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANTIDAD:"
         Height          =   195
         Left            =   435
         TabIndex        =   6
         Top             =   645
         Width           =   1020
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPCIÓN:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   1080
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
            Picture         =   "frmContratosExternosEdit.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratosExternosEdit.frx":039A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   635
      ButtonWidth     =   2011
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
End
Attribute VB_Name = "frmContratosExternosEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
If Len(Trim(Me.txtDescExterno.Text)) <> 0 Then
Me.txtDescExterno.SetFocus
    Me.txtDescExterno.SelStart = 0
    Me.txtDescExterno.SelLength = Len(Me.txtDescExterno.Text)
End If
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
 If SoloNumeros(KeyAscii) Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Me.txtPrecio.SetFocus
End Sub

Private Sub txtDescExterno_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.txtCantidad.SetFocus
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
If KeyAscii = vbKeyReturn Then ConfirmaCambios
End Sub

Private Sub ConfirmaCambios()

    If Len(Trim(Me.txtDescExterno.Text)) = 0 Then
        MsgBox "Debe ingresar la Descripción.", vbCritical, TituloSistema
        Me.txtDescExterno.SetFocus
    ElseIf Len(Trim(Me.txtCantidad.Text)) = 0 Then
        MsgBox "Debe ingresar la Cantidad.", vbCritical, TituloSistema
        Me.txtCantidad.SetFocus
    ElseIf val(Me.txtCantidad.Text) = 0 Or Not IsNumeric(Me.txtCantidad.Text) Then
        MsgBox "La Cantidad proporcionada es incorrecta.", vbCritical, TituloSistema
        Me.txtCantidad.SetFocus
        Me.txtCantidad.SelStart = 0
        Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
    ElseIf Len(Trim(Me.txtPrecio.Text)) = 0 Then
        MsgBox "Debe ingresar el Precio.", vbCritical, TituloSistema
        Me.txtPrecio.SetFocus
    ElseIf val(Me.txtPrecio.Text) = 0 Or Not IsNumeric(Me.txtPrecio.Text) Then
        MsgBox "El precio proporcinado es incorrecto.", vbCritical, TituloSistema
        Me.txtPrecio.SetFocus
        Me.txtPrecio.SelStart = 0
        Me.txtPrecio.SelLength = Len(Me.txtPrecio.Text)
    Else
        With frmContratosExternos.lvExternos.SelectedItem
            .Text = txtDescExterno.Text
            .SubItems(1) = Me.txtCantidad.Text
            .SubItems(2) = Me.txtPrecio.Text
            .SubItems(3) = val(.SubItems(1)) * val(.SubItems(2))
        End With
        Unload Me
    End If

End Sub
