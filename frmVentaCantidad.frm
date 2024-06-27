VERSION 5.00
Begin VB.Form frmVentaCantidad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingrese cantidad"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3045
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
   ScaleHeight     =   1425
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   600
      Left            =   1560
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   600
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmVentaCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vAcepta As Boolean
Public gCantidad As Double

Private Sub cmdAceptar_Click()

    If Len(Trim(Me.txtCantidad.Text)) = 0 Then
        MsgBox "Debe ingresar la cantidad.", vbInformation, Pub_Titulo
        Me.txtCantidad.SetFocus
    ElseIf Not IsNumeric(Me.txtCantidad.Text) Then
        MsgBox "Cantidad incorrecta.", vbCritical, Pub_Titulo
        Me.txtCantidad.SetFocus
    Else
        vAcepta = True
        gCantidad = Me.txtCantidad.Text
        Unload Me
    End If

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Me.txtCantidad.SelStart = 0
Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdCancelar_Click
End Sub

Private Sub Form_Load()
vAcepta = False

End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
cmdAceptar_Click
End If
End Sub
