VERSION 5.00
Begin VB.Form frmClientesUrbAdd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingrese Urbanización"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4110
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
   ScaleHeight     =   1290
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   3000
      TabIndex        =   3
      Top             =   840
      Width           =   990
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   990
   End
   Begin VB.TextBox txtUrb 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "frmClientesUrbAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gAcepta As Boolean
Public gNombre As String
Public gIde As Integer

Private Sub cmdAceptar_Click()

    If Len(Trim(Me.txtUrb.Text)) = 0 Then
        MsgBox "Debe ingresar el Nombre.", vbCritical, Pub_Titulo
        Me.txtUrb.SetFocus
    Else
    
        On Error GoTo xSAVE

        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_URBANIZACION_REGISTRAR"
    
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NOMBRE", adVarChar, adParamInput, 100, Me.txtUrb.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDURB", adBigInt, adParamOutput, , 0)
        oCmdEjec.Execute
        gIde = oCmdEjec.Parameters("@IDURB").Value
        gNombre = Me.txtUrb.Text
        gAcepta = True
        Unload Me

        Exit Sub

xSAVE:
        MsgBox Err.Description, vbCritical, Pub_Titulo
    End If

End Sub

Private Sub cmdCancelar_Click()
gAcepta = False
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdCancelar_Click
End Sub

Private Sub Form_Load()
gAcepta = False
End Sub

Private Sub txtUrb_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdAceptar_Click
End Sub
