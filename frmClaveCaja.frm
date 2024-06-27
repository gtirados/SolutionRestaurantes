VERSION 5.00
Begin VB.Form frmClaveCaja 
   BackColor       =   &H8000000C&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6345
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
   ScaleHeight     =   1575
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUsuario 
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtclave 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "INGRESE LA CLAVE:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "INGRESE EL USUARIO:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1980
   End
End
Attribute VB_Name = "frmClaveCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vAceptar As Boolean
Public vClave As String
Public vUSUARIO As String


Private Sub cmdAceptar_Click()

    If Len(Trim(Me.txtclave.Text)) = 0 Then
        MsgBox "Debe ingresar la clave", vbCritical, NombreProyecto
    Else
        vAceptar = True
        vUSUARIO = Trim(Me.txtUsuario.Text)
        vClave = Trim(Me.txtclave.Text)
        Unload Me
    End If

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
vAceptar = False
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdAceptar_Click
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
 KeyAscii = Mayusculas(KeyAscii)
End Sub
