VERSION 5.00
Begin VB.Form frmComandaClave 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingrese Clave"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3900
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   1860
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   990
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   990
   End
   Begin VB.TextBox txtClave 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblMozo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   2265
   End
End
Attribute VB_Name = "frmComandaclave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vAcepta As Boolean
Public vclave As String

Private Sub cmdAceptar_Click()
vAcepta = True
vclave = Me.txtClave.Text
Unload Me
End Sub

Private Sub cmdCancelar_Click()
vAcepta = False
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdCancelar_Click
End Sub

Private Sub Form_Load()
vAcepta = False
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdAceptar_Click
End Sub
