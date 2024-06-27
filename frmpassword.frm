VERSION 5.00
Begin VB.Form frmpassword 
   Caption         =   "Ingreso de Contraseña"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtpassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frmpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'If UCase(Trim(txtpassword.Text)) <> UCase(Trim(vclave)) Then
If UCase(Trim(txtpassword.Text)) <> 123 Then
   MsgBox "Password Incorrecto", vbCritical + vbDefaultButton2, Pub_Titulo
   If txtpassword.Enabled Then txtpassword.SetFocus
   txtpassword.Text = ""
Else
frmpassword.Visible = False

End If
End If
End Sub
