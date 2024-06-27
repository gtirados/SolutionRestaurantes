VERSION 5.00
Begin VB.Form frmAsigCantFac 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingrese Cantidad"
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3105
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
   ScaleHeight     =   600
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CANTIDAD."
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1005
   End
End
Attribute VB_Name = "frmAsigCantFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vAcepta As Boolean
Public vCANTIDAD As Double

Private Sub Form_Activate()
Me.txtCantidad.SelStart = 0
Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
Me.txtCantidad.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
vAcepta = False
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If val(Me.txtCantidad.Text) <= 0 Then
            MsgBox "Debe ingresar un Valor Valido", vbCritical, "Error"
            Exit Sub

        End If

        vAcepta = True
        vCANTIDAD = Me.txtCantidad.Text
        Unload Me
    End If

End Sub
