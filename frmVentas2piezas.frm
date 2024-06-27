VERSION 5.00
Begin VB.Form frmVentas2piezas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingrese Piezas"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   4470
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
   ScaleHeight     =   765
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   990
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   990
   End
   Begin VB.TextBox txtPiezas 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmVentas2piezas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gAcepta As Boolean
Public gPieza As Integer

Private Sub cmdAceptar_Click()
If Len(Trim(Me.txtPiezas.Text)) = 0 Then
    MsgBox "Debe ingresar las piezas.", vbInformation, Pub_Titulo
    Me.txtPiezas.SetFocus
ElseIf Not IsNumeric(Me.txtPiezas.Text) Then
    MsgBox "Las piezas son incorrectas", vbInformation, Pub_Titulo
    Me.txtPiezas.SetFocus
Else
    gAcepta = True
    gPieza = Trim(Me.txtPiezas.Text)
    Unload Me
End If

End Sub

Private Sub cmdCancelar_Click()
gAcepta = False
gPieza = 0
Unload Me
End Sub

Private Sub Form_Activate()
Me.txtPiezas.SelStart = 0
Me.txtPiezas.SelLength = Len(Me.txtPiezas.Text)
End Sub

Private Sub Form_Load()
gAcepta = False
gPieza = 0
Me.txtPiezas.Text = "0"
End Sub

Private Sub txtPiezas_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdAceptar_Click
End If
End Sub


