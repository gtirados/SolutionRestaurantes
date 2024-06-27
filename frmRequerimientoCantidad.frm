VERSION 5.00
Begin VB.Form frmRequerimientoCantidad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingrese Cantidad"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7410
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
   ScaleHeight     =   1035
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   480
      Left            =   5880
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   480
      Left            =   4440
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtCantidad 
      Height          =   405
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7665
   End
End
Attribute VB_Name = "frmRequerimientoCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gAcepta As Boolean
Public gCant As Double

Private Sub cmdAceptar_Click()

    If Not IsNumeric(Me.txtCantidad.Text) Then
        MsgBox "Cantidad incorrecta.", vbCritical, Pub_Titulo

        Exit Sub

    End If

    gAcepta = True
    gCant = Me.txtCantidad.Text
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Me.txtCantidad.SelStart = 0
Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdAceptar_Click
End Sub
