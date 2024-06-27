VERSION 5.00
Begin VB.Form frmsubfamiliacaracteristica_edit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modificar Caracteristica"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
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
   ScaleHeight     =   1215
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   990
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   2400
      TabIndex        =   2
      Top             =   720
      Width           =   990
   End
   Begin VB.TextBox txtCaracteristica 
      Height          =   285
      Left            =   120
      MaxLength       =   30
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caracteristica:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1260
   End
End
Attribute VB_Name = "frmsubfamiliacaracteristica_edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gAcepta As Boolean
Public gCarac As String

Private Sub cmdAceptar_Click()
gCarac = Me.txtCaracteristica.Text
gAcepta = True
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
gAcepta = False
End Sub
