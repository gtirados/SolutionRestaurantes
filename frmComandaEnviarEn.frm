VERSION 5.00
Begin VB.Form frmComandaEnviarEn 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Enviar En"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4080
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
   ScaleHeight     =   2190
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   600
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   600
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox ComTiempo 
      Height          =   315
      ItemData        =   "frmComandaEnviarEn.frx":0000
      Left            =   960
      List            =   "frmComandaEnviarEn.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Elegir tiempo para enviar a Cocina"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2985
   End
End
Attribute VB_Name = "frmComandaEnviarEn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gAcepta As Boolean
Public gItem As Integer

Private Sub cmdAceptar_Click()
    gAcepta = True
    If IsNumeric(Split(Me.ComTiempo.Text, " ", , vbBinaryCompare)(0)) Then
        gItem = Split(Me.ComTiempo.Text, " ", , vbBinaryCompare)(0)
    Else
        gItem = 0
    End If

    Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    gAcepta = False
    gItem = 0
    Me.ComTiempo.AddItem "AHORA"
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_TIEMPOS_LISTAR"

    Dim orsT As ADODB.Recordset

    Set orsT = oCmdEjec.Execute(, Array(LK_CODCIA, 1))

    Do While Not orsT.EOF
        Me.ComTiempo.AddItem orsT!DENOMINACION
        orsT.MoveNext
    Loop

End Sub
