VERSION 5.00
Begin VB.Form frmAgreFamSubFam 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   6120
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
   ScaleHeight     =   1650
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDesCor 
      Height          =   285
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Desccripción Corta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label lblcodigo 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   60
   End
   Begin VB.Label lblfamilia 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Desccripción:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Familia:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   525
   End
End
Attribute VB_Name = "frmAgreFamSubFam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VEsFam As Boolean
Public vCodFam As Integer
Public vCodSubFam As Integer
Public vAcepta As Boolean
Public vDescripcion As String


Private Sub cmdAceptar_Click()
On Error GoTo Graba
LimpiaParametros oCmdEjec
If VEsFam = False Then 'es una sub familia
 oCmdEjec.CommandText = "SpRegistraSubFamilia"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Descripcion", adVarChar, adParamInput, 40, Trim(Me.txtDescripcion.Text))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NomCorto", adVarChar, adParamInput, 10, Trim(Me.txtDesCor.Text))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TabcodArt", adInteger, adParamInput, , CInt(Me.lblfamilia.Tag))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TabNumTab", adInteger, adParamOutput, , 1)
    oCmdEjec.Execute
    vCodSubFam = oCmdEjec.Parameters("@TabNumTab").Value
    vDescripcion = Me.txtDescripcion.Text
    vAcepta = True
    Unload Me
Else
    oCmdEjec.CommandText = "SpRegistraFamilia"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Descripcion", adVarChar, adParamInput, 40, Trim(Me.txtDescripcion.Text))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NomCorto", adVarChar, adParamInput, 10, Trim(Me.txtDesCor.Text))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TabNumTab", adInteger, adParamOutput, , 1)
    oCmdEjec.Execute
    vCodFam = oCmdEjec.Parameters("@TabNumTab").Value
    vDescripcion = Me.txtDescripcion.Text
    vAcepta = True
    Unload Me
End If
Exit Sub
Graba:
MsgBox Err.Description
End Sub

Private Sub cmdCancelar_Click()
Unload Me
vAcepta = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    vAcepta = False
    Unload Me
End If
End Sub

Private Sub txtDesCor_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdAceptar_Click
End If
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.txtDesCor.SetFocus
End If
End Sub
