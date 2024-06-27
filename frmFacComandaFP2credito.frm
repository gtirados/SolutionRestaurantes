VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFacComandaFP2credito 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Días de Crédito"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4260
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   3136
      TabIndex        =   2
      Top             =   360
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtDiasCredito"
      BuddyDispid     =   196610
      OrigLeft        =   2280
      OrigTop         =   1200
      OrigRight       =   2520
      OrigBottom      =   1935
      Max             =   60
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtDiasCredito 
      Height          =   375
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblFechaVenc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Días de Crédito:"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1410
   End
End
Attribute VB_Name = "frmFacComandaFP2credito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gDiasCredito As Integer
Public gAcepta As Boolean

Private Sub cmdAceptar_Click()
If Len(Trim(Me.txtDiasCredito.Text)) = 0 Then
    MsgBox "Debe ingresar los días de crédito.", vbInformation, Pub_Titulo
    Exit Sub
End If
If Not IsNumeric(Me.txtDiasCredito.Text) Then
    MsgBox "Días de crédito incorrecto.", vbCritical, Pub_Titulo
    Me.txtDiasCredito.SetFocus
    Exit Sub
End If
If val(Me.txtDiasCredito.Text) = 0 Then
     MsgBox "Días de crédito incorrecto.", vbCritical, Pub_Titulo
    Me.txtDiasCredito.SetFocus
    Exit Sub
End If
gAcepta = True
gDiasCredito = Me.txtDiasCredito.Text
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
Me.txtDiasCredito.Text = 1
gAcepta = False
End Sub

Private Sub txtDiasCredito_Change()
If IsNumeric(Me.txtDiasCredito.Text) Then
Me.lblFechaVenc.Caption = DateAdd("d", Me.txtDiasCredito.Text, LK_FECHA_DIA)
End If
End Sub

Private Sub txtDiasCredito_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
If KeyAscii = vbKeyReturn Then cmdAceptar_Click
End Sub

Private Sub txtDiasCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' click derecho
If Button = 2 Then MsgBox "Acción no permitida", vbCritical, Pub_Titulo

End Sub
