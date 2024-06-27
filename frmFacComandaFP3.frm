VERSION 5.00
Begin VB.Form frmFacComandaFP3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importe"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
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
   ScaleHeight     =   2190
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3480
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   1920
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtReferencia 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox txtMonto 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. Referencia:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monto:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1260
      TabIndex        =   0
      Top             =   300
      Width           =   855
   End
End
Attribute VB_Name = "frmFacComandaFP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MuestraReferencia As Boolean

Public gACEPTA As Boolean
Public gMONTO As Double
Public gREF As String
Private vPUNTO As Boolean

Private Sub cmdAceptar_Click()

    If IsNumeric(Me.txtMonto.Text) Then
        gACEPTA = True
        gMONTO = Me.txtMonto.Text
        gREF = Me.txtReferencia.Text
        Unload Me
    Else
        MsgBox "Monto incorrecto.", vbCritical, Pub_Titulo
    End If

End Sub

Private Sub cmdCancelar_Click()
gACEPTA = False
Unload Me
End Sub

Private Sub Form_Activate()
Me.txtMonto.SelStart = 0
Me.txtMonto.SelLength = Len(Me.txtMonto.Text)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Me.txtReferencia.Visible = MuestraReferencia
    Me.Label2.Visible = MuestraReferencia
End Sub

Private Sub txtMonto_Change()
If InStr(Me.txtMonto.Text, ".") Then
    vPUNTO = True
Else
    vPUNTO = False
End If
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)

    If NumerosyPunto(KeyAscii) Then KeyAscii = 0
    If KeyAscii = 46 Then
        If vPUNTO Or Len(Trim(Me.txtMonto.Text)) = 0 Then
            KeyAscii = 0
        End If
    End If
    
    If KeyAscii = vbKeyReturn Then
        If Me.txtReferencia.Visible Then
            Me.txtReferencia.SetFocus
            Me.txtReferencia.SelStart = 0
            Me.txtReferencia.SelLength = Len(Me.txtReferencia.Text)
        Else
            cmdAceptar_Click
        End If
    End If
    
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
If KeyAscii = vbKeyReturn Then
    cmdAceptar_Click
End If
End Sub
