VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmComandaFamilia 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Elija"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   1740
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   4080
      TabIndex        =   3
      Top             =   1200
      Width           =   990
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   4080
      TabIndex        =   2
      Top             =   480
      Width           =   990
   End
   Begin MSDataListLib.DataCombo DatFamilia 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "Familia"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Familia:"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   540
   End
End
Attribute VB_Name = "frmComandaFamilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gIDfam As Integer
Public vAcepta As Boolean

Private Sub cmdAceptar_Click()
gIDfam = Me.DatFamilia.BoundText
vAcepta = True
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
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_COMANDA_CARGARFAMILIA"

    Dim ORSd As ADODB.Recordset

    Set ORSd = oCmdEjec.Execute(, LK_CODCIA)
    Set Me.DatFamilia.RowSource = ORSd
    Me.DatFamilia.ListField = ORSd.Fields(1).Name
    Me.DatFamilia.BoundColumn = ORSd.Fields(0).Name
    Me.DatFamilia.BoundText = gIDfam
End Sub
