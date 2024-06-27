VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDeliveryCambiaRepartidor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Repartidor"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
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
   ScaleHeight     =   1695
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4800
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo DatRepartidor 
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      Text            =   "Repartidor"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cambiar por:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label lblRepartidor 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Repartidor Actual:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1560
   End
End
Attribute VB_Name = "frmDeliveryCambiaRepartidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gSerie As String
Public gNumero As String

Private Sub cmdAceptar_Click()

If Me.DatRepartidor.BoundText = -1 Then
MsgBox "Debe elegir el repartidor.", vbInformation, Pub_Titulo

Exit Sub
End If
    On Error GoTo grabar

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_DELIVERY_CAMBIAREPARTIDOR"
    oCmdEjec.Execute , Array(LK_CODCIA, gSerie, gNumero, Me.DatRepartidor.BoundText)
    Unload Me

    Exit Sub

grabar:
    MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub

Private Sub DatRepartidor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_DELIVERY_CARGAR_REPARTIDORES"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim oRSr As ADODB.Recordset

    Set oRSr = oCmdEjec.Execute(, LK_CODCIA)

    Set Me.DatRepartidor.RowSource = oRSr
    Me.DatRepartidor.ListField = oRSr.Fields(1).Name
    Me.DatRepartidor.BoundColumn = oRSr.Fields(0).Name
    Me.DatRepartidor.BoundText = -1
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_DELIVERY_CARGAREPARTIDOR"

    Dim orsD As ADODB.Recordset

    Set orsD = oCmdEjec.Execute(, Array(LK_CODCIA, gSerie, gNumero))

    If Not orsD.EOF Then
        Me.lblRepartidor.Caption = orsD!repartidor
    End If

End Sub
