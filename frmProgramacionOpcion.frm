VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmProgramacionOpcion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccione Turno"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4065
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
   ScaleHeight     =   2145
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   600
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   600
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo DatTurno 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Turno:"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   570
   End
End
Attribute VB_Name = "frmProgramacionOpcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gNUEVO As Boolean
Public gFECHA As Date

Private Sub cmdAceptar_Click()

    If Me.DatTurno.BoundText = -1 Then
        MsgBox "Debe seleccionar el Turno.", vbCritical, Pub_Titulo
        Me.DatTurno.SetFocus

        Exit Sub

    End If

'    LimpiaParametros oCmdEjec
'    oCmdEjec.CommandText = "SPPROGRAMACION_VALIDAREPETIDO"
'
'    Dim vVAL As Boolean
'
'    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
'    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , gFECHA)
'    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDTURNO", adBigInt, adParamInput, , Me.DatTurno.BoundText)
'    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DATO", adBoolean, adParamOutput, , vVAL)
'
'    oCmdEjec.Execute
'
'    vVAL = oCmdEjec.Parameters("@DATO").Value
'
'    If vVAL Then
'        MsgBox "REPETIDO", vbCritical, Pub_Titulo
'    Else
        frmProgramacion.VNuevo = gNUEVO
        frmProgramacion.xTurno = Me.DatTurno.BoundText
        frmProgramacion.Show vbModal
        Unload Me
'    End If

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub DatTurno_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.cmdAceptar.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SPLISTARTURNOS"

oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@LISTADO", adBoolean, adParamInput, , 1)

Dim ORSt As ADODB.Recordset
Set ORSt = oCmdEjec.Execute

Set Me.DatTurno.RowSource = ORSt
Me.DatTurno.ListField = ORSt.Fields(1).Name
Me.DatTurno.BoundColumn = ORSt.Fields(0).Name
Me.DatTurno.BoundText = -1


End Sub
