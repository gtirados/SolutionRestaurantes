VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acceso"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtClave 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   2400
      Picture         =   "frmLogin.frx":06EA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   1200
      Picture         =   "frmLogin.frx":0DD4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo dcboCia 
      Height          =   360
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcboUsuarios 
      Height          =   360
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "CLAVE:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   7
      Top             =   1147
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "LOCAL:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   6
      Top             =   660
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "USUARIO:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
On Error GoTo Accesar
'cargar las cia a las q el usuario tiene acceso
Dim oRsCias As ADODB.Recordset
Dim vclave As String
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SpCargarCias"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodUser", adVarChar, adParamInput, 10, Me.dcboUsuarios.BoundText)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Clave", adChar, adParamOutput, 10)
Set oRsCias = oCmdEjec.Execute
vclave = oCmdEjec.Parameters("@Clave").Value
oRsCias.Filter = "cia='" & Me.dcboCia.BoundText & "'"
If Not oRsCias.EOF Then
    If Trim(Me.txtClave.Text) = Trim(vclave) Then
        CodCia = Me.dcboCia.BoundText
        'cargar la fecha de la tabla pargen
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SpDevuelveFechaParGen"
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, CodCia)
        Dim orsTmp As ADODB.Recordset
        Set orsTmp = oCmdEjec.Execute
        If Not orsTmp.EOF Then FechaGen = orsTmp!FechaGen
    
        Usuario = Trim(Me.dcboUsuarios.BoundText)
        frmZonas.Show
        Unload Me
    Else
        MsgBox "La Clave proporcionada es Incorrecta.", vbCritical, NombreProyecto
        End If
    Else
    MsgBox "No Tiene permicos para dicha opción.", vbCritical, NombreProyecto
End If
Exit Sub
Accesar:
MostrarErrores Err
End Sub

Private Sub cmdCancelar_Click()
End
End Sub

Private Sub dcboCia_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.txtClave.SetFocus
End Sub

Private Sub dcboUsuarios_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.dcboCia.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then End
End Sub

Private Sub Form_Load()
Dim oRsTemp As ADODB.Recordset
EstablecerConexion
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SpListarUsuarios"
Set oRsTemp = oCmdEjec.Execute
Set Me.dcboUsuarios.RowSource = oRsTemp
Me.dcboUsuarios.ListField = oRsTemp.Fields(1).Name
Me.dcboUsuarios.BoundColumn = oRsTemp.Fields(0).Name
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SpListarParGen"
Set oRsTemp = oCmdEjec.Execute
Set Me.dcboCia.RowSource = oRsTemp
Me.dcboCia.ListField = oRsTemp.Fields(1).Name
Me.dcboCia.BoundColumn = oRsTemp.Fields(0).Name
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdAceptar_Click
End Sub
