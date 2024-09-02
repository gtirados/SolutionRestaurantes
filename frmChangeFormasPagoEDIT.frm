VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmChangeFormasPagoEDIT 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Forma de Pago"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8385
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
   ScaleHeight     =   3345
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   5160
      TabIndex        =   9
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdChange 
         Caption         =   "Realizar Cambio"
         Height          =   480
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   2655
      End
      Begin MSDataListLib.DataCombo DatFormaPago 
         Height          =   315
         Left            =   360
         TabIndex        =   10
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "FormaPago"
      End
      Begin VB.Label lblFBG 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   840
         TabIndex        =   16
         Top             =   240
         Width           =   60
      End
      Begin VB.Label lblNumOper 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   3240
         TabIndex        =   13
         Top             =   360
         Width           =   60
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CAMBIAR POR:"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.Label lblImporte 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Importe:"
         Height          =   195
         Left            =   1440
         TabIndex        =   14
         Top             =   2040
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label lblFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FORMA DE PAGO:"
         Height          =   195
         Left            =   2520
         TabIndex        =   6
         Top             =   1080
         Width           =   1530
      End
      Begin VB.Label lblFormaPago 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label lblNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NÚMERO:"
         Height          =   195
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SERIE:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmChangeFormasPagoEDIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gAcepta As Boolean

Private Sub cmdChange_Click()
If Me.DatFormaPago.BoundText = "" Then
    MsgBox "Debe elegir la forma de pago.", vbInformation, Pub_Titulo
    Me.DatFormaPago.SetFocus
    Exit Sub
End If
gAcepta = False
On Error GoTo cChange

LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "[dbo].[USP_DOCTOVENTAS_CAMBIARFORMAPAGO]"
oCmdEjec.Prepared = True
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adVarChar, adParamInput, 3, Trim(Me.lblSerie.Caption))
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMERO", adInteger, adParamInput, , Me.lblNumero.Caption)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Trim(Me.lblFBG.Caption))
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , Me.lblFecha.Caption)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMOPER", adInteger, adParamInput, , Me.lblNumOper.Caption)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDFORMAPAGO", adInteger, adParamInput, , Me.DatFormaPago.BoundText)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FORMAPAGO", adVarChar, adParamInput, 30, Me.DatFormaPago.Text)
oCmdEjec.Execute
MsgBox "Cambio realizado correctamente.", vbInformation, Pub_Titulo
gAcepta = True
Unload Me
Exit Sub
cChange:
MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
Me.lblNumOper.Visible = False
Me.lblFBG.Visible = False
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SPCARGARFORMASPAGO"
Set oRSmain = oCmdEjec.Execute(, 2401)

Set Me.DatFormaPago.RowSource = oRSmain
Me.DatFormaPago.ListField = oRSmain.Fields(1).Name
Me.DatFormaPago.BoundColumn = oRSmain.Fields(0).Name
End Sub

