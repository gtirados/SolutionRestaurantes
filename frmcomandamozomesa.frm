VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmcomandamozomesa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Validar Usuario"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5205
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
   ScaleHeight     =   4260
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSDataListLib.DataCombo DatMozos 
      Height          =   315
      Left            =   840
      TabIndex        =   14
      Top             =   360
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3360
      TabIndex        =   13
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "<="
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2520
      TabIndex        =   11
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdnum 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   9
      Left            =   2520
      TabIndex        =   10
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdnum 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   8
      Left            =   1680
      TabIndex        =   9
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdnum 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   7
      Left            =   840
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdnum 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   6
      Left            =   2520
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdnum 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   5
      Left            =   1680
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdnum 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   4
      Left            =   840
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdnum 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   3
      Left            =   2520
      TabIndex        =   4
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdnum 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   2
      Left            =   1680
      TabIndex        =   3
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdnum 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdnum 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtClave 
      Alignment       =   1  'Right Justify
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3360
      TabIndex        =   12
      Top             =   2520
      Width           =   1335
   End
End
Attribute VB_Name = "frmcomandamozomesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gEntro As Boolean
Public gModificaMozo As Boolean
Public gCodMesa As String
Public gCodMozo As String
Public gMozo As String

Private Sub cmdAceptar_Click()

    Dim ORSd As ADODB.Recordset

    If gModificaMozo Then

        'aqui verifica clave
       
            LimpiaParametros oCmdEjec

            oCmdEjec.CommandText = "SP_COMANDA_DATOSMOZO1"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)

            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MOZO", adDouble, adParamInput, , Me.DatMozos.BoundText)

            Set ORSd = oCmdEjec.Execute
     
            If ORSd!Clave = txtClave.Text Then
                gEntro = True
                gCodMozo = Me.DatMozos.BoundText
                Unload Me
            Else
                MsgBox "Datos Incorrectos", vbCritical, Pub_Titulo
                        
            End If
        
         
    Else
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPCOMANDA_EVALUA_MOZO"

        Set ORSd = oCmdEjec.Execute(, Array(LK_CODCIA, Me.DatMozos.BoundText, Me.txtClave.Text))

        If ORSd!pasa Then
            gEntro = True
            gCodMozo = Me.DatMozos.BoundText
            Unload Me
        Else
            MsgBox "Datos Incorrectos", vbCritical, Pub_Titulo
        End If
    End If

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If Len(Trim(Me.txtClave.Text)) <> 0 Then
Me.txtClave.Text = Left(Me.txtClave.Text, Len(Trim(txtClave.Text)) - 1)
End If
End Sub

Private Sub cmdNum_Click(Index As Integer)
Me.txtClave.Text = Me.txtClave.Text & Index
End Sub

Private Sub Form_Load()
    gEntro = False

    If gModificaMozo Then

        Dim orsM As ADODB.Recordset

        Set orsM = New ADODB.Recordset
        orsM.CursorType = adOpenDynamic ' setting cursor type
        orsM.Fields.Append "codmozo", adBigInt
        orsM.Fields.Append "mozo", adVarChar, 120
    
        orsM.Fields.Refresh
        orsM.Open
  
        orsM.AddNew
        orsM!CODMOZO = gCodMozo
        orsM!mozo = gMozo
        
        orsM.Update
        
        Set Me.DatMozos.RowSource = orsM
    '
        Me.DatMozos.ListField = orsM.Fields(1).Name
        Me.DatMozos.BoundColumn = orsM.Fields(0).Name
        If Not orsM.EOF Then Me.DatMozos.BoundText = orsM!CODMOZO
    End If

    '    Dim ORSm As ADODB.Recordset
    '
    '
    '
    '    LimpiaParametros oCmdEjec
    '    oCmdEjec.CommandText = "SPCARGA_MOZOSxMESAS"
    '    Set ORSm = oCmdEjec.Execute(, Array(LK_CODCIA, gCodMesa))
    '
    '    Set Me.DatMozos.RowSource = ORSm
    '
    '    Me.DatMozos.ListField = ORSm.Fields(1).Name
    '    Me.DatMozos.BoundColumn = ORSm.Fields(0).Name
    '
    '    If Not ORSm.EOF Then Me.DatMozos.BoundText = ORSm!CODMOZO
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdAceptar_Click
End Sub
