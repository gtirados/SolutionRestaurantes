VERSION 5.00
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Begin VB.Form frmMozosEstado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mozos"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10785
   BeginProperty Font 
      Name            =   "Tahoma"
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10785
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   840
      Left            =   9360
      Picture         =   "frmMozosEstado.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1335
   End
   Begin ubGridControl.ubGrid ubgDatos 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   10186
      Rows            =   1
      Cols            =   2
      Redraw          =   -1  'True
      ShowGrid        =   -1  'True
      GridSolid       =   -1  'True
      GridLineColor   =   12632256
      BackColorFixed  =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontEdit {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMozosEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
 Dim i    As Integer

            Dim vENC As Boolean

            vENC = False

            For i = 1 To Me.ubgDatos.Rows

                If Len(Trim(Me.ubgDatos.TextMatrix(i, 2))) = 0 Then
                    vENC = True

                    Exit For

                End If

            Next
            
            If vENC Then
                MsgBox "Hay Mozos que no tienen nombre.", vbInformation, Pub_Titulo

                Exit Sub

            End If

            On Error GoTo grabar

            Pub_ConnAdo.BeginTrans

            oCmdEjec.CommandText = "SP_MOZOS_UPDATE_INFO"

            For i = 1 To Me.ubgDatos.Rows
                LimpiaParametros oCmdEjec
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDMOZO", adBigInt, adParamInput, , Me.ubgDatos.TextMatrix(i, 3))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NOMBRE", adVarChar, adParamInput, 30, Me.ubgDatos.TextMatrix(i, 2))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ESTADO", adBoolean, adParamInput, , Me.ubgDatos.TextMatrix(i, 1))
                oCmdEjec.Execute
            Next

            Pub_ConnAdo.CommitTrans
            MsgBox "Información Almacenada Correctamente.", vbInformation, Pub_Titulo

            Exit Sub

grabar:
            Pub_ConnAdo.RollbackTrans
            MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub

Private Sub Form_Load()
    Me.ubgDatos.AutoSetup 0, 2, True, True, "Activo     |Mozo  |Codigo"
    Me.ubgDatos.AutoRedraw = True
   
    Me.ubgDatos.ColMask(1) = checkmark
    Me.ubgDatos.ColWidth(2) = 600
    Me.ubgDatos.ColWidth(3) = 0
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_MOZOS_LIST"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim ORSDatos As ADODB.Recordset

    Set ORSDatos = oCmdEjec.Execute(, LK_CODCIA)

    Do While Not ORSDatos.EOF
        Me.ubgDatos.AddItem (ORSDatos!ESTADO & vbTab & ORSDatos!Nombre & vbTab & ORSDatos!Codigo)
        ORSDatos.MoveNext
    Loop

End Sub
