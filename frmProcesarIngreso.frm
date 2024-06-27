VERSION 5.00
Begin VB.Form frmProcesarIngreso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesar Ingreso"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13005
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   13005
   Begin VB.Timer tmrMensaje 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5880
      Top             =   3000
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "=>"
      Enabled         =   0   'False
      Height          =   4200
      Left            =   12240
      TabIndex        =   8
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "<="
      Enabled         =   0   'False
      Height          =   4200
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtDni 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   960
      MaxLength       =   8
      TabIndex        =   1
      Top             =   300
      Width           =   2175
   End
   Begin VB.Frame fraOpcion 
      Height          =   4335
      Index           =   0
      Left            =   10440
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "ELEGIR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   0
         Left            =   4200
         TabIndex        =   6
         Top             =   3600
         Width           =   1230
      End
      Begin VB.Label lblMenu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5235
      End
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label lblCliente 
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
      Height          =   645
      Left            =   4320
      TabIndex        =   3
      Top             =   1080
      Width           =   8325
   End
   Begin VB.Label lblTitulo 
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
      Height          =   645
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   5445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI:"
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
      Left            =   360
      TabIndex        =   0
      Top             =   405
      Width           =   570
   End
End
Attribute VB_Name = "frmProcesarIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vPaginaActual As Integer
Private vTotalPaginas As Integer
Private vTotalOpciones As Integer
Private vOpcionElegida As Integer
Private C As Integer

Private Sub cmdNext_Click()

    Dim ini, fin, f As Integer

    If vPaginaActual = 1 Then
        ini = 1
        fin = ini * 2
    ElseIf vPaginaActual = vTotalPaginas Then

        Exit Sub

    Else
        ini = (2 * vPaginaActual) - 1
        fin = 2 * vPaginaActual
    End If

    For f = ini To fin
        Me.fraOpcion(f).Visible = False
    Next

    If vPaginaActual < vTotalPaginas Then
        vPaginaActual = vPaginaActual + 1

        If vPaginaActual = vTotalPaginas Then: Me.cmdNext.Enabled = False
        Me.cmdPreview.Enabled = True
    End If

    ''If vPagActFam = 1 Then
    ''    ini = 1
    ''    fin = ini * 14
    ''ElseIf vPagActFam = vPagTotFam Then
    ''    Exit Sub
    ''Else
    ''    ini = (14 * vPagActFam) - 13
    ''    fin = 14 * vPagActFam
    ''End If
    ''
    ''For f = ini To fin
    ''    Me.cmdFam(f).Visible = False
    ''Next
    ''If vPagActFam < vPagTotFam Then
    ''vPagActFam = vPagActFam + 1
    ''    If vPagActFam = vPagTotFam Then: Me.cmdFamSig.Enabled = False
    ''
    ''    Me.cmdFamAnt.Enabled = True
    ''End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)


Dim C As Integer
For C = 1 To vTotalOpciones
    Me.fraOpcion(C).BackColor = -2147483633
Next
Me.fraOpcion(Index).BackColor = vbRed
Me.txtDni.SetFocus
Me.txtDni.SelStart = 0
Me.txtDni.SelLength = Len(Me.txtDni.Text)
vOpcionElegida = Me.cmdOpcion(Index).Tag
End Sub

Private Sub cmdPreview_Click()

    Dim ini, fin, f As Integer

    If vPaginaActual = 2 Then
        ini = 1
        fin = ini * 2
    ElseIf vPaginaActual = 1 Then

        Exit Sub

    Else
        FF = vPaginaActual - 1
        ini = (2 * FF) - 1
        fin = 2 * FF
    End If

    For f = ini To fin
        Me.fraOpcion(f).Visible = True
    Next

    If vPaginaActual > 1 Then
        vPaginaActual = vPaginaActual - 1

        If vPaginaActual = 1 Then: Me.cmdPreview.Enabled = False
    
        Me.cmdNext.Enabled = True
    End If

End Sub





Private Sub Form_Load()
CargarOpciones
CentrarFormulario MDIForm1, Me
C = 1
End Sub

Private Sub tmrMensaje_Timer()

If C = 3 Then
    Me.lblMensaje.Caption = ""
    Me.lblMensaje.ForeColor = vbBlack
    Me.tmrMensaje.Enabled = False
Else
C = C + 1
End If
End Sub

Private Sub txtDni_Change()
If Len(Trim(Me.txtDni.Text)) = 8 Then
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_ATENCION_CLIENTE_DATOS"

        Dim orsC As ADODB.Recordset

        Set orsC = oCmdEjec.Execute(, Array(LK_CODCIA, Me.txtDni.Text))

        If Not orsC.EOF Then
            If orsC.RecordCount = 0 Then
                Me.lblMensaje.Caption = "DNI NO EXISTE."
                Me.lblMensaje.ForeColor = vbRed
                Me.txtDni.SelStart = 0
                Me.txtDni.SelLength = Len(Me.txtDni.Text)
                Me.txtDni.SetFocus
Me.tmrMensaje.Enabled = True
                Exit Sub

            Else
            Me.lblMensaje.Caption = ""
                Me.lblCliente.Caption = orsC!EMP
                Me.lblCliente.Tag = orsC!cod
            End If

        Else
            Me.lblMensaje.Caption = "DNI NO EXISTE."
            Me.lblMensaje.ForeColor = vbRed
            Me.tmrMensaje.Enabled = True
            Me.txtDni.SelStart = 0
            Me.txtDni.SelLength = Len(Me.txtDni.Text)
            Me.txtDni.SetFocus

            Exit Sub

        End If
        
      

       
        Else
        Me.lblCliente.Tag = ""
        Me.lblCliente.Caption = ""
        
End If

End Sub

Private Sub txtDni_GotFocus()
Me.txtDni.BackColor = vbYellow
End Sub

Private Sub txtDni_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
    
        
        
        If Len(Trim(Me.lblTitulo.Caption)) = 0 Then
            Me.lblMensaje.Caption = "No hay ningun turno vigente."
            Me.lblMensaje.ForeColor = vbRed
            Me.tmrMensaje.Enabled = True
            Me.txtDni.SetFocus
            Me.txtDni.SelStart = 0
            Me.lblCliente.Caption = ""
            Me.txtDni.SelLength = Len(Me.txtDni.Text)
            Exit Sub
        End If
        
        
          If Len(Trim(Me.txtDni.Text)) < 8 Then
            Me.lblMensaje.Caption = "Debe ingresar 8 digitos para el DNI."
            Me.lblMensaje.ForeColor = vbRed
            Me.tmrMensaje.Enabled = True
            Me.txtDni.SetFocus
            Me.txtDni.SelStart = 0
            Me.lblCliente.Caption = ""
            Me.txtDni.SelLength = Len(Me.txtDni.Text)

            Exit Sub

        End If
        
     If vOpcionElegida = 0 Then
            Me.lblMensaje.Caption = "Debe elegir una Opcion"
            Me.lblMensaje.ForeColor = vbRed
            Me.tmrMensaje.Enabled = True
          Me.txtDni.SetFocus
            Me.txtDni.SelStart = 0
            Me.lblCliente.Caption = ""
            Me.txtDni.SelLength = Len(Me.txtDni.Text)
            Exit Sub

        End If
        
        If Len(Trim(Me.lblCliente.Tag)) = 0 Then
        Exit Sub
        End If
        
      'VALIDANDO EMPLEADO REPETIDO
        Dim vATENDIDO As Boolean

        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_ATENCION_VALIDA"
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDTURNO", adInteger, adParamInput, , Me.lblTitulo.Tag)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPLEADO", adBigInt, adParamInput, , Me.lblCliente.Tag)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DATO", adBoolean, adParamOutput, , vATENDIDO)
        oCmdEjec.Execute
        
        vATENDIDO = oCmdEjec.Parameters("@DATO").Value
        
        If Not vATENDIDO Then
            Grabar_Atencion
        Else
            MsgBox "Ya fue registrado.", vbCritical, Pub_Titulo
            Me.lblCliente.Tag = ""
            Me.lblCliente.Caption = ""
            LimpiarTemporales
            CargarOpciones
            Me.txtDni.Text = ""
            Me.txtDni.SetFocus
        End If
    End If

End Sub

Private Sub Grabar_Atencion()

    On Error GoTo SaveA

    Pub_ConnAdo.BeginTrans
    LimpiaParametros oCmdEjec

    Dim vIDe As Double

    oCmdEjec.CommandText = "SP_ATENCION_REGISTRAR"

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDTURNO", adInteger, adParamInput, , Me.lblTitulo.Tag)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDEMPLEADO", adBigInt, adParamInput, , Me.lblCliente.Tag)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDATENCION", adBigInt, adParamOutput, , vIDe)
    oCmdEjec.Execute
        
    vIDe = oCmdEjec.Parameters("@IDATENCION").Value
        
    'GRABANDO EN DETALLE DE ATENCION
        
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_ATENCION_DETALLE_REGISTRAR"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDATENCION", adBigInt, adParamInput, , vIDe)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPROGRAMACION", adBigInt, adParamInput, , vOpcionElegida)
    oCmdEjec.Execute
      
    'OBTENIENDO LOS DETALLES DE LA ATENCION
    Dim ORSa As ADODB.Recordset

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_ATENCION_DETALLE_PRINTLIST"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDATENCION", adBigInt, adParamInput, , vIDe)

    Set ORSa = oCmdEjec.Execute

    Do While Not ORSa.EOF
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPACTUALIZASTOCK_ATENCION"
    
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Usuario", adVarChar, adParamInput, 20, LK_CODUSU)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodArt", adDouble, adParamInput, , ORSa!IDE)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@cp", adInteger, adParamInput, , ORSa!cant)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@nro", adInteger, adParamInput, , vIDe)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@tipo", adBoolean, adParamInput, , 1) '0 cuando es extorno
        oCmdEjec.Execute
    
        ORSa.MoveNext
    Loop
        
       
    Print_Atencion (vIDe)
    
    Pub_ConnAdo.CommitTrans
    Me.lblMensaje.Caption = "Datos Registrados Correctamente."
    Me.lblMensaje.ForeColor = vbBlack
    Me.tmrMensaje.Enabled = True
    LimpiarTemporales
    CargarOpciones
    Me.txtDni.Text = ""
    Me.txtDni.SetFocus

    Exit Sub

SaveA:
    Pub_ConnAdo.RollbackTrans
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub Print_Atencion(VidATENCION As Integer)

    Dim objCrystal  As New CRAXDRT.APPLICATION

    Dim RutaReporte As String

    RutaReporte = "c:\Admin\Nordi\Atencion.rpt"

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
  
            
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.CommandText = "SP_ATENCION_PRINT"
    'oCmdEjec.CommandText = "SpPrintComanda"

    Dim rsd As ADODB.Recordset
            
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDATENCION", adBigInt, adParamInput, , VidATENCION)

    Set rsd = oCmdEjec.Execute
            
    VReporte.DataBase.SetDataSource rsd, 3, 1
    'frmprint.CRViewer1.ReportSource = VReporte
    'frmprint.CRViewer1.ViewReport
    VReporte.PrintOut False, 1, , 1, 1
    Set objCrystal = Nothing
    Set VReporte = Nothing
End Sub


Private Sub LimpiarTemporales()
Dim Opc As Integer

For Opc = 1 To Me.fraOpcion.count - 1
    Unload Me.cmdOpcion(Opc)
    Unload Me.lblMenu(Opc)
    Unload Me.fraOpcion(Opc)
Next
vPaginaActual = 0
vTotalOpciones = 0
vOpcionElegida = 0

End Sub

Private Sub CargarOpciones()

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPPROGRAMACION_VIGENTE1"

    Dim orsG  As ADODB.Recordset

    Dim ORSd  As ADODB.Recordset
    
    Dim pLEFT As Double

    Dim pTOP  As Double
    
    pLEFT = 960
    pTOP = 2040

    Set orsG = oCmdEjec.Execute(, LK_CODCIA)

    Dim C As Integer

    C = 1
    Set ORSd = orsG.NextRecordset
    vTotalOpciones = orsG.RecordCount
    
    Dim valor As Double

    Dim sMENU As String

    valor = orsG.RecordCount / 2

    pos = InStr(Trim(str(valor)), ".")
    pos2 = Right(Trim(str(valor)), Len(Trim(str(valor))) - pos)

    If pos <> 0 Then ent = Left(Trim(str(valor)), pos - 1)

    If ent = "" Then: ent = 0
    If val(pos2) > 0 Then: vTotalPaginas = val(ent) + 1

    If orsG.RecordCount >= 1 Then: vPaginaActual = 1
    If orsG.RecordCount > 2 Then: Me.cmdNext.Enabled = True



    If Not orsG.EOF Then



        Dim iM As Integer
        
        Me.lblTitulo.Caption = orsG!turno
        Me.lblTitulo.Tag = orsG!IDETURNO

        For iM = 1 To orsG.RecordCount
            Load Me.fraOpcion(iM)

            Me.fraOpcion(iM).Visible = True
            
            If C = 1 Then
                C = C + 1
                Me.fraOpcion(iM).Left = pLEFT
            Else
                Me.fraOpcion(iM).Left = Me.fraOpcion(iM - 1).Width + Me.cmdPreview.Width + 300
                C = 1
            End If

            Me.fraOpcion(iM).Top = pTOP

            'agregando boton
            Load Me.cmdOpcion(iM)

            Set Me.cmdOpcion(iM).Container = Me.fraOpcion(iM)
            Me.cmdOpcion(iM).Visible = True
            Me.cmdOpcion(iM).Tag = orsG!IDPROGRAMACION
            'agregando label
            Load Me.lblMenu(iM)
            Set Me.lblMenu(iM).Container = Me.fraOpcion(iM)
            Me.lblMenu(iM).Visible = True
            
            
            ORSd.Filter = "IDEPROGRAMACION=" & CStr(orsG!IDPROGRAMACION)
            
            sMENU = ""
            
            sMENU = orsG!Opcion & vbCrLf & vbCrLf
            
            Do While Not ORSd.EOF
                sMENU = sMENU + Left(Trim(str(ORSd!Cantidad)) + Space(2), 5) + Trim(ORSd!producto) + vbCrLf
                ORSd.MoveNext
            Loop
            
            Me.lblMenu(iM).Caption = sMENU
            
            orsG.MoveNext
        Next
       
       orsG.MoveFirst
       
       If orsG.RecordCount = 1 Then
vOpcionElegida = orsG!IDPROGRAMACION
Me.fraOpcion(1).BackColor = vbRed
End If

    End If
End Sub

Private Sub txtDni_LostFocus()
Me.txtDni.BackColor = vbWhite
End Sub
