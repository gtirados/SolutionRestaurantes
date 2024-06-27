VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConsultaAtenciones1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Consumos"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11355
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
   ScaleHeight     =   5895
   ScaleWidth      =   11355
   Begin VB.CommandButton cmdAnular 
      Caption         =   "&Anular"
      Height          =   480
      Left            =   9000
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   480
      Left            =   7560
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvDatos 
      Height          =   4695
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "C&onsultar"
      Height          =   480
      Left            =   6120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   300
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      Format          =   160563201
      CurrentDate     =   41499
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   300
      Left            =   4200
      TabIndex        =   3
      Top             =   300
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      Format          =   160563201
      CurrentDate     =   41499
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HASTA:"
      Height          =   195
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESDE:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   675
   End
End
Attribute VB_Name = "frmConsultaAtenciones1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAnular_Click()
If Me.lvDatos.ListItems.count = 0 Then Exit Sub
    On Error GoTo Anular

    Pub_ConnAdo.BeginTrans
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_ATENCION_ANULAR"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDATENCION", adBigInt, adParamInput, , Me.lvDatos.SelectedItem.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fecha", adDBTimeStamp, adParamInput, , CDate(Me.lvDatos.SelectedItem.SubItems(1)))
    oCmdEjec.Execute

    'OBTENIENDO LOS DETALLES DE LA ATENCION
    Dim ORSa As ADODB.Recordset

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_ATENCION_DETALLE_PRINTLIST"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDATENCION", adBigInt, adParamInput, , Me.lvDatos.SelectedItem.Text)

    Set ORSa = oCmdEjec.Execute

    Do While Not ORSa.EOF
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPACTUALIZASTOCK_ATENCION"
    
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Usuario", adVarChar, adParamInput, 20, LK_CODUSU)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodArt", adDouble, adParamInput, , ORSa!IDE)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@cp", adInteger, adParamInput, , ORSa!cant)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@nro", adInteger, adParamInput, , Me.lvDatos.SelectedItem.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@tipo", adBoolean, adParamInput, , 0) '0 cuando es extorno
        oCmdEjec.Execute
    
        ORSa.MoveNext
    Loop

    Pub_ConnAdo.CommitTrans

MsgBox "Datos Almacenados Correctamente.", vbInformation, Pub_Titulo
Me.lvDatos.ListItems.Remove Me.lvDatos.SelectedItem.Index
    Exit Sub

Anular:
    Pub_ConnAdo.RollbackTrans
    MsgBox Err.Description, vbCritical, Pub_Titulo
    
End Sub

Private Sub cmdConsultar_Click()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_ATENCION_FILL"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@F1", adDBTimeStamp, adParamInput, , Me.dtpDesde.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@F2", adDBTimeStamp, adParamInput, , Me.dtpHasta.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    
    Dim ORSd As ADODB.Recordset
    
    Set ORSd = oCmdEjec.Execute
    
    Me.lvDatos.ListItems.Clear
    
    Dim itemX As Object

    Do While Not ORSd.EOF
        Set itemX = Me.lvDatos.ListItems.Add(, , ORSd!ideatencion)
        itemX.SubItems(1) = ORSd!fecha
        itemX.SubItems(2) = ORSd!empleado
        itemX.SubItems(3) = ORSd!turno
            ORSd.MoveNext
    Loop
                
End Sub

Private Sub CmdImprimir_Click()

If Me.lvDatos.ListItems.count = 0 Then Exit Sub

    Dim objCrystal  As New CRAXDRT.APPLICATION

    Dim RutaReporte As String

    RutaReporte = "c:\Admin\Nordi\Atencion.rpt"

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
  
            
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.CommandText = "SP_ATENCION_PRINT1"
    'oCmdEjec.CommandText = "SpPrintComanda"

    Dim rsd As ADODB.Recordset
            
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDATENCION", adBigInt, adParamInput, , Me.lvDatos.SelectedItem.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fecha", adDBTimeStamp, adParamInput, , CDate(Me.lvDatos.SelectedItem.SubItems(1)))

    Set rsd = oCmdEjec.Execute
            
    VReporte.DataBase.SetDataSource rsd, 3, 1
    'frmprint.CRViewer1.ReportSource = VReporte
    'frmprint.CRViewer1.ViewReport
    VReporte.PrintOut False, 1, , 1, 1
    Set objCrystal = Nothing
    Set VReporte = Nothing

End Sub

Private Sub dtpDesde_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.dtpHasta.SetFocus
End Sub

Private Sub dtpHasta_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdConsultar_Click
End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
ConfigurarLV
Me.dtpHasta.Value = Date
Me.dtpDesde.Value = DateAdd("d", -1, Date)

End Sub

Private Sub ConfigurarLV()

    With Me.lvDatos

        .ColumnHeaders.Add , , "Nro Atención", 600
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "Empleado", 6000
        .ColumnHeaders.Add , , "Turno", 1200
        
        .Gridlines = True
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .View = lvwReport
        .MultiSelect = False

    End With

End Sub
