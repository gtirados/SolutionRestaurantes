VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmConsultaPrintCuentaLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Historica de Impresiones de Cuentas"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12705
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   12705
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   480
      Left            =   8400
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   480
      Left            =   6720
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvLog 
      Height          =   5895
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSDataListLib.DataCombo DatMesa 
      Height          =   315
      Left            =   4680
      TabIndex        =   5
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "Mesa"
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   160497665
      CurrentDate     =   41754
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   160497665
      CurrentDate     =   41754
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MESA:"
      Height          =   195
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HASTA:"
      Height          =   195
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESDE:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   675
   End
End
Attribute VB_Name = "frmConsultaPrintCuentaLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBuscar_Click()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_COMANDA_PRINT_LOG_LIST"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA1", adDBTimeStamp, adParamInput, , Me.dtpDesde.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA2", adDBTimeStamp, adParamInput, , Me.dtpHasta.Value)

    If Me.DatMesa.BoundText <> "-1" Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDMESA", adVarChar, adParamInput, 10, Me.DatMesa.BoundText)
    End If
    
    Dim ORSd As ADODB.Recordset

    Set ORSd = oCmdEjec.Execute
    
    Me.lvLog.ListItems.Clear

    Do While Not ORSd.EOF
        Set itemX = Me.lvLog.ListItems.Add(, , ORSd!cod)
        itemX.SubItems(1) = ORSd!PEDIDO
        itemX.SubItems(2) = ORSd!mesa
        itemX.SubItems(3) = ORSd!mozo
        itemX.SubItems(4) = ORSd!fecha
        itemX.SubItems(5) = ORSd!hORA
        itemX.SubItems(6) = ORSd!USUARIO
        itemX.SubItems(7) = ORSd!Total
        ORSd.MoveNext
    Loop
    
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
Me.dtpHasta.Value = LK_FECHA_DIA
Me.dtpDesde.Value = DateAdd("d", -30, LK_FECHA_DIA)
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_COMANDA_PRINT_LOG_LIST_MESAS"

    Dim ORSd As ADODB.Recordset

    Set ORSd = oCmdEjec.Execute(, LK_CODCIA)
    Set Me.DatMesa.RowSource = ORSd
    Me.DatMesa.BoundColumn = "COD"
    Me.DatMesa.ListField = "DEN"
    Me.DatMesa.BoundText = "-1"

    With Me.lvLog
        .Gridlines = True
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .View = lvwReport
    
        .HideSelection = False
        .ColumnHeaders.Add , , "Codigo", 1200
        .ColumnHeaders.Add , , "Pedido", 1200
        .ColumnHeaders.Add , , "Mesa", 1000
        .ColumnHeaders.Add , , "Mozo", 1400
        .ColumnHeaders.Add , , "Fecha", 1400
        .ColumnHeaders.Add , , "Hora", 1400
        .ColumnHeaders.Add , , "Usuario", 1000
        .ColumnHeaders.Add , , "Total", 1000
        
        .MultiSelect = True
    
    End With

End Sub
