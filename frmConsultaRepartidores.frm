VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmConsultaRepartidores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidación de Repartidores"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13725
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   13725
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Ver"
      Height          =   480
      Left            =   9360
      TabIndex        =   16
      Top             =   240
      Width           =   1335
   End
   Begin VB.CheckBox chkMarca 
      Caption         =   "Todos"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   975
   End
   Begin VB.CheckBox chkCobrado 
      Caption         =   "Incluir Todos"
      Height          =   195
      Left            =   12120
      TabIndex        =   14
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdPagar 
      Caption         =   "Pagar"
      Height          =   480
      Left            =   7920
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvData 
      Height          =   4455
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   480
      Left            =   6480
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   315
      Left            =   4800
      TabIndex        =   5
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   197853185
      CurrentDate     =   41923
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   315
      Left            =   3120
      TabIndex        =   4
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   197853185
      CurrentDate     =   41923
   End
   Begin MSDataListLib.DataCombo DatRepartidor 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "Repartidor"
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   11520
      TabIndex        =   13
      Top             =   6060
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Tarifas:"
      Height          =   195
      Left            =   10920
      TabIndex        =   12
      Top             =   5700
      Width           =   1140
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   12120
      TabIndex        =   10
      Top             =   6000
      Width           =   1515
   End
   Begin VB.Label lblTarifas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   12120
      TabIndex        =   9
      Top             =   5640
      Width           =   1515
   End
   Begin VB.Label lblItems 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   4920
      TabIndex        =   3
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      Height          =   195
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Repartidor:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmConsultaRepartidores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkMarca_Click()
For Each ditem In Me.lvData.ListItems
    ditem.Checked = Me.chkMarca.Value
Next
End Sub

Private Sub cmdBuscar_Click()

    If Me.DatRepartidor.BoundText = -1 Then
        MsgBox "Debe elegir el repartidor.", vbInformation, Pub_Titulo

        Exit Sub

    End If

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_REPARTIDORES_CONSULTA"

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESDE", adDBTimeStamp, adParamInput, , Me.dtpDesde.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@HASTA", adDBTimeStamp, adParamInput, , Me.dtpHasta.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDREPARTIDOR", adInteger, adParamInput, , Me.DatRepartidor.BoundText)
    
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TODOS", adBoolean, adParamInput, , Me.chkCobrado.Value)
    Me.lvData.ListItems.Clear

    Dim ORSl As ADODB.Recordset

    Set ORSl = oCmdEjec.Execute
    
    Dim itemX As Object

    Dim xTotal, xTarifas As Double

    xTotal = 0
    xTarifas = 0

    Do While Not ORSl.EOF
        Set itemX = Me.lvData.ListItems.Add(, , ORSl!Comanda)
        itemX.SubItems(1) = ORSl!NRODOCTO
        itemX.SubItems(2) = ORSl!fecha
        itemX.SubItems(3) = ORSl!cliente
        itemX.SubItems(4) = ORSl!direccion
        itemX.SubItems(5) = FormatCurrency(ORSl!subtotal, 2)
        itemX.SubItems(6) = FormatCurrency(ORSl!tarifa, 2)
        itemX.SubItems(7) = FormatCurrency(ORSl!dscto, 2)
        itemX.SubItems(8) = FormatCurrency(ORSl!Total, 2)
        itemX.SubItems(9) = ORSl!NumSer
        itemX.SubItems(10) = ORSl!NumFac
        xTotal = xTotal + ORSl!Total
        xTarifas = xTarifas + ORSl!tarifa
        ORSl.MoveNext
    
    Loop

    Me.lblItems.Caption = Me.lvData.ListItems.count & " Pedidos"
    Me.lblTarifas.Caption = FormatCurrency(xTarifas, 2)
    Me.lblTotal.Caption = FormatCurrency(xTotal, 2)
            
End Sub

Private Sub cmdPagar_Click()

    If Me.lvData.ListItems.count = 0 Then Exit Sub
   

    Dim oMSN  As String

    Dim xITEM As Object
    Dim xMar As Boolean
    xMar = False
    
    For Each xITEM In Me.lvData.ListItems
        If xITEM.Checked Then
            xMar = True
            Exit For
        End If
    Next
    
    If Not xMar Then
        MsgBox "Debe marcar algun documento.", vbInformation, Pub_Titulo
        Exit Sub
    End If

 If MsgBox("¿Desea continuar con la operación?.", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
    LimpiaParametros oCmdEjec
    
    On Error GoTo xPaga

    Pub_ConnAdo.BeginTrans
    oMSN = ""
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.CommandText = "SP_DELIVERY_PAGAR"
    
    For Each xITEM In Me.lvData.ListItems

        If xITEM.Checked Then
            LimpiaParametros oCmdEjec
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIECOMANDA", adVarChar, adParamInput, 3, xITEM.SubItems(9))
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMEROCOMANDA", adBigInt, adParamInput, , xITEM.SubItems(10))
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@exito", adVarChar, adParamOutput, 300, oMSN)
            oCmdEjec.Execute
        End If

    Next

    oMSN = oCmdEjec.Parameters("@exito").Value

    If Len(Trim(oMSN)) <> 0 Then
        Pub_ConnAdo.RollbackTrans
        MsgBox oMSN, vbCritical, Pub_Titulo
    Else
        Pub_ConnAdo.CommitTrans
        MsgBox "Datos Almacenados Correctamente.", vbInformation, Pub_Titulo
        
    End If

    Exit Sub

xPaga:
    Pub_ConnAdo.RollbackTrans
    MsgBox Err.Description, vbCritical, Pub_Titulo
  
End Sub

Private Sub cmdPrint_Click()
  Dim rsd As ADODB.Recordset

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_REPARTIDORES_CONSULTA"

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESDE", adDBTimeStamp, adParamInput, , Me.dtpDesde.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@HASTA", adDBTimeStamp, adParamInput, , Me.dtpHasta.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDREPARTIDOR", adInteger, adParamInput, , Me.DatRepartidor.BoundText)
    
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TODOS", adBoolean, adParamInput, , Me.chkCobrado.Value)
    

    Set rsd = oCmdEjec.Execute
    
  
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal  As New CRAXDRT.APPLICATION

    Dim RutaReporte As String

    RutaReporte = "C:\Admin\Nordi\ConsultaRepartos.rpt"
    

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
 
 

    VReporte.DataBase.SetDataSource rsd, , 1

 
    frmVisor.cr.ReportSource = VReporte
    frmVisor.cr.ViewReport
    frmVisor.Show
    Set objCrystal = Nothing
    Set VReporte = Nothing
End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_DELIVERY_CARGAR_REPARTIDORES"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim oRSr As ADODB.Recordset

    Set oRSr = oCmdEjec.Execute(, LK_CODCIA)

    Set Me.DatRepartidor.RowSource = oRSr
    Me.DatRepartidor.ListField = oRSr.Fields(1).Name
    Me.DatRepartidor.BoundColumn = oRSr.Fields(0).Name
    Me.DatRepartidor.BoundText = -1
    
    ConfiguraLV
    Me.dtpHasta.Value = LK_FECHA_DIA
    Me.dtpDesde.Value = LK_FECHA_DIA
End Sub


Private Sub ConfiguraLV()
With Me.lvData
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Comanda", 1000
    .ColumnHeaders.Add , , "Docto", 1200
    .ColumnHeaders.Add , , "Fecha", 1200
    .ColumnHeaders.Add , , "Cliente", 3000
    .ColumnHeaders.Add , , "Direccion", 2000
    .ColumnHeaders.Add , , "SubTotal", 1200, lvwColumnRight
    .ColumnHeaders.Add , , "Tarifa", 1200, lvwColumnRight
    .ColumnHeaders.Add , , "Dscto", 1200, lvwColumnRight
    .ColumnHeaders.Add , , "Total", 1200, lvwColumnRight
     .ColumnHeaders.Add , , "Serie", 0
    .ColumnHeaders.Add , , "Nro", 0
    .MultiSelect = False
    '.ColumnHeaders(9).Alignment = lvwColumnLeft
End With
End Sub

