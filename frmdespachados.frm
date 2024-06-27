VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmdespachados 
   Caption         =   "Items Despachados"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13185
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   13185
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdprint 
      Caption         =   "Imprimir"
      Height          =   615
      Left            =   11280
      TabIndex        =   14
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame FraPorDespachar 
      Height          =   6135
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   12975
      Begin MSComctlLib.ListView lvDespachados 
         Height          =   5535
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9763
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ITEMS DESPACHADOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   3405
      End
   End
   Begin VB.Frame FraEjecucion 
      Caption         =   "Opción de Ejecución"
      Height          =   975
      Left            =   6720
      TabIndex        =   7
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "&Actualizar"
         Height          =   600
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FraFILTRO 
      Caption         =   "Filtro"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin MSDataListLib.DataCombo DatFamilia 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   315
         Left            =   3240
         TabIndex        =   2
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   79626241
         CurrentDate     =   41400
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   315
         Left            =   4920
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   79626241
         CurrentDate     =   41400
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   4920
         TabIndex        =   6
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         Height          =   195
         Left            =   3240
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Familias"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.Label lblPromedioDespacho 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Promedio Demora"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9000
      TabIndex        =   13
      Top             =   120
      Width           =   1875
   End
   Begin VB.Label lblPromedio 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   9000
      TabIndex        =   12
      Top             =   480
      Width           =   1530
   End
End
Attribute VB_Name = "frmdespachados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub RealizarBusqueda()
    Me.lvDespachados.ListItems.Clear
    'Me.lvTotales.ListItems.Clear
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPITEMSDESPACHADOS"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA1", adDBTimeStamp, adParamInput, , Me.dtpDesde.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA2", adDBTimeStamp, adParamInput, , Me.dtpHasta.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDFAMILIA", adBigInt, adParamInput, , Me.DatFamilia.BoundText)

    Dim ORSDatos As ADODB.Recordset

    Set ORSDatos = oCmdEjec.Execute

    Dim oRSr  As ADODB.Recordset

    Dim ITEMd As Object
Dim xTiempo As String
xTiempo = ""
    Do While Not ORSDatos.EOF
        Set ITEMd = Me.lvDespachados.ListItems.Add(, , ORSDatos!producto)
        ITEMd.SubItems(1) = ORSDatos!Cantidad
        ITEMd.SubItems(2) = ORSDatos!mesa
        ITEMd.SubItems(3) = ORSDatos!mozo
        ITEMd.SubItems(4) = ORSDatos!hORA
        ITEMd.SubItems(5) = ORSDatos!tiempo
        ITEMd.SubItems(6) = ORSDatos!DETALLE
        ITEMd.SubItems(7) = ORSDatos!NUMERO
        ITEMd.SubItems(8) = ORSDatos!SEC
        ITEMd.SubItems(9) = ORSDatos!serie
        ITEMd.SubItems(10) = ORSDatos!CodArt
        ITEMd.SubItems(11) = ORSDatos!tiempo2

If Len(Trim(xTiempo)) <> 0 Then
        xTiempo = Format(TimeValue(ORSDatos!tiempo2) + TimeValue(xTiempo), "hh:mm:ss")
        Else
        xTiempo = Format(TimeValue(ORSDatos!tiempo2), "hh:mm:ss")
        End If
        ORSDatos.MoveNext
    Loop

Me.lblPromedio.Caption = Format(TimeValue(xTiempo) / Me.lvDespachados.ListItems.count, "hh:mm:ss")
    Set oRSr = ORSDatos.NextRecordset

    Dim i As Integer

    For i = 1 To Me.lvDespachados.ListItems.count
        Me.lvDespachados.ListItems(i).Selected = False
    
    Next

   

End Sub

Private Sub cmdActualizar_Click()

    If Me.DatFamilia.BoundText = "" Then
        MsgBox "Debe elegir la familia.", vbCritical, Pub_Titulo

        Exit Sub

    End If

    RealizarBusqueda
End Sub

Private Sub cmdprint_Click()
Dim rsd As ADODB.Recordset

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPITEMSDESPACHADOS"
    Me.lvDespachados.ListItems.Clear

    With oCmdEjec
        .Parameters.Append .CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        .Parameters.Append .CreateParameter("@DESDE", adDBTimeStamp, adParamInput, , Me.dtpDesde.Value)
        .Parameters.Append .CreateParameter("@HASTA", adDBTimeStamp, adParamInput, , Me.dtpHasta.Value)
      '  .Parameters.Append .CreateParameter("@CODUSUARIO", adVarChar, adParamInput, 10, Me.DatUsuario.BoundText)

        If Me.DatFamilia.BoundText <> "-1" Then
            .Parameters.Append .CreateParameter("@IDFAMILIA", adVarChar, adParamInput, 10, Me.DatFamilia.BoundText)
        End If

    End With

    Set rsd = oCmdEjec.Execute
    
  
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal  As New CRAXDRT.APPLICATION

    Dim RutaReporte As String

    RutaReporte = "C:\Admin\Nordi\PlatosDespachados.rpt"
    

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
 
 

    VReporte.DataBase.SetDataSource rsd, , 1

 
    frmVisor.cr.ReportSource = VReporte
    frmVisor.cr.ViewReport
    frmVisor.Show
    Set objCrystal = Nothing
    Set VReporte = Nothing
End Sub

Private Sub Form_Load()
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SpListarFamilias2"
 oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    
    
    Dim ORSf As ADODB.Recordset
    Set ORSf = oCmdEjec.Execute
    
    Set Me.DatFamilia.RowSource = ORSf
    Me.DatFamilia.BoundColumn = ORSf.Fields(0).Name
    Me.DatFamilia.ListField = ORSf.Fields(1).Name
    Me.DatFamilia.BoundText = -1
    
    Me.dtpDesde.Value = LK_FECHA_DIA
    Me.dtpHasta.Value = LK_FECHA_DIA
    ConfigurarLV
End Sub

Private Sub ConfigurarLV()

    With Me.lvDespachados
        .ColumnHeaders.Add , , "PRODUCTO", 5000
        .ColumnHeaders.Add , , "CANT.", 1200
        .ColumnHeaders.Add , , "MESA", 1200
        .ColumnHeaders.Add , , "MOZO"
        .ColumnHeaders.Add , , "HORA DESP.", 2400
        .ColumnHeaders.Add , , "TIEMPO", 1640
        .ColumnHeaders.Add , , "DETALLE", 0
        .ColumnHeaders.Add , , "NRO", 0
        .ColumnHeaders.Add , , "SEC", 0
        .ColumnHeaders.Add , , "SERIE", 0
        .ColumnHeaders.Add , , "CODART", 0
        .ColumnHeaders.Add , , "DEMORA", 1700
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
    End With

   

End Sub

Private Sub Form_Resize()
If (Me.ScaleWidth - 6300) <= 0 Then Exit Sub
If (Me.ScaleHeight - 1800) <= 0 Then Exit Sub

Me.FraPorDespachar.Width = Me.ScaleWidth - 200


'listview
Me.lvDespachados.Width = Me.ScaleWidth - 500

'ALTO
Me.FraPorDespachar.Height = Me.ScaleHeight - 1200

Me.lvDespachados.Height = Me.ScaleHeight - 1800

End Sub

