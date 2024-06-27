VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmConsultaDespachos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Despachos"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18180
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
   ScaleHeight     =   8415
   ScaleWidth      =   18180
   Begin VB.CommandButton cmdVer 
      Caption         =   "Visualizar"
      Height          =   480
      Left            =   11040
      TabIndex        =   6
      Top             =   210
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DatMesas 
      Height          =   315
      Left            =   6000
      TabIndex        =   5
      Top             =   300
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "Mostrar"
      Height          =   480
      Left            =   9360
      TabIndex        =   4
      Top             =   210
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvDatos 
      Height          =   7455
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   18015
      _ExtentX        =   31776
      _ExtentY        =   13150
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   270
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   161087489
      CurrentDate     =   41813
   End
   Begin VB.Label lblLabel2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mesa:"
      Height          =   195
      Left            =   5400
      TabIndex        =   2
      Top             =   360
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   570
   End
End
Attribute VB_Name = "frmConsultaDespachos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdMostrar_Click()

    Dim oRSdata As ADODB.Recordset

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_CONSULTA_DESPACHOS"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , Me.dtpFecha.Value)

    If Me.DatMesas.BoundText <> "-1" Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MESA", adVarChar, adParamInput, 10, Me.DatMesas.BoundText)
    
    Set oRSdata = oCmdEjec.Execute
    Me.lvDatos.ListItems.Clear

    Do While Not oRSdata.EOF
        Set itemX = Me.lvDatos.ListItems.Add(, , oRSdata!Comanda)
        itemX.SubItems(1) = oRSdata!mesa
        itemX.SubItems(2) = oRSdata!HEE
        itemX.SubItems(3) = Trim(oRSdata!ENTRADAS)
        
        itemX.SubItems(4) = Trim(oRSdata!HSE)
        itemX.SubItems(5) = Trim(oRSdata!HES)
        itemX.SubItems(6) = Trim(oRSdata!SEGUNDOS)
        itemX.SubItems(7) = Trim(oRSdata!HSS)
        itemX.SubItems(8) = Trim(oRSdata!PERS)
        oRSdata.MoveNext
    Loop

End Sub

Private Sub cmdVer_Click()



frmConsultaDespachosPrint.gMesa = Me.DatMesas.BoundText
frmConsultaDespachosPrint.gFecha = Me.dtpFecha.Value
frmConsultaDespachosPrint.Show
End Sub

Private Sub Form_Load()
  CentrarFormulario MDIForm1, Me
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_CONSULTA_DESPACHOS_LISTARMESAS"

    Dim oRS As ADODB.Recordset

    Set oRS = oCmdEjec.Execute(, LK_CODCIA)

    Set Me.DatMesas.RowSource = oRS
    Me.DatMesas.BoundColumn = oRS.Fields(0).Name
    Me.DatMesas.ListField = oRS.Fields(1).Name
    Me.DatMesas.BoundText = "-1"
    Me.dtpFecha.Value = LK_FECHA_DIA
    
        With Me.lvDatos
            .View = lvwReport
            .LabelEdit = lvwManual
            .FullRowSelect = True
            .Gridlines = True
            .ColumnHeaders.Add , , "COMANDA", 1200
            .ColumnHeaders.Add , , "MESA", 1000
            .ColumnHeaders.Add , , "HEE", 1500
            .ColumnHeaders.Add , , "ENTRADAS", 5000
            .ColumnHeaders.Add , , "HSE", 1500
            .ColumnHeaders.Add , , "HES", 1500
            .ColumnHeaders.Add , , "SEGUNDOS", 5000
            .ColumnHeaders.Add , , "HSS", 1500
            .ColumnHeaders.Add , , "PERS", 700
        End With

  

End Sub
