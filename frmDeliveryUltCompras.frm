VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDeliveryUltCompras 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ultimos Pedidos"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9495
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
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSearch 
      Height          =   315
      Left            =   8760
      Picture         =   "frmDeliveryUltCompras.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   675
   End
   Begin MSComCtl2.DTPicker dtpDel 
      Height          =   315
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   197394433
      CurrentDate     =   41547
   End
   Begin MSComctlLib.ListView lvPedidos 
      Height          =   2655
      Left            =   30
      TabIndex        =   2
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4683
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
   Begin MSComCtl2.DTPicker dtpAl 
      Height          =   315
      Left            =   7200
      TabIndex        =   6
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   197394433
      CurrentDate     =   41547
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   2655
      Left            =   30
      TabIndex        =   8
      Top             =   3240
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4683
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Al:"
      Height          =   195
      Left            =   6960
      TabIndex        =   4
      Top             =   180
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Del:"
      Height          =   195
      Left            =   4920
      TabIndex        =   3
      Top             =   180
      Width           =   360
   End
   Begin VB.Label lblCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   780
      TabIndex        =   1
      Top             =   120
      Width           =   4035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   180
      Width           =   675
   End
End
Attribute VB_Name = "frmDeliveryUltCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gICLIENTE As Double

Private Sub ConfigurarLV()

    With Me.lvPedidos
    
        .ColumnHeaders.Add , , "SERIE"
        .ColumnHeaders.Add , , "NUMERO"
        
        .ColumnHeaders.Add , , "FECHA"
        .ColumnHeaders.Add , , "TOTAL"
        .ColumnHeaders.Add , , "ESTADO"
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True
    End With

    With Me.lvDetalle
        .ColumnHeaders.Add , , "CANT", 800
        .ColumnHeaders.Add , , "PRODUCTO", 5000
        .ColumnHeaders.Add , , "PRECIO"
        .ColumnHeaders.Add , , "IMPORTE"
        
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True

    End With

End Sub

Private Sub CargarUltimosPedidos()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_DELIVERY_DOCUMENTOS_CAB"

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , gICLIENTE)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHAINI", adDBTimeStamp, adParamInput, , Me.dtpDel.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHAFIN", adDBTimeStamp, adParamInput, , Me.dtpAl.Value)

    Dim orsDoctos As ADODB.Recordset

    Set orsDoctos = oCmdEjec.Execute

    Me.lvPedidos.ListItems.Clear

    Dim itemO As Object

    Do While Not orsDoctos.EOF
        Set itemO = Me.lvPedidos.ListItems.Add(, , orsDoctos!serie)
        
        itemO.SubItems(1) = orsDoctos!NUMERO
        
        itemO.SubItems(2) = orsDoctos!fecha
        itemO.SubItems(3) = orsDoctos!Total
        itemO.SubItems(4) = orsDoctos!ESTADO
    
        orsDoctos.MoveNext
    Loop
    If Me.lvPedidos.ListItems.count <> 0 Then Me.lvPedidos.SelectedItem.Selected = False
End Sub

Private Sub cmdSearch_Click()
CargarUltimosPedidos
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Me.dtpAl.Value = LK_FECHA_DIA
    Me.dtpDel.Value = LK_FECHA_DIA
    ConfigurarLV
    CargarUltimosPedidos

End Sub

Private Sub lvPedidos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_DELIVERY_DOCUMENTOS_DET"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , Me.lvPedidos.SelectedItem.SubItems(2))
    
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adInteger, adParamInput, , Me.lvPedidos.SelectedItem.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , Me.lvPedidos.SelectedItem.SubItems(1))

    Dim oRSdet As ADODB.Recordset

    Set oRSdet = oCmdEjec.Execute
        
    Me.lvDetalle.ListItems.Clear

Dim ITEMd As Object
    Do While Not oRSdet.EOF
    Set ITEMd = Me.lvDetalle.ListItems.Add(, , oRSdet!cant)
    ITEMd.SubItems(1) = Trim(oRSdet!producto)
    ITEMd.SubItems(2) = oRSdet!PRECIO
    ITEMd.SubItems(3) = oRSdet!Importe
        oRSdet.MoveNext
    Loop
    
End Sub
