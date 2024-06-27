VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmInsumos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insumos Por Agotarse"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmri 
      Interval        =   1000
      Left            =   2760
      Top             =   2160
   End
   Begin MSFlexGridLib.MSFlexGrid msdata 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8281
      _Version        =   393216
      FixedCols       =   0
   End
End
Attribute VB_Name = "frmInsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oData As rdoResultset

Dim vquery As rdoQuery
Dim vdata As rdoResultset
Private Sub cargadata()
With Me.msdata
    .Cols = 3
    .Rows = 1
End With

Me.msdata.TextMatrix(0, 0) = "Insumo"
Me.msdata.TextMatrix(0, 1) = "Stock"
Me.msdata.RowHeight(0) = 300
Me.msdata.ColWidth(0) = 3600
Me.msdata.TextMatrix(0, 2) = "Minimo"
   
   Dim fila As Integer
   
   
   Do While Not oData.EOF
   
    fila = fila + 1
    Me.msdata.Rows = msdata.Rows + 1
    Me.msdata.TextMatrix(fila, 0) = UCase(Trim(oData!art_nombre))
    Me.msdata.TextMatrix(fila, 1) = Trim(oData!arm_stock)
    Me.msdata.TextMatrix(fila, 2) = Trim(oData!art_stock_min)
    
     oData.MoveNext
   Loop
End Sub
Private Sub Form_Load()
cargadata
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oFrmIns = Nothing
End Sub

Private Sub tmri_Timer()
pub_cadena = "select ar.art_nombre,a.arm_stock,ar.art_stock_min " & _
"from articulo a inner join arti ar on " & _
"a.arm_codcia = ar.art_codcia And a.arm_codart = ar.art_key " & _
"where ar.art_flag_stock='M' and a.arm_stock <= ar.art_stock_min " & _
"and ar.art_codcia = ? "

'pub_cadena = "SELECT sum(PED_CANTIDAD) as total  FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_FECHA >= ? AND PED_FECHA <= ?  AND  PED_CODART = ? AND PED_ESTADO <> 'E' and PED_FAMILIA = 2 AND PED_CANATEN = 0"
Set vquery = CN.CreateQuery("", pub_cadena)
'Set PS_CON02 = CN.CreateQuery("", pub_cadena)
vquery(0) = LK_CODCIA


'Set llave_con04 = PS_CON02.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)
Set vdata = vquery.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)
Set oData = vdata
cargadata
End Sub
