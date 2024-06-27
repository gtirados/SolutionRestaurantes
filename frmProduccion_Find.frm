VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProduccion_Find 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Producción"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10815
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
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   10575
      Begin MSComctlLib.ListView lvDetalle 
         Height          =   2535
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   4471
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
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   10575
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   9960
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
      Begin MSComctlLib.ListView lvCabecera 
         Height          =   2535
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4471
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
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   10575
      Begin MSComCtl2.DTPicker dtpFI 
         Height          =   375
         Left            =   1800
         TabIndex        =   0
         Top             =   270
         Width           =   1400
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         Format          =   102957057
         CurrentDate     =   44357
      End
      Begin MSComCtl2.DTPicker dtpFF 
         Height          =   375
         Left            =   4440
         TabIndex        =   1
         Top             =   270
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         Format          =   102957057
         CurrentDate     =   44357
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   360
         Left            =   6120
         Picture         =   "frmProduccion_Find.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   277
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HASTA:"
         Height          =   195
         Left            =   3600
         TabIndex        =   9
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESDE:"
         Height          =   195
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmProduccion_Find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ConfigurarLV()
With Me.lvCabecera
    .ColumnHeaders.Add , , "NUMERO", 700
    .ColumnHeaders.Add , , "PRODUCTO", 3000
    .ColumnHeaders.Add , , "CANTIDAD"
    .ColumnHeaders.Add , , "COSTO"
    .ColumnHeaders.Add , , "OPERACION", 0
    .FullRowSelect = True
    .Gridlines = True
    .View = lvwReport
    .HideSelection = False
End With
With Me.lvDetalle
    .ColumnHeaders.Add , , "NUMERO", 700
    .ColumnHeaders.Add , , "INSUMO", 3000
    .ColumnHeaders.Add , , "CANTIDAD"
    .ColumnHeaders.Add , , "COSTO"
    .ColumnHeaders.Add , , "OPERACION", 0
    .FullRowSelect = True
    .Gridlines = True
    .View = lvwReport
    .HideSelection = False
End With
End Sub

Private Sub cmdPrint_Click()
frmProduccion.Imprime Me.lvDetalle.SelectedItem.SubItems(4), Me.lvCabecera.SelectedItem.SubItems(1), Me.lvCabecera.SelectedItem.SubItems(2)
End Sub

Private Sub cmdSEARCH_Click()

    Dim oRS As ADODB.Recordset
Me.lvCabecera.ListItems.Clear
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_PRODUCCION_SEARCH]"
    oCmdEjec.CommandType = adCmdStoredProc

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FI", adDBTimeStamp, adParamInput, , Me.dtpFI.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FF", adDBTimeStamp, adParamInput, , Me.dtpFF.Value)
    
    Set oRS = oCmdEjec.Execute
 
    Dim itemX As Object
        
    If Not oRS.EOF Then

        Do While Not oRS.EOF
            Set itemX = Me.lvCabecera.ListItems.Add(, , Trim(oRS!NUMERO))
            itemX.SubItems(1) = oRS!INSUMO
            itemX.SubItems(2) = oRS!Cantidad
            itemX.SubItems(3) = oRS!costo
            itemX.SubItems(4) = oRS!operacion
            oRS.MoveNext
        Loop
     
    Else
        MsgBox "No se encontraron registros", vbInformation, Pub_Titulo
    End If
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
Me.dtpFF.Value = LK_FECHA_DIA
Me.dtpFI.Value = DateAdd("d", -30, LK_FECHA_DIA)
ConfigurarLV
End Sub

Private Sub lvCabecera_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim oRS As ADODB.Recordset
Me.lvDetalle.ListItems.Clear
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_PRODUCCION_DETALLE_SEARCH]"
    oCmdEjec.CommandType = adCmdStoredProc

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMERO", adInteger, adParamInput, , Me.lvCabecera.SelectedItem.Text)
    
    Set oRS = oCmdEjec.Execute
 
    Dim itemX As Object
        
  

        Do While Not oRS.EOF
            Set itemX = Me.lvDetalle.ListItems.Add(, , Trim(oRS!Codigo))
            itemX.SubItems(1) = oRS!INSUMO
            itemX.SubItems(2) = oRS!Cantidad
            itemX.SubItems(3) = oRS!costo
            itemX.SubItems(4) = oRS!NumOper
            oRS.MoveNext
        Loop
     
   
End Sub
