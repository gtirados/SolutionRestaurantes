VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComandaProdCaracteristicas 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7800
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
   ScaleHeight     =   6570
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   6000
      TabIndex        =   2
      Top             =   5640
      Width           =   1710
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   4200
      TabIndex        =   0
      Top             =   5640
      Width           =   1710
   End
   Begin MSComctlLib.ListView lvCaracteristicas 
      Height          =   5535
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   9763
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmComandaProdCaracteristicas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gNUMFAC As Double
Public gNUMSER As String
Public gNUMSEC As Integer
Public gIDPRODUCTO As Double

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdGrabar_Click()
If Me.lvCaracteristicas.ListItems.count = 0 Then Exit Sub

    On Error GoTo xGraba
              
    LimpiaParametros oCmdEjec
    Pub_ConnAdo.BeginTrans
    oCmdEjec.Prepared = True
  
    oCmdEjec.CommandText = "SP_PEDIDO_CARACTERISTICA_ELIMINAR"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , gNUMFAC)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adVarChar, adParamInput, 3, gNUMSER)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSEC", adInteger, adParamInput, , gNUMSEC)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , gIDPRODUCTO)
    
    oCmdEjec.Execute
    
    Dim itemM As Object

    For Each itemM In Me.lvCaracteristicas.ListItems

        If itemM.Checked Then
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SP_PEDIDO_CARACTERISTICA_REGISTRAR"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , gNUMFAC)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adVarChar, adParamInput, 3, gNUMSER)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSEC", adInteger, adParamInput, , gNUMSEC)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , gIDPRODUCTO)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCARACTERISTICA", adInteger, adParamInput, , itemM.Tag)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CARACTERISTICA", adVarChar, adParamInput, 30, itemM.Text)
            oCmdEjec.Execute
        End If

    Next

    Pub_ConnAdo.CommitTrans
'MsgBox "Datos Almacenados correctamente.", vbInformation, Pub_Titulo
Unload Me
    Exit Sub

xGraba:
   
    MsgBox Err.Description, vbCritical, Pub_Titulo
    ' Pub_ConnAdo.RollbackTrans
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    ConfigurarLV
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_PEDIDO_PRODUCTO_FILLCARACTERISTICAS"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , gNUMFAC)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adVarChar, adParamInput, 3, gNUMSER)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSEC", adInteger, adParamInput, , gNUMSEC)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , gIDPRODUCTO)
    
    Dim orsC As ADODB.Recordset

    Dim orsM As ADODB.Recordset

    Set orsC = oCmdEjec.Execute

    Dim itemO As Object

    Do While Not orsC.EOF
        Set itemO = Me.lvCaracteristicas.ListItems.Add(, , orsC!CARAC)
        itemO.Tag = orsC!ide
        orsC.MoveNext
    Loop

    'items marcados
    Set orsM = orsC.NextRecordset

    Do While Not orsM.EOF

        For Each itemO In Me.lvCaracteristicas.ListItems

            If CStr(itemO.Tag) = CStr(orsM!ide) Then
                itemO.Checked = True

                Exit For

            End If

        Next

        orsM.MoveNext
    Loop

End Sub

Private Sub ConfigurarLV()

With Me.lvCaracteristicas
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "CARACTERISTICA", 7000
End With
End Sub

Private Sub lvCaracteristicas_ItemClick(ByVal Item As MSComctlLib.ListItem)
Item.Checked = Not Item.Checked
End Sub
