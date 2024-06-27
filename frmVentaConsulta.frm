VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVentaConsulta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta de Documentos"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10095
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
   ScaleHeight     =   5130
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnular 
      Caption         =   "&Anular"
      Height          =   600
      Left            =   8520
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   600
      Left            =   6960
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvData 
      Height          =   4095
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   201916417
      CurrentDate     =   41976
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   201916417
      CurrentDate     =   41976
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmVentaConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAnular_Click()

    If MsgBox("¿desea continuar con la operación.?", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_FACTURACION_ANULAR"

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fecha", adDBTimeStamp, adParamInput, , Me.lvData.SelectedItem.SubItems(2))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@tipodocto", adChar, adParamInput, 1, Left(Me.lvData.SelectedItem.Text, 1))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@serie", adChar, adParamInput, 3, Me.lvData.SelectedItem.SubItems(5))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numero", adBigInt, adParamInput, , Me.lvData.SelectedItem.SubItems(6))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@usuario", adVarChar, adParamInput, 10, LK_CODUSU)
    
    Set oRSfp = oCmdEjec.Execute
    
    If Split(oRSfp!Mensaje, "=")(0) <> 0 Then
        MsgBox Split(oRSfp!Mensaje, "=")(1), vbCritical, Pub_Titulo
    Else
        MsgBox Split(oRSfp!Mensaje, "=")(1), vbInformation, Pub_Titulo
        Me.lvData.ListItems.Remove Me.lvData.SelectedItem.Index
    End If

End Sub

Private Sub cmdBuscar_Click()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_FACTURACION_SEARCH"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESDE", adDBTimeStamp, adParamInput, , Me.dtpDesde.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@HASTA", adDBTimeStamp, adParamInput, , Me.dtpHasta.Value)
    
    Set oRSfp = oCmdEjec.Execute

    Dim ITEMs As Object

    Me.lvData.ListItems.Clear

    Do While Not oRSfp.EOF
        Set ITEMs = Me.lvData.ListItems.Add(, , oRSfp!DOCTO)
        ITEMs.SubItems(1) = oRSfp!NUMERO
        ITEMs.SubItems(2) = oRSfp!fecha
        ITEMs.SubItems(3) = oRSfp!cliente
        ITEMs.SubItems(4) = oRSfp!Total
        ITEMs.SubItems(5) = oRSfp!serie
        ITEMs.SubItems(6) = oRSfp!Nro
        oRSfp.MoveNext
    Loop

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
Me.dtpDesde.Value = LK_FECHA_DIA
Me.dtpHasta.Value = LK_FECHA_DIA

With Me.lvData
.ColumnHeaders.Add , , "Docto"
.ColumnHeaders.Add , , "Numero"
.ColumnHeaders.Add , , "Fecha"
.ColumnHeaders.Add , , "Cliente"
.ColumnHeaders.Add , , "Total"
.ColumnHeaders.Add , , "serie"
.ColumnHeaders.Add , , "numero"
.View = lvwReport
.HideColumnHeaders = False
End With
End Sub

