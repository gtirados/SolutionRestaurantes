VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmRequerimientoRecepcion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Requerimientos"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11040
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   11040
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   600
      Left            =   8760
      TabIndex        =   4
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   600
      Left            =   7320
      TabIndex        =   3
      Top             =   5280
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvDatos 
      Height          =   3735
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   6588
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
   Begin VB.TextBox txtRequerimiento 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Requerimiento:"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1680
   End
End
Attribute VB_Name = "frmRequerimientoRecepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdGrabar_Click()

    If Len(Trim(Me.txtRequerimiento.Text)) = 0 Then
        MsgBox "Debe ingresar el requerimiento.", vbInformation, Pub_Titulo
        Me.txtRequerimiento.SetFocus
    
        Exit Sub

    End If

    On Error GoTo xGraba

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_REQUERIMIENTO_INGRESO"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim Xitem As Object

    Dim xDET  As String

    xDET = ""

    If Me.lvDatos.ListItems.count <> 0 Then
        xDET = "<r>"

        For Each Xitem In Me.lvDatos.ListItems

            xDET = xDET + "<d "
            xDET = xDET + "cp=""" & Xitem.Tag & """ "
            xDET = xDET + "st=""" & Xitem.Text & """ "
            xDET = xDET + "pr=""" & "0" & """ "
            xDET = xDET + "un=""" & "UND" & """ "
            xDET = xDET + "sc=""" & Xitem.Index & """ "
            xDET = xDET + "cam=""" & "" & """ "
            xDET = xDET + "/>"
        Next

        xDET = xDET + "</r>"
    End If

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@XMLDET", adVarChar, adParamInput, 4000, xDET)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDREQ", adBigInt, adParamInput, , Me.txtRequerimiento.Text)
    '    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NroDoc", adInteger, adParamInput, , Me.txtNumero.Text)
    '    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, IIf(Me.optBoleta.Value, "B", IIf(Me.optFactura.Value, "F", "P")))
    '
    '
    '
    '    If Me.optFactura.Value Then
    '        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcli", adInteger, adParamInput, , Me.txtRuc.Tag)
    '    Else
    '        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcli", adInteger, adParamInput, , 1)
    '    End If
    '
    '    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@totalfac", adDouble, adParamInput, , Me.lblTotal.Caption)
    '    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@diascre", adDouble, adParamInput, , 0)
    '    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@farjabas", adInteger, adParamInput, , 0)

    oCmdEjec.Execute
 
    '    If MsgBox("¿Desea imprimir?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
    '        Imprimir IIf(Me.optBoleta.Value, "B", IIf(Me.optFactura.Value, "F", "P")), False
    '    End If
    
    Me.lvDatos.ListItems.Clear
    '    Me.txtCliente.Text = ""
    '    Me.txtDireccion.Text = ""
    '    Me.lblSubTotal.Caption = "0.00"
    '    Me.lblIgv.Caption = "0.00"
    '    Me.txtRuc.Text = ""
    '    Me.lblTotal.Caption = "0.00"
    '    Me.optBoleta.Value = True
    '    Me.txtDireccion.Text = ""
    '    CargarNumeracion "B"
    Me.txtRequerimiento.Text = ""

    MsgBox "Datos almacenados correctamente", vbInformation, Pub_Titulo
    
    Exit Sub

xGraba:
    MsgBox Err.Description
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
With Me.lvDatos
.ColumnHeaders.Add , , "Cantidad"
.ColumnHeaders.Add , , "Producto", 6000
'.ColumnHeaders.Add , , "idproducto"
    '.ColumnHeaders.Add , , ""

End With
End Sub

Private Sub lvDatos_DblClick()
frmRequerimientoCantidad.txtCantidad.Text = Me.lvDatos.SelectedItem.Text
frmRequerimientoCantidad.Show vbModal
If frmRequerimientoCantidad.gAcepta Then
    Me.lvDatos.SelectedItem.Text = frmRequerimientoCantidad.gCant
End If
End Sub

Private Sub txtRequerimiento_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Not IsNumeric(Me.txtRequerimiento.Text) Then
            MsgBox "Nro de requerimiento no existe.", vbCritical, Pub_Titulo

            Exit Sub

        End If

        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_REQUERIMIENTO_RECEPCION"
    
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@idreq", adBigInt, adParamInput, 80, Me.txtRequerimiento.Text)

        Dim orsP As ADODB.Recordset

        Set orsP = oCmdEjec.Execute
        Me.lvDatos.ListItems.Clear

        Dim itemX As Object
        
        If orsP.RecordCount = 0 Then
            MsgBox "Requerimiento no existe.", vbCritical, Pub_Titulo
        Else

            If CBool(orsP!pendiente) Then

                Do While Not orsP.EOF
                    Set itemX = Me.lvDatos.ListItems.Add(, , orsP!Cantidad)
                    itemX.SubItems(1) = orsP!producto
                    itemX.Tag = orsP!idproducto
                    orsP.MoveNext
                Loop

            Else
                MsgBox "Requerimiento ya fue atendido.", vbCritical, Pub_Titulo
            End If
        End If

    End If

End Sub
