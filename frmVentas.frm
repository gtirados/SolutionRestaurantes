VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas de Entradas"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
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
   ScaleHeight     =   7080
   ScaleWidth      =   11145
   Begin VB.CommandButton cmdSunat 
      Caption         =   "Sunat"
      Height          =   360
      Left            =   10200
      TabIndex        =   27
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdFP 
      Caption         =   "Forma Pago"
      Height          =   465
      Left            =   3000
      TabIndex        =   26
      Top             =   6360
      Width           =   1815
   End
   Begin MSComctlLib.ListView listview1 
      Height          =   1455
      Left            =   4080
      TabIndex        =   20
      Top             =   840
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   2880
      TabIndex        =   21
      Top             =   1680
      Width           =   7215
      Begin VB.CheckBox chkEdit 
         Caption         =   "Edit"
         Height          =   255
         Left            =   4440
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtNro 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   22
         Top             =   270
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DatTiposDoctos 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   255
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "TiposDoctos"
      End
      Begin VB.Label lblSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2280
         TabIndex        =   25
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.TextBox txtDireccion 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   18
      Top             =   1200
      Width           =   6015
   End
   Begin VB.TextBox txtRuc 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   17
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Del"
      Height          =   360
      Left            =   10200
      TabIndex        =   15
      Top             =   3120
      Width           =   630
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "&Limpiar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8640
      TabIndex        =   12
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7200
      TabIndex        =   11
      Top             =   6360
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvVenta 
      Height          =   3015
      Left            =   2880
      TabIndex        =   10
      Top             =   2640
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5318
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
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox txtCliente 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      TabIndex        =   2
      Top             =   480
      Width           =   6015
   End
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton cmdEntradaSig 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   120
         TabIndex        =   14
         Top             =   5880
         Width           =   2415
      End
      Begin VB.CommandButton cmdEntradaAnt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   2415
      End
      Begin VB.CommandButton cmdEntrada 
         Caption         =   "Entrada"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCIÓN:"
      Height          =   195
      Left            =   2940
      TabIndex        =   19
      Top             =   1283
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC:"
      Height          =   195
      Left            =   3600
      TabIndex        =   16
      Top             =   923
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Igv:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6000
      TabIndex        =   9
      Top             =   5760
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8280
      TabIndex        =   8
      Top             =   5760
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2880
      TabIndex        =   7
      Top             =   5760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblIgv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   5708
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9120
      TabIndex        =   5
      Top             =   5715
      Width           =   1590
   End
   Begin VB.Label lblSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   5708
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE:"
      Height          =   195
      Left            =   3240
      TabIndex        =   1
      Top             =   555
      Width           =   810
   End
End
Attribute VB_Name = "frmVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loc_key  As Integer
Private vPagActEnt, vPagTotEnt As Integer
Private vIniLeft As Integer
Private vIniTop As Integer
Private vBuscar As Boolean 'variable para la busqueda de clientes
Private ORStd As ADODB.Recordset 'VARIABLE PARA SABER SI EL TIPO DE DOCUMENTO ES EDITABLE


Private Sub CrearRecordSet()

If oRSfp.State = 1 Then
oRSfp.Close
Set oRSfp = Nothing
End If



        Set oRSfp = New ADODB.Recordset
        oRSfp.CursorType = adOpenDynamic ' setting cursor type
        oRSfp.Fields.Append "IDFORMAPAGO", adInteger
        oRSfp.Fields.Append "FORMAPAGO", adVarChar, 100
        oRSfp.Fields.Append "REFERENCIA", adVarChar, 200
        oRSfp.Fields.Append "MONTO", adDouble
        oRSfp.Fields.Append "PAGACON", adDouble
        oRSfp.Fields.Append "VUELTO", adDouble
        oRSfp.Fields.Append "TIPO", adVarChar, 2
        'oRSfp.Fields.Append "formapago", adVarChar, 120
    
        oRSfp.Fields.Refresh
        oRSfp.Open
        
        
        
           oRSfp.AddNew
            oRSfp!idformapago = 1
            oRSfp!formaPAGO = "EFECTIVO" ' Me.lblFormaPago(DD).Caption
            oRSfp!referencia = "" 'IIf(Me.cbomoneda(DD).ListIndex = 0, "S", "D")
            oRSfp!monto = Me.lblTotal.Caption
            oRSfp!tipo = "E"
            oRSfp!pagacon = Me.lblTotal.Caption
            oRSfp!VUELTO = 0
            oRSfp.Update
      
End Sub


Private Sub chkEdit_Click()
Me.txtNro.Enabled = Me.chkEdit.Value
If Me.chkEdit.Value Then
Me.txtNro.SetFocus
Me.txtNro.SelStart = 0
Me.txtNro.SelLength = Len(Me.txtNro.Text)
End If
End Sub

Private Sub cmdDel_Click()
If Me.lvVenta.ListItems.count = 0 Then Exit Sub
Me.lvVenta.ListItems.Remove Me.lvVenta.SelectedItem.index
sumatoria
End Sub

Private Sub cmdEntrada_Click(index As Integer)

    frmVentasCantidad.Show vbModal

    If frmVentasCantidad.gAcepta Then

        Dim itemX As Object

        Set itemX = Me.lvVenta.ListItems.Add(, , frmVentasCantidad.gCantidad)
        itemX.Tag = Split(Me.cmdEntrada(index).Tag, "|")(0)
        itemX.SubItems(1) = Me.cmdEntrada(index).Caption
        itemX.SubItems(2) = Split(Me.cmdEntrada(index).Tag, "|")(1)
        itemX.SubItems(3) = val(frmVentasCantidad.gCantidad) * CDec(Split(Me.cmdEntrada(index).Tag, "|")(1))
        sumatoria
        CrearRecordSet
    End If

End Sub

Private Sub cmdEntradaAnt_Click()
Dim ini, fin, f As Integer
If vPagActEnt = 2 Then
    ini = 1
    fin = ini * 5
ElseIf vPagActEnt = 1 Then
    Exit Sub
Else
    FF = vPagActEnt - 1
    ini = (5 * FF) - 4
    fin = 5 * FF
End If

For f = ini To fin
    Me.cmdEntrada(f).Visible = True
Next
If vPagActEnt > 1 Then
vPagActEnt = vPagActEnt - 1
    If vPagActEnt = 1 Then: Me.cmdEntradaAnt.Enabled = False
    
    Me.cmdEntradaSig.Enabled = True
End If
End Sub

Private Sub cmdEntradaSig_Click()
Dim ini, fin, f As Integer
If vPagActEnt = 1 Then
    ini = 1
    fin = ini * 5
ElseIf vPagActEnt = vPagTotEnt Then
    Exit Sub
Else
    ini = (5 * vPagActEnt) - 4
    fin = 5 * vPagActEnt
End If

For f = ini To fin
    Me.cmdEntrada(f).Visible = False
Next
If vPagActEnt < vPagTotEnt Then
vPagActEnt = vPagActEnt + 1
    If vPagActEnt = vPagTotEnt Then: Me.cmdEntradaSig.Enabled = False
    
    Me.cmdEntradaAnt.Enabled = True
End If
End Sub

Private Sub cmdFP_Click()

'If oRSfp.RecordCount = 0 Then
'            oRSfp.AddNew
'            oRSfp!idformapago = 1
'            oRSfp!formapago = "EFECTIVO" ' Me.lblFormaPago(DD).Caption
'            oRSfp!referencia = "" 'IIf(Me.cbomoneda(DD).ListIndex = 0, "S", "D")
'            oRSfp!monto = Me.lblTotal.Caption
'            oRSfp!tipo = "E"
'            oRSfp!pagacon = Me.lblTotal.Caption
'            oRSfp!VUELTO = 0
'            oRSfp.Update
'
'End If
frmVentas2FP.lblTotalPagar.Caption = FormatCurrency(Me.lblTotal.Caption, 2)
frmVentas2FP.Show vbModal
End Sub

Private Sub cmdGrabar_Click()

     If Me.DatTiposDoctos.BoundText = "01" And Len(Trim(Me.txtRuc.Text)) = 0 Then
        MsgBox "Debe ingresar el cliente para una factura.", vbCritical, Pub_Titulo

        Exit Sub

    End If

    If Me.lvVenta.ListItems.count = 0 Then
        MsgBox "Debe agregar productos.", vbInformation, Pub_Titulo

        Exit Sub

    End If

    On Error GoTo xGraba

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_ENTRADAS_REGISTRAR"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim xITEM As Object

    Dim xDET  As String

    xDET = ""

    If Me.lvVenta.ListItems.count <> 0 Then
        xDET = "<r>"

        For Each xITEM In Me.lvVenta.ListItems

            xDET = xDET + "<d "
            xDET = xDET + "cp=""" & xITEM.Tag & """ "
            xDET = xDET + "st=""" & xITEM.Text & """ "
            xDET = xDET + "pr=""" & xITEM.SubItems(2) & """ "
            xDET = xDET + "un=""" & "UND" & """ "
            xDET = xDET + "sc=""" & xITEM.index & """ "
            xDET = xDET + "cam=""" & "" & """ "
            xDET = xDET + "/>"
        Next

        xDET = xDET + "</r>"
    End If
    
    
     Dim xPag  As String

    xPag = ""

    If oRSfp.RecordCount > 0 Then
        xPag = "<r>"

        For i = 1 To oRSfp.RecordCount

            xPag = xPag + "<d "
            xPag = xPag + "idfp=""" & oRSfp!idformapago & """ "
            xPag = xPag + "fp=""" & oRSfp!formaPAGO & """ "
            xPag = xPag + "mon=""" & "S" & """ "
            xPag = xPag + "monto=""" & oRSfp!monto & """ "
            xPag = xPag + "ref=""" & oRSfp!referencia & """ "
            xPag = xPag + "/>"
            oRSfp.MoveNext
        Next

        xPag = xPag + "</r>"
    End If
    

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SerDoc", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NroDoc", adInteger, adParamInput, , Me.txtNro.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1))

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@XMLDET", adVarChar, adParamInput, 4000, xDET)

    '    If Me.DatTiposDoctos.BoundText = "01" Then GTS
'        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcli", adInteger, adParamInput, , Me.txtRuc.Tag)
'    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcli", adInteger, adParamInput, , IIf(Len(Trim(Me.txtRuc.Tag)) = 0, 1, Me.txtRuc.Tag))
        
'    End If

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@totalfac", adDouble, adParamInput, , Me.lblTotal.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@diascre", adDouble, adParamInput, , 0)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@farjabas", adInteger, adParamInput, , 0)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGODOCTO", adChar, adParamInput, 2, Me.DatTiposDoctos.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@xmlFP", adVarChar, adParamInput, 4000, xPag)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@AUTONUMFAC", adInteger, adParamOutput, , 0)    'gts agregado

    oCmdEjec.Execute
    
    Me.txtNro.Text = oCmdEjec.Parameters("@AUTONUMFAC").Value   'gts agregado
 
    If MsgBox("¿Desea imprimir?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
        Imprimir Left(Me.DatTiposDoctos.Text, 1), False
    End If
    
    'If Me.DatTiposDoctos.BoundText = "01" Or Me.DatTiposDoctos.BoundText = "03" Then
    If Me.DatTiposDoctos.BoundText = "01" Then
     CrearArchivoPlano Left(Me.DatTiposDoctos.Text, 1), Me.lblSerie.Caption, Me.txtNro.Text
    End If
    
    Me.lvVenta.ListItems.Clear
    Me.txtCliente.Text = ""
    Me.txtDireccion.Text = ""
    Me.lblSubTotal.Caption = "0.00"
    Me.lblIgv.Caption = "0.00"
    Me.txtRuc.Text = ""
    Me.lblTotal.Caption = "0.00"
    
    Me.txtDireccion.Text = ""
   ' Me.DatTiposDoctos.BoundText = "03"
    
    Exit Sub

xGraba:
    MsgBox Err.Description
End Sub

Private Sub cmdLimpiar_Click()
Limpiar
End Sub

Private Sub cmdSunat_Click()
   If Len(Trim(Me.txtRuc.Text)) = 0 Then
        If IsNumeric(Me.txtCliente.Text) Then
            If Len(Trim(Me.txtCliente.Text)) = 11 Then
                frmFacComandaSunat.gRUC = Trim(Me.txtCliente.Text)
                frmFacComandaSunat.Show vbModal

                If frmFacComandaSunat.gAcepta Then
                Dim xsalida As Integer
                xsalida = 0
                    
                    Me.txtCliente.Text = Trim(frmFacComandaSunat.gRS)
                    Me.txtRuc.Text = frmFacComandaSunat.gRUC
                    Me.txtDireccion.Text = Trim(frmFacComandaSunat.gDIR)
                    LimpiaParametros oCmdEjec
                    oCmdEjec.CommandText = "SP_CLIENTES_UPDATE_DATOS_SUNAT"
                    oCmdEjec.CommandType = adCmdStoredProc
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RAZONSOCIAL", adVarChar, adParamInput, 200, frmFacComandaSunat.gRS)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIRECCION", adVarChar, adParamInput, 200, Left(frmFacComandaSunat.gDIR, 200))
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RUC", adVarChar, adParamInput, 11, frmFacComandaSunat.gRUC)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , 0)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SALIDA", adBigInt, adParamOutput, , xsalida)
                    oCmdEjec.Execute
                    Me.txtRuc.Tag = oCmdEjec.Parameters("@salida").Value
                End If

            Else
                MsgBox "El ruc debe tener 11 caracteres", vbCritical, Pub_Titulo
            End If

        Else
            MsgBox "El ruc debe ser Numeros", vbCritical, Pub_Titulo
        End If

    Else
        frmFacComandaSunat.gRUC = Trim(Me.txtRuc.Text)
        frmFacComandaSunat.Show vbModal

        If frmFacComandaSunat.gAcepta Then
        Dim xdato As Double
        xdato = Me.txtRuc.Tag
            
            Me.txtCliente.Text = Trim(frmFacComandaSunat.gRS)
            Me.txtRuc.Tag = xdato
            Me.txtDireccion.Text = Trim(frmFacComandaSunat.gDIR)
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SP_CLIENTES_UPDATE_DATOS_SUNAT"
            oCmdEjec.CommandType = adCmdStoredProc
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RAZONSOCIAL", adVarChar, adParamInput, 200, frmFacComandaSunat.gRS)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIRECCION", adVarChar, adParamInput, 200, Left(frmFacComandaSunat.gDIR, 200))
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RUC", adVarChar, adParamInput, 11, frmFacComandaSunat.gRUC)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , xdato)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SALIDA", adBigInt, adParamOutput, , 0)
            oCmdEjec.Execute
            'MsgBox Me.txtRuc.Tag
        End If
    End If
End Sub

Private Sub DatTiposDoctos_Click(Area As Integer)
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_SERIES_CARGAR"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim xSerie As String

    Dim xNro   As Double

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1)) ' Me.cboTipoDocto.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamOutput, 3, 1)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MAXIMO", adBigInt, adParamOutput, , 1)
    oCmdEjec.Execute

    xSerie = oCmdEjec.Parameters("@SERIE").Value
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)

    Me.lblSerie.Caption = xSerie
    Me.txtNro.Text = oCmdEjec.Parameters("@MAXIMO").Value

    ORStd.Filter = "CODIGO='" & Me.DatTiposDoctos.BoundText & "'"

    If ORStd.RecordCount <> 0 Then
        Me.chkEdit.Enabled = ORStd!Editable
    End If

    ORStd.Filter = ""
    Me.txtNro.Enabled = False
    Me.chkEdit.Value = False
End Sub

Private Sub Form_Load()
ConfiguraLV
CentrarFormulario MDIForm1, Me
CrearRecordSet
    vIniTop = 120
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_ARTICULO_ENTRADAS"

    Dim ORSe As ADODB.Recordset

    Set ORSe = oCmdEjec.Execute(, LK_CODCIA)

    Dim vEntradas As Integer

    vEntradas = ORSe.RecordCount

    Dim c As Integer

    c = 1
    
    Dim valor As Double

    valor = vEntradas / 5

    pos = InStr(Trim(str(valor)), ".")
    pos2 = Right(Trim(str(valor)), Len(Trim(str(valor))) - pos)
    If pos <> 0 Then
     ent = Left(Trim(str(valor)), pos - 1)
     Else
     ent = 0
    End If
    
   

    If ent = "" Then: ent = 0
    If pos2 > 0 Then: vPagTotEnt = ent + 1

    If vEntradas >= 1 Then: vPagActEnt = 1
    If vEntradas > 5 Then: Me.cmdEntradaSig.Enabled = True

    For i = 1 To vEntradas
        Load Me.cmdEntrada(i)
    
        vIniTop = vIniTop + Me.cmdEntradaAnt.Height
        Me.cmdEntrada(i).Left = Me.cmdEntradaAnt.Left
        Me.cmdEntrada(i).Top = vIniTop

        'Me.cmdEntrada(i).Left = vIniLeft
        Me.cmdEntrada(i).Top = vIniTop
        'Me.cmdFam(i).Visible = vPri
        Me.cmdEntrada(i).Visible = True
        Me.cmdEntrada(i).Caption = Trim(ORSe!Prod)
        Me.cmdEntrada(i).Tag = ORSe!cod & "|" & ORSe!pv1
   
        '    If c <= 14 Then
        '        Me.cmdFam(i).Visible = True
        '    Else
        '        Me.cmdFam(i).Visible = False
        '    End If
        ORSe.MoveNext

        If c = 5 Then
            '        vPri = False
            c = 1
            'vuelve a empezar
            vIniTop = 120
        Else
            c = c + 1
        End If
   
    Next
    'CargarNumeracion "B"
      LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_TIPOS_DOCTOS_LIST"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    Set ORStd = oCmdEjec.Execute
    Set Me.DatTiposDoctos.RowSource = ORStd
    Me.DatTiposDoctos.ListField = ORStd.Fields(1).Name
    Me.DatTiposDoctos.BoundColumn = ORStd.Fields(0).Name

    If ORStd.RecordCount <> 0 Then
        Me.DatTiposDoctos.BoundText = ORStd!Codigo
    
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_SERIES_CARGAR"
        oCmdEjec.CommandType = adCmdStoredProc

        Dim xSerie As String

        Dim xNro   As Double

        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1)) ' Me.cboTipoDocto.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamOutput, 3, 1)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MAXIMO", adBigInt, adParamOutput, , 1)
        oCmdEjec.Execute

        xSerie = oCmdEjec.Parameters("@SERIE").Value
        'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)

        Me.lblSerie.Caption = xSerie
        Me.txtNro.Text = oCmdEjec.Parameters("@MAXIMO").Value
        ORStd.Filter = "CODIGO='" & Me.DatTiposDoctos.BoundText & "'"

        If ORStd.RecordCount <> 0 Then
            Me.chkEdit.Enabled = ORStd!Editable
        End If

        ORStd.Filter = ""
    End If
    
    Me.lblTotal.Caption = "0.00"
End Sub

Private Sub ConfiguraLV()
With Me.lvVenta

    .Gridlines = True
    .LabelEdit = lvwManual
    .FullRowSelect = True
    .View = lvwReport
    
    .HideSelection = False
    .ColumnHeaders.Add , , "Cant."
    .ColumnHeaders.Add , , "Entrada"
    .ColumnHeaders.Add , , "Precio"
    .ColumnHeaders.Add , , "Importe"
    


End With

With Me.ListView1
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Codigo", 1000
    .ColumnHeaders.Add , , "Cliente", 5000
    .ColumnHeaders.Add , , "Ruc", 0
    .ColumnHeaders.Add , , "Direcion", 0
    .MultiSelect = False
End With
End Sub

Private Sub sumatoria()

    Dim itemX As Object

    Dim vT    As Double

    vT = 0

    For Each itemX In Me.lvVenta.ListItems

        vT = vT + itemX.SubItems(3)
    Next
 
    Me.lblTotal.Caption = FormatCurrency(vT)
    Me.lblSubTotal.Caption = FormatCurrency(Round(CDec(vT) / CDec((LK_IGV / 100) + 1), 2))
    Me.lblIgv.Caption = FormatCurrency(CDec(vT) - CDec(Me.lblSubTotal.Caption), 2)

End Sub

Private Sub CrearArchivoPlano(cTipoDocto As String, cSerie As String, cNumero As Double)

    Dim oRS As ADODB.Recordset

    LimpiaParametros oCmdEjec

    If cTipoDocto = "F" Then
           oCmdEjec.CommandText = "SP_VENTA_FACTURA_SFS"
    ElseIf cTipoDocto = "B" Then
           oCmdEjec.CommandText = "SP_VENTA_BOLETA_SFS"
    ElseIf LK_CODTRA = 1111 Then
    
    End If
    
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@serie", adVarChar, adParamInput, 3, IIf(LK_CODTRA = 1111, PUB_NUMSER_C, cSerie))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numero", adDouble, adParamInput, , IIf(LK_CODTRA = 1111, PUB_NUMFAC_C, cNumero))
    If LK_CODTRA = 1111 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adChar, adParamInput, 1, PUB_FBG)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TRANSACCION", adBigInt, adParamInput, , LK_CODTRA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    End If
    
    Set oRS = oCmdEjec.Execute
    
    Dim sCadena As String

    sCadena = ""
    
    Dim obj_FSO     As Object

    Dim ArchivoCab  As Object
    Dim ArchivoTri As Object
    Dim ArchivoDet  As Object
    Dim ArchivoLey As Object
    
    Dim sARCHIVOcab As String

    Dim sARCHIVOdet As String
    Dim sARCHIVOtri As String
    Dim sARCHIVOley As String

    Dim sRUC        As String
    
    sRUC = Leer_Ini(App.Path & "\config.ini", "RUC", "C:\")
     
    sARCHIVOcab = sRUC & "-" & oRS!Nombre + IIf(LK_CODTRA = 2412, ".not", IIf(LK_CODTRA = 1111, ".cba", ".cab"))
    sARCHIVOtri = sRUC & "-" & oRS!Nombre + IIf(LK_CODTRA = 2412, ".not", IIf(LK_CODTRA = 1111, ".tri", ".tri"))
    sARCHIVOley = sRUC & "-" & oRS!Nombre + IIf(LK_CODTRA = 2412, ".not", IIf(LK_CODTRA = 1111, ".ley", ".ley"))
    
    If LK_CODTRA <> 1111 Then
        sARCHIVOdet = sRUC & "-" & oRS!Nombre + ".det"
    End If
    
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")

    'Creamos un archivo con el método CreateTextFile
    Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
    
    Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
    Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOley, True)
    If LK_CODTRA <> 1111 Then
    Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
    End If
    'Set Archivo = obj_FSO.CreateTextFile("C:\" + sARCHIVO, True)
    
    If LK_CODTRA = 2412 Then

        Do While Not oRS.EOF
            sCadena = sCadena & oRS!fecemision & "|" & oRS!CODMOTIVO & "|" & oRS!DESCMOTIVO & "|" & oRS!TIPODOCAFECTADO & "|" & oRS!NUMDOCAFECTADO & "|" & oRS!TIPDOCUSUARIO & "|" & oRS!NUMDOCUSUARIO & "|" & oRS!CLI1 & "|" & oRS!TIPMONEDA & "|" & oRS!SUMOTROSCARGOS & "|" & oRS!MTOOPERGRAVADAS & "|" & oRS!MTOOPERINAFECTAS & "|" & oRS!MTOOPEREXONERADAS & "|" & oRS!MTOIGV & "|" & oRS!MTOISC & "|" & oRS!MTOOTROSTRIBUTOS & "|" & oRS!MTOIMPVENTA & "|"
            oRS.MoveNext
        Loop
    
    ElseIf LK_CODTRA = 1111 Then
         Do While Not oRS.EOF
            sCadena = sCadena & oRS!FEC_GENERACcION & "|" & oRS!FEC_COMUNICACION & "|" & oRS!TIPDOCBAJA & "|" & oRS!NUMDOCBAJA & "|" & oRS!DESMOTIVOBAJA & "|"
            oRS.MoveNext
        Loop
    Else

        Do While Not oRS.EOF
            sCadena = sCadena & oRS!TIPOPERACION & "|" & oRS!fecemision & "|" & oRS!hORA & "|" & oRS!FECHAVENC & "|" & oRS!codlocalemisor & "|" & oRS!TIPDOCUSUARIO & "|" & oRS!NUMDOCUSUARIO & "|" & oRS!rznsocialusuario & "|" & oRS!TIPMONEDA & "|" & oRS!MTOIGV & "|" & oRS!MTOOPERGRAVADAS & "|" & oRS!MTOIMPVENTA & "|" & oRS!SUMDSCTOGLOBAL & "|" & oRS!SUMOTROSCARGOS & "|" & oRS!TOTANTICIPOS & "|" & oRS!IMPTOTALVENTA & "|" & oRS!UBL & "|" & oRS!CUSTOMDOC & "|"
            oRS.MoveNext
        Loop

    End If
   
    'Escribimos lineas
    ArchivoCab.WriteLine sCadena
    
    'Cerramos el fichero
    ArchivoCab.Close
    Set ArchivoCab = Nothing
   
    Dim oRSdet As ADODB.Recordset

    Set oRSdet = oRS.NextRecordset
   
    sCadena = ""
    Dim c As Integer
    c = 1

    If LK_CODTRA = 2412 Then

        Do While Not oRSdet.EOF
         
            sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & oRSdet!DESITEM & "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!CODTIPTRIBUTOIGV & "|" & oRSdet!MTOIGVITEM1 & "|" & oRSdet!NOMTRIBITEM & "|" & oRSdet!CODTIPTRIBUTOITEM & "|" & oRSdet!TIPAFEIGV & "|" & FormatNumber(oRSdet!PORCIGV, 2) & "|" & oRSdet!CODISC & "|" & oRSdet!CODOTROITEM & "|" & oRSdet!GRATUITO & "|"
            If c < oRSdet.RecordCount Then
                sCadena = sCadena + vbCrLf
            End If
             c = c + 1
            oRSdet.MoveNext
            
        Loop

    ElseIf LK_CODTRA <> 1111 Then
    

        Do While Not oRSdet.EOF
       
           ' sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & oRSdet!DESITEM & "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTODSCTOITEM & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!TIPAFEIGV & "|" & oRSdet!MTOISCITEM & "|" & oRSdet!TIPSISISC & "|" & oRSdet!MTOPRECIOVENTAITEM & "|" & oRSdet!MTOVALORVENTAITEM & "|"
           sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & Trim(oRSdet!DESITEM) & "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!CODTIPTRIBUTOIGV & "|" & oRSdet!MTOIGVITEM1 & "|" & oRSdet!BASEIMPIGV & "|" & oRSdet!NOMTRIBITEM & "|" & oRSdet!CODTIPTRIBUTOITEM & "|" & oRSdet!TIPAFEIGV & "|" & FormatNumber(oRSdet!PORCIGV, 2) & "|" & oRSdet!CODISC & "|" & oRSdet!MONTOISC & "|" & oRSdet!BASEIMPONIBLEISC & "|" & oRSdet!NOMBRETRIBITEM & "|" & oRSdet!CODTRIBITEM & "|" & oRSdet!CODSISISC & "|" & oRSdet!PORCISC & "|" & oRSdet!CODTRIBOTO & "|" & oRSdet!MONTOTRIBOTO & "|" & oRSdet!BASEIMPONIBLEOTO & "|" & oRSdet!NOMBRETRIBOTO & "|" & oRSdet!TIPSISISC & "|" & oRSdet!PORCOTO & "|" & oRSdet!PRECIOVTAUNITARIO & "|" & oRSdet!VALORVTAXITEM & "|" & oRSdet!GRATUITO & "|"
            If c < oRSdet.RecordCount Then
                sCadena = sCadena + vbCrLf
            End If
             c = c + 1
            oRSdet.MoveNext
             
        Loop

    End If

    'Escribimos lineas
    If LK_CODTRA <> 1111 Then
    ArchivoDet.WriteLine sCadena
    
     'Cerramos el fichero
    ArchivoDet.Close
    Set ArchivoDet = Nothing
    
    Dim orsTri As ADODB.Recordset
    Set orsTri = oRS.NextRecordset
    
    sCadena = ""
    c = 1
    'ARCIVO .TRI
    Do While Not orsTri.EOF
    sCadena = sCadena & orsTri!Codigo & "|" & orsTri!Nombre & "|" & orsTri!cod & "|" & orsTri!BASEIMPONIBLE & "|" & orsTri!TRIBUTO & "|"
    If c < orsTri.RecordCount Then
        sCadena = sCadena & vbCrLf
    End If
    c = c + 1
        orsTri.MoveNext
    Loop
    
    
     ArchivoTri.WriteLine sCadena
    
     'Cerramos el fichero
    ArchivoTri.Close
    Set ArchivoTri = Nothing
    
    Dim orsLey As ADODB.Recordset
    Set orsLey = oRS.NextRecordset
    
    c = 1
    sCadena = ""
    Do While Not orsLey.EOF
        sCadena = sCadena & orsLey!cod & "|" & Trim(CONVER_LETRAS(Me.lblTotal.Caption, "S")) & "|"
        If c < orsLey.RecordCount Then
            sCadena = sCadena & vbCrLf
        End If
        c = c + 1
        orsLey.MoveNext
    Loop
    
    ArchivoLey.WriteLine sCadena
    ArchivoLey.Close
    Set ArchivoLey = Nothing
    
    End If
    
   
    
    Set obj_FSO = Nothing
End Sub



Private Sub Limpiar()
Me.lblIgv.Caption = FormatCurrency(0)
Me.lblSubTotal.Caption = FormatCurrency(0)
Me.lblTotal.Caption = FormatCurrency(0)
Me.lvVenta.ListItems.Clear
Me.txtCliente.Text = ""

End Sub




Private Sub txtCliente_Change()
vBuscar = True
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key = loc_key + 1

        If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 38 Then
        loc_key = loc_key - 1

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 34 Then
        loc_key = loc_key + 17

        If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 33 Then
        loc_key = loc_key - 17

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 27 Then
        Me.ListView1.Visible = False
        Me.txtCliente.Text = ""
        Me.txtRuc.Text = ""
        Me.txtDireccion.Text = ""
    End If

    GoTo fin
posicion:
    ListView1.ListItems.Item(loc_key).Selected = True
    ListView1.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    txtCliente.SelStart = Len(Me.txtCliente.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscar Then
            Me.ListView1.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpListarCliProv"
            Set oRsPago = oCmdEjec.Execute(, Array(LK_CODCIA, "C", Me.txtCliente.Text))

            Dim Item As Object
        
            If Not oRsPago.EOF Then

                Do While Not oRsPago.EOF
                    Set Item = Me.ListView1.ListItems.Add(, , oRsPago!CodClie)
                    Item.SubItems(1) = Trim(oRsPago!Nombre)
                    Item.SubItems(2) = IIf(IsNull(oRsPago!RUC), "", oRsPago!RUC)
                    Item.SubItems(3) = Trim(oRsPago!dir)
                    oRsPago.MoveNext
                Loop

                Me.ListView1.Visible = True
                Me.ListView1.ListItems(1).Selected = True
                loc_key = 1
                Me.ListView1.ListItems(1).EnsureVisible
                vBuscar = False
            Else
Me.txtRuc.Text = ""
Me.txtDireccion.Text = ""
                If MsgBox("Cliente no existe." + vbCrLf + "¿Desea Crearlo.?", vbQuestion + vbYesNo, "Restaurantes") = vbYes Then
                    frmCLI.Show vbModal
                End If
            End If
        
        Else
            Me.txtRuc.Tag = Me.ListView1.ListItems(loc_key)
            Me.txtRuc.Text = Me.ListView1.ListItems(loc_key).SubItems(2)
            Me.txtDireccion.Text = Me.ListView1.ListItems(loc_key).SubItems(3)
            Me.txtCliente.Text = Me.ListView1.ListItems(loc_key).SubItems(1)
            Me.ListView1.Visible = False
            Me.lvVenta.SetFocus
        End If
    End If
End Sub

Private Sub Imprimir(TipoDoc As String, Esconsumo As Boolean)

    On Error GoTo printe

    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal  As New CRAXDRT.APPLICATION

    Dim vIgv        As Currency

    Dim vSubTotal   As Currency

    Dim RutaReporte As String

    If Esconsumo Then
        If TipoDoc = "B" Then
            RutaReporte = "C:\Admin\Nordi\BolCon.rpt"
        ElseIf TipoDoc = "F" Then
            RutaReporte = "C:\Admin\Nordi\FacCon.rpt"
            vSubTotal = Round((Me.lblTotal.Caption / ((100 + LK_IGV) / 100)), 2)
            'vSubTotal = Round((Me.lblImporte.Caption / ((100 + LK_IGV + 5) / 100)), 2)
            'vrec = Round(vSubTotal * 0.05, 2)
            vIgv = Me.lblTotal.Caption - vSubTotal
            'vIgv = Me.lblImporte.Caption - vSubTotal - vrec
       ElseIf TipoDoc = "T" Then
            RutaReporte = "C:\Admin\Nordi\TikCon.rpt"
            vSubTotal = Round((Me.lblTotal.Caption / ((100 + LK_IGV) / 100)), 2)
            'vSubTotal = Round((Me.lblImporte.Caption / ((100 + LK_IGV + 5) / 100)), 2)
            'vrec = Round(vSubTotal * 0.05, 2)
            vIgv = Me.lblTotal.Caption - vSubTotal
            'vIgv = Me.lblImporte.Caption - vSubTotal - vrec
        
        End If

    Else

        If TipoDoc = "B" Then
            RutaReporte = "C:\Admin\Nordi\BolDet.rpt"
            vSubTotal = Round((Me.lblTotal.Caption / ((100 + LK_IGV) / 100)), 2)
             vIgv = Me.lblTotal.Caption - vSubTotal
        ElseIf TipoDoc = "F" Then
            RutaReporte = "C:\Admin\Nordi\FacDet.rpt"
            vSubTotal = Round((Me.lblTotal.Caption / ((100 + LK_IGV) / 100)), 2)
            'vSubTotal = Round((Me.lblImporte.Caption / ((100 + LK_IGV + 5) / 100)), 2)
            'vrec = Round(vSubTotal * 0.05, 2)
            vIgv = Me.lblTotal.Caption - vSubTotal
            'vIgv = Me.lblImporte.Caption - vSubTotal - vrec
        ElseIf TipoDoc = "T" Then
            RutaReporte = "C:\Admin\Nordi\TikDet.rpt"
            vSubTotal = Round((Me.lblTotal.Caption / ((100 + LK_IGV) / 100)), 2)
            'vSubTotal = Round((Me.lblImporte.Caption / ((100 + LK_IGV + 5) / 100)), 2)
            'vrec = Round(vSubTotal * 0.05, 2)
            vIgv = Me.lblTotal.Caption - vSubTotal
            'vIgv = Me.lblImporte.Caption - vSubTotal - vrec
        End If
    End If

    'If TipoDoc = "B" Then
  
    'Else
    '    oCmdEjec.CommandText = "SpPrintFacDet"
    'End If

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    Set crParamDefs = VReporte.ParameterFields

    For Each crParamDef In crParamDefs

        Select Case crParamDef.ParameterFieldName

            Case "cliente"
                'crParamDef.AddCurrentValue IIf(Len(Trim(Me.txtRS.Text)) = 0, "CLIENTES VARIOS", Trim(Me.txtRS.Text))
                crParamDef.AddCurrentValue IIf(Len(Trim(Me.txtCliente.Text)) = 0, Trim(Me.txtCliente.Text), Trim(Me.txtCliente.Text))

            Case "FechaEmi"
                crParamDef.AddCurrentValue LK_FECHA_DIA

            Case "Son"
                crParamDef.AddCurrentValue CONVER_LETRAS(Me.lblTotal.Caption, "S")

            Case "total"
                crParamDef.AddCurrentValue FormatNumber(Me.lblTotal.Caption, 2)

            Case "subtotal"
                crParamDef.AddCurrentValue CStr(vSubTotal)

            Case "igv"
                crParamDef.AddCurrentValue CStr(vIgv)

            Case "SerFac"
                crParamDef.AddCurrentValue Me.lblSerie.Caption

            Case "NumFac"
                crParamDef.AddCurrentValue Me.txtNro.Text

            Case "DirClie"
                crParamDef.AddCurrentValue Me.txtDireccion.Text

            Case "RucClie"
                crParamDef.AddCurrentValue Me.txtRuc.Text

            Case "Importe" 'linea nueva
                crParamDef.AddCurrentValue FormatNumber(Me.lblTotal.Caption, 2)  'linea nueva
                ' Case "rec"          'SR BEFE
                '     crParamDef.AddCurrentValue CStr(vrec)
            Case "Dni" 'linea nueva
                crParamDef.AddCurrentValue xdni 'linea nueva

        End Select

    Next

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc

    'oCmdEjec.CommandText = "SpPrintComanda"

    Dim rsd As ADODB.Recordset

    oCmdEjec.CommandText = "SpPrintFacturacion"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Serie", adChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@nro", adInteger, adParamInput, , Me.txtNro.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fbg", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1))
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NroCom", adInteger, adParamInput, , vNroCom)

    Set rsd = oCmdEjec.Execute

    'COCINA
    'rsd.Filter = "PED_FAMILIA=2"
    Dim DD As ADODB.Recordset

    ' For i = 0 To Printers.count - 1
    '        MsgBox Printers(i).DeviceName
    '    Next
    If Not rsd.EOF Then

        VReporte.DataBase.SetDataSource rsd, 3, 1 'lleno el objeto reporte
        'VReporte.SelectPrinter Printer.DriverName, "\\laptop\doPDF v6", Printer.Port
        '
        VReporte.PrintOut False, 1, , 1, 1
        frmVisor.cr.ReportSource = VReporte
        'frmVisor.cr.ViewReport
        'frmVisor.Show vbModal
    
    End If



    Set objCrystal = Nothing
    Set VReporte = Nothing

    Exit Sub

printe:
    MostrarErrores Err

End Sub
