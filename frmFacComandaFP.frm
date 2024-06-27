VERSION 5.00
Object = "{BB35AEF3-E525-4F8B-81F2-511FF805ABB1}#2.1#0"; "scrollerII.ocx"
Begin VB.Form frmFacComandaFP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Formas de Pago"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   225
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescuento 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6600
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Frame FraDeno 
      Height          =   3015
      Left            =   9720
      TabIndex        =   23
      Top             =   0
      Width           =   3015
      Begin VB.CommandButton cmdBorrar 
         Height          =   480
         Left            =   1560
         TabIndex        =   11
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdDeno 
         Caption         =   "1"
         Height          =   480
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdDeno 
         Caption         =   "200"
         Height          =   480
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdDeno 
         Caption         =   "100"
         Height          =   480
         Index           =   1
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdDeno 
         Caption         =   "50"
         Height          =   480
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdDeno 
         Caption         =   "20"
         Height          =   480
         Index           =   3
         Left            =   1560
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdDeno 
         Caption         =   "10"
         Height          =   480
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdDeno 
         Caption         =   "5"
         Height          =   480
         Index           =   5
         Left            =   1560
         TabIndex        =   4
         Top             =   1560
         Width           =   1335
      End
   End
   Begin VB.TextBox txtPago 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8160
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.ComboBox cboMoneda 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   0
      ItemData        =   "frmFacComandaFP.frx":0000
      Left            =   3840
      List            =   "frmFacComandaFP.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   0
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   8160
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCUENTO"
      Height          =   195
      Left            =   6720
      TabIndex        =   24
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label lblVuelto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   8160
      TabIndex        =   22
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VUELTO:"
      Height          =   195
      Left            =   8520
      TabIndex        =   21
      Top             =   840
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAGA CON:"
      Height          =   195
      Left            =   8280
      TabIndex        =   20
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label lbltotal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   6600
      TabIndex        =   19
      Top             =   2520
      Width           =   1500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL:"
      Height          =   195
      Left            =   7035
      TabIndex        =   18
      Top             =   2280
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL MOVILIDAD:"
      Height          =   195
      Left            =   6495
      TabIndex        =   17
      Top             =   840
      Width           =   1710
   End
   Begin VB.Label lblMovilidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   6600
      TabIndex        =   16
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL PEDIDO:"
      Height          =   195
      Left            =   6660
      TabIndex        =   15
      Top             =   120
      Width           =   1380
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   6600
      TabIndex        =   14
      Top             =   360
      Width           =   1500
   End
   Begin VB.Label lblFormaPago 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   3585
   End
   Begin ScrollerII.FormScroller FormScroller1 
      Left            =   7320
      Top             =   4800
      _ExtentX        =   2170
      _ExtentY        =   1085
      SmallChange     =   100
      LargeChange     =   1
      BackColor       =   -2147483632
      ScaleMode       =   0
   End
End
Attribute VB_Name = "frmFacComandaFP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vAcepta As Boolean
Private oRStp As ADODB.Recordset
Public gMostrador As Boolean
Public gDELIVERY As Boolean
Private cMontoTarifa As Double
Private descuento As Boolean
Private pIngresoPass As Boolean

Private Sub cmdAceptar_Click()

    If gDELIVERY Then
        
        If CDec(frmDeliveryApp.lblTot.Caption) + cMontoTarifa < val(Me.lbltotal2.Caption) Then
            MsgBox "Excede", vbCritical, Pub_Titulo

            Exit Sub

        End If
        
    Else

        If frmFacComanda.lvDetalle.ListItems.count = 0 Then
            If val(frmComanda.lblTot.Caption) > val(frmFacComanda.lblImporte.Caption) Then
                MsgBox "Excede", vbCritical, Pub_Titulo

                Exit Sub

            End If

        Else

            If CDec(Me.lbltotal2.Caption) > CDec(frmFacComanda.lblImporte.Caption) Then
                MsgBox "Excede", vbCritical, Pub_Titulo

                Exit Sub

            End If
        End If

    End If
    
    '--------------OJO revisar julio------------------------------------------------
    'If CDec(Me.lblTotal.Caption) <> CDec(frmFacComanda.lblImporte.Caption) Then
    If gDELIVERY Then
        If CDec(Me.lbltotal2.Caption) <> CDec(frmDeliveryApp.lblTot.Caption) + CDec(cMontoTarifa) - val(Me.txtDescuento.Text) Then
            MsgBox "El importe no coindice", vbCritical, Pub_Titulo

            Exit Sub

        End If

    Else
    
        If gMostrador Then
            If CDec(Me.lblTotal.Caption) <> CDec(frmComanda2.lblTot.Caption) Then
                MsgBox "El importe no coindice", vbCritical, Pub_Titulo

                Exit Sub

            End If

        Else

            '        If CDec(Me.lblTotal.Caption) <> CDec(frmComanda.lblTot.Caption) Then
            If CDec(Me.lblTotal.Caption) <> CDec(frmFacComanda.lblImporte.Caption) Then
                MsgBox "El importe no coindice", vbCritical, Pub_Titulo

                Exit Sub

            End If
        End If
    End If

    '--------------OJO revisar julio------------------------------------------------
  
    If Not oRSfp Is Nothing Then

        'If Not oRSfp.EOF Then oRSfp.Delete
        If oRSfp.RecordCount <> 0 Then
            oRSfp.MoveFirst

            Do While Not oRSfp.EOF
                oRSfp.Delete adAffectCurrent
                oRSfp.MoveNext
            Loop

        End If
    End If
   
    vAcepta = True
    
    If Not oRStp.EOF Then
        oRStp.Filter = ""
        oRStp.MoveFirst
        
        Dim DD As Integer

        For DD = 1 To oRStp.RecordCount

            If Me.txtImporte(DD).Text <> "" Then
                ' oRSFPingresados.AddNew oC.Tag & vbTab & oC.Caption & vbTab & Me.txtImporte(oC.Tag).Text
            
                oRSfp.AddNew
                oRSfp!idformapago = Me.lblFormaPago(DD).Tag
                oRSfp!formapago = Me.lblFormaPago(DD).Caption
                oRSfp!moneda = IIf(Me.cboMoneda(DD).ListIndex = 0, "S", "D")
                oRSfp!monto = Me.txtImporte(DD).Text
   
                oRSfp.Update
            End If

        Next

    End If
    
    If gDELIVERY Then

        Dim sSERIE As String, sNUMERO As Double

        Dim xEXITO As String

        sSERIE = frmDeliveryApp.lblSerie.Caption
        sNUMERO = frmDeliveryApp.lblNumero.Caption
            
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_DELIVERY_ASIGNAR_FORMAPAGO"
        oCmdEjec.CommandType = adCmdStoredProc
            
        xEXITO = ""
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, sSERIE)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , sNUMERO)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PAGO", adDouble, adParamInput, , IIf(Len(Trim(Me.txtPago.Text)) = 0, 0, Me.txtPago.Text))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESCUENTO", adDouble, adParamInput, , Me.txtDescuento.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@VUELTO", adDouble, adParamInput, , Me.lblVuelto.Caption)
        
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@EXITO", adVarChar, adParamOutput, 300, xEXITO)
        oCmdEjec.Execute
        frmFacComanda.gDESCUENTO = Me.txtDescuento.Text
    
        xEXITO = oCmdEjec.Parameters("@EXITO").Value

        If Len(Trim(xEXITO)) = 0 Then
            'GRABA EN TABLA
            'borra LOS APGOS
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SP_DELIVERY_ELIMINARFORMASPAGO"
            oCmdEjec.Execute , Array(LK_CODCIA, frmDeliveryApp.lblSerie.Caption, frmDeliveryApp.lblNumero.Caption)
       
            oCmdEjec.CommandText = "SP_DELIVERY_REGISTRAFORMASPAGO"

            For i = 1 To Me.txtImporte.count - 1

                If Len(Trim(Me.txtImporte(i).Text)) <> 0 And IsNumeric(Me.txtImporte(i).Text) Then
                    LimpiaParametros oCmdEjec
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, sSERIE)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , sNUMERO)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDFORMAPAGO", adInteger, adParamInput, , Me.lblFormaPago(i).Tag) ' Me.lblFormaPago(oCTRL.Tag).Caption)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MONTO", adDouble, adParamInput, , Me.txtImporte(i).Text) '  Me.txtImporte(oCTRL.Tag).Text)
                    oCmdEjec.Execute
                End If

            Next

        Else
            MsgBox xEXITO, vbCritical, Pub_Titulo
        End If

    Else
        frmFacComanda.gDESCUENTO = Me.txtDescuento.Text
        frmFacComanda.gPAGO = IIf(Len(Trim(Me.txtPago.Text)) = 0, 0, Me.txtPago.Text)
        frmFacComanda.lblImporte.Caption = Me.lbltotal2.Caption
    End If

    If oRSfp.RecordCount <> 0 Then oRSfp.MoveFirst
    Unload Me
End Sub

Private Sub cmdBorrar_Click()

    If descuento Then
        If Not pIngresoPass Then
            frmClaveCaja.Show vbModal

            Dim vS As String

            If frmClaveCaja.vAceptar Then
                If VerificaPass(frmClaveCaja.vUSUARIO, frmClaveCaja.vClave, vS) Then
                    pIngresoPass = True
                     Me.txtDescuento.Text = "0.00"
                Else
                    MsgBox vS, vbCritical, Pub_Titulo
                End If
            End If
            Else
                   Me.txtDescuento.Text = "0.00"
        End If
        
       
    Else

        Me.txtPago.Text = "0.00"
    End If

End Sub

Private Sub cmdDeno_Click(Index As Integer)

    If descuento Then
        If Not pIngresoPass Then
        
            frmClaveCaja.Show vbModal

            Dim vS As String

            If frmClaveCaja.vAceptar Then
                If VerificaPass(frmClaveCaja.vUSUARIO, frmClaveCaja.vClave, vS) Then
                    pIngresoPass = True

                    If Len(Trim(Me.txtDescuento.Text)) = 0 Then
                        Me.txtDescuento.Text = val(Me.cmdDeno(Index).Caption)
                    Else

                        If IsNumeric(Me.txtDescuento.Text) Then
                            Me.txtDescuento.Text = val(Me.txtDescuento.Text) + val(Me.cmdDeno(Index).Caption)
                        End If
                    End If

                Else
                    MsgBox vS, vbCritical, Pub_Titulo
                End If
            End If
        
        Else

            If Len(Trim(Me.txtDescuento.Text)) = 0 Then
                Me.txtDescuento.Text = val(Me.cmdDeno(Index).Caption)
            Else

                If IsNumeric(Me.txtDescuento.Text) Then
                    Me.txtDescuento.Text = val(Me.txtDescuento.Text) + val(Me.cmdDeno(Index).Caption)
                End If
            End If

        End If
     
    Else

        If Len(Trim(Me.txtPago.Text)) = 0 Then
            Me.txtPago.Text = val(Me.cmdDeno(Index).Caption)
        Else

            If IsNumeric(Me.txtPago.Text) Then
                Me.txtPago.Text = val(Me.txtPago.Text) + val(Me.cmdDeno(Index).Caption)
            End If
        End If
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    vAcepta = False
    pIngresoPass = False

    Dim vTop  As Integer

    Dim vLeft As Integer
    
    If Not gDELIVERY Then
        Me.Label3.Visible = False
        'Me.Label4.Visible = False
        Me.lblMovilidad.Visible = False
        'Me.lbltotal2.Visible = False
        ' Me.txtDescuento.Text = frmFacComanda.gDESCUENTO
    End If

    vTop = 120
    vLeft = 120

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpCargarFormasPago"

    If gDELIVERY Then
    
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodTran", adBigInt, adParamInput, 2, 2401)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DELIVERY", adBoolean, adParamInput, , True)
        
        Set oRStp = oCmdEjec.Execute
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodTran", adBigInt, adParamInput, 2, 2401)
        
        Set oRStp = oCmdEjec.Execute
    End If

    Dim f As Integer

    For f = 1 To oRStp.RecordCount
        Load Me.lblFormaPago(f)
        Load Me.txtImporte(f)
        Load Me.cboMoneda(f)
      
        Me.lblFormaPago(f).Top = vTop
        Me.lblFormaPago(f).Left = vLeft
        Me.txtImporte(f).Top = vTop
        Me.txtImporte(f).Left = Me.txtImporte(0).Left
        Me.txtImporte(f).Tag = f
        Me.cboMoneda(f).Top = vTop
        Me.cboMoneda(f).Left = Me.cboMoneda(0).Left

        vTop = vTop + Me.lblFormaPago(f).Height + 50
        
        Me.lblFormaPago(f).Visible = True
        Me.lblFormaPago(f).Caption = oRStp!formapago
        Me.lblFormaPago(f).Tag = oRStp!Codigo
        Me.txtImporte(f).Visible = True
        Me.cboMoneda(f).Visible = True
        Me.cboMoneda(f).AddItem "S/."
        Me.cboMoneda(f).AddItem "U$."
        Me.cboMoneda(f).Text = "S/."
        oRStp.MoveNext
    Next

    Dim oCTRL As Object

    Dim xpago As Double

    xpago = 0

    If gDELIVERY Then

        Dim ORSpr As ADODB.Recordset

        oCmdEjec.CommandText = "SP_DELIVERY_CARGAFORMASPAGO"
        LimpiaParametros oCmdEjec
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, frmDeliveryApp.lblSerie.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , frmDeliveryApp.lblNumero.Caption)
                
        Set ORSpr = oCmdEjec.Execute
        
        Do While Not ORSpr.EOF
            For i = 1 To Me.txtImporte.count - 1
                If ORSpr!idfp = Me.lblFormaPago(i).Tag Then
                    Me.txtImporte(i).Text = ORSpr!monto
                    Exit For
                End If
            Next
            ORSpr.MoveNext
        Loop
            
    Else

        If Not oRSfp Is Nothing Then
            If Not oRSfp.EOF Then oRSfp.MoveFirst
            If oRSfp.RecordCount <> 0 Then oRSfp.MoveFirst
     
            Do While Not oRSfp.EOF

                For Each oCTRL In Me.Controls

                    If TypeOf oCTRL Is label And oCTRL.Tag <> "" Then
            
                        If CInt(oCTRL.Tag) = oRSfp!idformapago Then
                            Me.txtImporte(oCTRL.Tag).Text = oRSfp!monto
                            xpago = xpago + oRSfp!monto
                            Exit For

                        End If
                    End If
            
                Next

                oRSfp.MoveNext
            Loop

        End If

    End If
  
    oRStp.Filter = ""
    oRStp.MoveFirst

    If gDELIVERY Then
    
        'AQUI CARGA LA TARIFA DE ENVIO DE DELIVERY
        oCmdEjec.CommandText = "SP_DELIVERY_CARGATARIFAPEDIDO"
        LimpiaParametros oCmdEjec

        Dim orsT As ADODB.Recordset

        Set orsT = oCmdEjec.Execute(, Array(LK_CODCIA, frmDeliveryApp.lblSerie.Caption, frmDeliveryApp.lblNumero.Caption))
        Me.lblTotal.Caption = Format(frmDeliveryApp.lblTot.Caption, "##0.#0")
        Me.lbltotal2.Caption = Format(frmDeliveryApp.lblTot.Caption + cMontoTarifa, "##0.#0")

        If Not orsT.EOF Then
        
            cMontoTarifa = orsT!tarifa
            Me.lblMovilidad.Caption = cMontoTarifa
            Me.txtDescuento.Text = orsT!descuento
            Me.txtPago.Text = IIf(IsNull(orsT!PAGO), 0, orsT!PAGO)
            Me.lblVuelto.Caption = orsT!VUELTO
            
        End If

    End If

    If oRSfp.RecordCount = 0 Then
   
        If gDELIVERY Then
            Me.txtImporte(1).Text = Format(frmDeliveryApp.lblTot.Caption, "##0.#0")
        Else
   
            If gMostrador Then
                Me.txtImporte(1).Text = Format(frmComanda2.lblTot.Caption, "##0.#0")
            Else
                Me.txtImporte(1).Text = Format(frmComanda.lblTot.Caption, "##0.#0")
            End If
        End If
    End If

    If gDELIVERY Then
       
       'CUANDO ES NULLO EL PAGO, CARGA EL TOTAL EN EL IMPORTE DE PAGO
       If IsNull(orsT!PAGO) Then
       Me.txtImporte(1).Text = val(Me.lblTotal.Caption) + val(Me.lblMovilidad.Caption)
       End If
       
    Else

        If frmFacComanda.lvDetalle.ListItems.count = 0 Then

            'EVALUAR SI ES DELIVERY
            If gMostrador Then
                Me.lblTotal.Caption = Format(frmComanda2.lblTot.Caption, "##0.#0")
            Else
                Me.lblTotal.Caption = Format(frmComanda.lblTot.Caption, "##0.#0")
            End If

        Else

            If xpago <> 0 Then
                Me.lblTotal.Caption = Format(xpago, "##0.#0")
            Else

                Me.lblTotal.Caption = Format(Me.lblTotal.Caption, "##0.#0")
            End If
            
            Me.txtDescuento.Text = frmFacComanda.gDESCUENTO

            If frmFacComanda.gPAGO = 0 Then
                Me.txtPago.Text = xpago
            Else
                Me.txtPago.Text = frmFacComanda.gPAGO
            End If
        End If
    End If

End Sub


Private Sub txtDescuento_Change()
If IsNumeric(Me.txtDescuento.Text) Then
    Me.lbltotal2.Caption = (val(Me.lblTotal.Caption) + val(Me.lblMovilidad.Caption)) - val(Me.txtDescuento.Text) ' val(Me.txtPago.Text) - val(Me.txtImporte(1).Text)
    If IsNumeric(Me.txtPago.Text) Then
        Me.lblVuelto.Caption = val(Me.txtPago.Text) - val(Me.lbltotal2.Caption)
    End If
End If
End Sub

Private Function VerificaPass(vUSUARIO As String, _
                              vClave As String, _
                              ByRef vMSN As String) As Boolean

    Dim orsPass As ADODB.Recordset

    Dim vtpass  As String, vPasa As Boolean

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpDevuelveClaveCaja"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, vUSUARIO)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CLAVE", adVarChar, adParamInput, 10, vClave)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MSN", adVarChar, adParamOutput, 200, 1)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)
    oCmdEjec.Execute

    'If Not orsPass.EOF Then vtpass = Trim(orsPass!Clave)
    vtpass = oCmdEjec.Parameters("@MSN").Value
    vPasa = oCmdEjec.Parameters("@PASA").Value
    vMSN = vtpass

    VerificaPass = vPasa
End Function

Private Sub txtDescuento_Click()

    If Not pIngresoPass Then
        frmClaveCaja.Show vbModal

        Dim vS As String

        If frmClaveCaja.vAceptar Then
            If VerificaPass(frmClaveCaja.vUSUARIO, frmClaveCaja.vClave, vS) Then
                pIngresoPass = True
            Else
                MsgBox vS, vbCritical, Pub_Titulo
            End If
        End If
    End If

End Sub

Private Sub txtDescuento_GotFocus()
descuento = True
End Sub

Private Sub txtImporte_Change(Index As Integer)

    Dim oCTRL  As Object

    Dim vTOTAL As Double

    vTOTAL = 0
    
    If Not oRStp.EOF Then
        oRStp.Filter = ""
        oRStp.MoveFirst

        Dim D As Integer

        For D = 1 To oRStp.RecordCount

            If Me.txtImporte(D).Text <> "" Then
                If IsNumeric(Me.txtImporte(D).Text) Then
                    vTOTAL = vTOTAL + Me.txtImporte(D).Text
                Else
                    vTOTAL = vTOTAL
                End If
                
            End If

        Next

        oRStp.Filter = ""
        oRStp.MoveFirst
    End If
    
    If Index = 1 Then
        Me.txtPago.Text = Me.txtImporte(Index).Text
    End If
    
    Me.lbltotal2.Caption = vTOTAL
    'Me.lbltotal2.Caption = vTotal + cMontoTarifa

End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
If NumerosyPunto(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtPago_Change()
If IsNumeric(Me.txtImporte(1).Text) Then
'    Me.lblVuelto.Caption = val(Me.txtPago.Text) - val(Me.txtImporte(1).Text)
Me.lblVuelto.Caption = val(Me.txtPago.Text) - val(Me.lbltotal2.Caption)
End If
End Sub

Private Sub txtPago_GotFocus()
descuento = False
End Sub
