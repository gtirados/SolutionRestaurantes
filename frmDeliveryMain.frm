VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmDeliveryMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deliverys"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16410
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
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   16410
   Begin MSComctlLib.ImageList ilDelivery 
      Left            =   7200
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeliveryMain.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbDelivery 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   16410
      _ExtentX        =   28945
      _ExtentY        =   635
      ButtonWidth     =   1720
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ilDelivery"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nuevo"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdNext 
      Enabled         =   0   'False
      Height          =   2880
      Left            =   13080
      Picture         =   "frmDeliveryMain.frx":039A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   3255
   End
   Begin VB.CommandButton cmdPrev 
      Enabled         =   0   'False
      Height          =   2880
      Left            =   120
      Picture         =   "frmDeliveryMain.frx":52A4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblPedido 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2880
      Index           =   0
      Left            =   6720
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "frmDeliveryMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vPagActual As Integer, vPagTotal As Integer

Private Sub CargarPedidos()
    Me.cmdPrev.Enabled = False

    If Me.lblPedido.count <> 1 Then

        For i = 1 To Me.lblPedido.count - 1
            Unload Me.lblPedido(i)
        Next

    End If

    Dim cont    As Integer

    Dim posleft As Integer, postop As Integer

    Dim fila    As Integer

    fila = 1
    posleft = 120
    postop = 480
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_DELIVERY_CARGAR"

    Dim orsP As ADODB.Recordset
    
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@usuario", adVarChar, adParamInput, 20, LK_CODUSU)
    
    Set orsP = oCmdEjec.Execute

    Dim strMSN As String

    For cont = 1 To orsP.RecordCount
        Load Me.lblPedido(cont)
  
        If fila <= 4 Then '1° fila
            If fila = 1 Then
                posleft = posleft + Me.cmdPrev.Width
            Else
                posleft = posleft + Me.lblPedido(fila - 1).Width
            End If

            ' MsgBox "james"
        ElseIf fila <= 9 Then '2° fila
          
            If fila = 5 Then
                posleft = 120
                postop = postop + Me.lblPedido(fila - 1).Height
            Else
                posleft = posleft + Me.lblPedido(fila - 1).Width
            End If

        ElseIf fila <= 13 Then

            If fila = 10 Then
                posleft = 120
                postop = postop + Me.lblPedido(fila - 1).Height
            Else
                posleft = posleft + Me.lblPedido(fila - 1).Width
            End If

        End If

        lblPedido(cont).Left = posleft
        lblPedido(cont).Top = postop
        
        lblPedido(cont).Visible = True
        lblPedido(cont).Tag = orsP!NumSer & "-" & orsP!NumFac

        strMSN = "Pedido Nro :" & CStr(orsP!NumFac) + vbCrLf & _
                 "Fecha: " & orsP!fecha + vbCrLf + _
                 "Hora Reg.: " & orsP!horaREG + vbCrLf + _
                 "Hora Salida: " & orsP!horasalida + vbCrLf + _
                 "Hora Lleg.: " & orsP!horallegada & vbCrLf & _
                 "Cliente: " & Trim(orsP!cliente) & vbCrLf & _
                 "Enviado a: " & Trim(orsP!direccion) & vbCrLf & _
                 "Referencia: " & Trim(orsP!ref) & vbCrLf & _
                 "Telefonos: " & Trim(orsP!fonos) & vbCrLf & _
                 "Repartidor: " & Trim(orsP!repartidor) & vbCrLf & _
                 "Observacion: " & Trim(orsP!OBS)
        
        lblPedido(cont).Caption = strMSN
        
        If orsP!ESTADO = "P" Then
            lblPedido(cont).BackColor = vbYellow
        ElseIf orsP!ESTADO = "E" Then
        
            lblPedido(cont).BackColor = vbMagenta
        Else
         
            lblPedido(cont).BackColor = &H8000000F
        End If

        If fila = 13 Then
            fila = 1
            postop = 480
            posleft = 120
        Else
            fila = fila + 1
        End If

        orsP.MoveNext
    Next
    
    Dim valor As Double

    valor = orsP.RecordCount / 13
    pos = InStr(Trim(str(valor)), ".")
    
    If pos <> 0 Then
        If pos = 1 Then
            ent = Left(CStr(valor), pos)
        Else
            ent = Left(CStr(valor), pos - 1)
        End If

    Else
        ent = Int(valor)
    End If
    
    If pos <> 0 Then
        pos2 = Right(Trim(str(valor)), Len(Trim(str(valor))) - pos)
    Else
        pos2 = 0
    End If

    If pos2 > 0 Then
        ent = ent + 1
    End If

    vPagTotal = ent

    If valor <> 0 Then vPagActual = 1

    If orsP.RecordCount > 13 Then: Me.cmdNext.Enabled = True

End Sub

Private Sub cmdNext_Click()

    Dim ini, fin, f As Integer

    If vPagActual = 1 Then
        ini = 1
        fin = 13
    ElseIf vPagActual = vPagTotal Then

        Exit Sub

    Else
        ini = (13 * vPagActual) - 12
        fin = 13 * vPagActual
    End If

    For f = ini To fin
        Me.lblPedido(f).Visible = False
    Next

    If vPagActual < vPagTotal Then
        vPagActual = vPagActual + 1

        If vPagActual = vPagTotal Then: Me.cmdNext.Enabled = False
    
        Me.cmdPrev.Enabled = True
    End If

End Sub

Private Sub cmdPrev_Click()

    Dim ini, fin, f As Integer

    If vPagActual = 2 Then
        ini = 1
        fin = 13
    ElseIf vPagActual = 1 Then

        Exit Sub

    Else
        FF = vPagActual - 1
        ini = (13 * FF) - 12
        fin = 13 * FF
    End If

    For f = ini To fin
        Me.lblPedido(f).Visible = True
    Next

    If vPagActual > 1 Then
        vPagActual = vPagActual - 1

        If vPagActual = 1 Then: Me.cmdPrev.Enabled = False
    
        Me.cmdNext.Enabled = True
    End If

End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
CargarPedidos
'Me.lblPedido(0).Caption = "Julio Mendoza" & vbCrLf & "Vinatea Reynoso 582 Urb. Santo Dominguito" & vbCrLf & "241368/950205108"
End Sub

Private Sub lblPedido_MouseDown(index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)

    If Button = vbRightButton Then

    ElseIf Button = vbLeftButton Then

        Dim MyMatriz() As String

        MyMatriz = Split(Me.lblPedido(index).Tag, "-")
        frmDeliveryApp.gEstaCargando = True
        frmDeliveryApp.vnumser = MyMatriz(0)
        frmDeliveryApp.vNumFac = MyMatriz(1)
        frmDeliveryApp.VNuevo = False
        frmDeliveryApp.vPrimero = False
        frmDeliveryApp.vEstado = "O"
        frmDeliveryApp.Show vbModal
        CargarPedidos
    End If

End Sub

Private Sub tbDelivery_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.index

        Case 1
             LEER_PAR_LLAVE
            If LK_FLAG_GRIFO <> "A" Then
                'If par_llave!PAR_FECHA_DIA <> LK_FECHA_DIA Then
                 '  MsgBox "!!!FECHA YA NO COINCIDE CON LA ACTUAL , OTRO USUARIO HA CERRADO EL DIA!!! SALGA Y REINICIE SU SISTEMA...", 48, Pub_Titulo

'                    End

                    'GoTo salirf
                'End If
            End If
        
            frmDeliveryApp.VNuevo = True
            frmDeliveryApp.vPrimero = True
            frmDeliveryApp.vEstado = "L"
            frmDeliveryApp.Show vbModal
            CargarPedidos
       
    End Select

End Sub
