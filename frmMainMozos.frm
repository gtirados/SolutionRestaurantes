VERSION 5.00
Object = "{798C3AED-5101-11D5-9278-0050FC0DD647}#93.1#0"; "CoolButton.ocx"
Begin VB.Form frmMainMozos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione mozo"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
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
   ScaleHeight     =   5505
   ScaleWidth      =   8640
   Begin CoolButton.CoolCommand cmdMozo 
      Height          =   1335
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2355
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      checkcaption    =   "Command"
      BackColor1      =   16744576
      BackColor2      =   16761024
      Icon            =   "frmMainMozos.frx":0000
      IconSize        =   3
      IconAlign       =   3
      IconWidth       =   52
      IconHeight      =   52
      TextAlign       =   4
      IcoDistFromEdge =   8
   End
   Begin CoolButton.CoolCommand cmdPrev 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2355
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      checkcaption    =   ""
      Caption         =   ""
      BackColor1      =   16744576
      BackColor2      =   16761024
      Icon            =   "frmMainMozos.frx":0786
      IconSize        =   3
      IconWidth       =   75
      IconHeight      =   75
      TextAlign       =   4
      IcoDistFromEdge =   0
   End
   Begin CoolButton.CoolCommand cmdNext 
      Height          =   1335
      Left            =   6900
      TabIndex        =   2
      Top             =   4125
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2355
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      checkcaption    =   ""
      Caption         =   ""
      BackColor1      =   16744576
      BackColor2      =   16761024
      Icon            =   "frmMainMozos.frx":12B7
      IconSize        =   3
      IconWidth       =   75
      IconHeight      =   75
      TextAlign       =   4
      IcoDistFromEdge =   0
   End
End
Attribute VB_Name = "frmMainMozos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vPagActual As Integer, vPagTotal As Integer
Public gCodMozo As String
Public gMozo As String


Private Sub cmdMozo_Click(Index As Integer)
'PEDIR CLAVE
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SP_REQUIERE_PASSWORD"
Set oRSfp = oCmdEjec.Execute(, LK_CODUSU)

If Not oRSfp.EOF Then
    If oRSfp!PASSW = "A" Then

        frmcomandamozomesa.gModificaMozo = True
        frmcomandamozomesa.gCodMozo = cmdMozo(Index).Tag
        frmcomandamozomesa.gMozo = cmdMozo(Index).Caption
        frmcomandamozomesa.Show vbModal
        
        If frmcomandamozomesa.gEntro Then
            gCodMozo = frmcomandamozomesa.gCodMozo ' Me.cmdMozo(Index).Tag
            gMozo = frmcomandamozomesa.gMozo ' Me.cmdMozo(Index).Caption
            frmMainMesas.Show
        End If
    Else
        gCodMozo = Me.cmdMozo(Index).Tag
        gMozo = Me.cmdMozo(Index).Caption
        frmMainMesas.Show
    End If
End If
End Sub

Private Sub cmdNext_Click()
Dim ini, fin, f As Integer

    If vPagActual = 1 Then
        ini = 1
        fin = 18
    ElseIf vPagActual = vPagTotal Then
        Exit Sub
    Else
        ini = (18 * vPagActual) - 17
        fin = 18 * vPagActual
    End If

    For f = ini To fin
        Me.cmdMozo(f).Visible = False
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
        fin = 18
    ElseIf vPagActual = 1 Then

        Exit Sub

    Else
        FF = vPagActual - 1
        ini = (18 * FF) - 17
        fin = 18 * FF
    End If

    For f = ini To fin
        Me.cmdMozo(f).Visible = True
    Next

    If vPagActual > 1 Then
        vPagActual = vPagActual - 1

        If vPagActual = 1 Then: Me.cmdPrev.Enabled = False
    
        Me.cmdNext.Enabled = True
    End If
End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
CargarMozos
End Sub

Private Sub CargarMozos()

    Dim oRsMozos As ADODB.Recordset

    Dim itemM    As ListItem

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpListarMozos"
    Set oRsMozos = oCmdEjec.Execute(, LK_CODCIA)
    
    Dim cont    As Integer

    Dim posleft As Integer, postop As Integer

    Dim fila    As Integer

    fila = 1
    posleft = 120
    postop = 120

    For cont = 1 To oRsMozos.RecordCount

        'Load Me.cmdPrev(cont)
        Load Me.cmdMozo(cont)
 
        If fila <= 4 Then '1° fila
            If fila = 1 Then
                posleft = posleft + Me.cmdPrev.Width
            Else
                posleft = posleft + Me.cmdMozo(fila - 1).Width
            End If

            ' MsgBox "james"
        ElseIf fila <= 9 Then '2° fila

            If fila = 5 Then
                posleft = 120
                postop = postop + Me.cmdMozo(fila - 1).Height
            Else
                posleft = posleft + Me.cmdMozo(fila - 1).Width
            End If

        ElseIf fila <= 14 Then

            If fila = 10 Then
                posleft = 120
                postop = postop + Me.cmdMozo(fila - 1).Height
            Else
                posleft = posleft + Me.cmdMozo(fila - 1).Width
            End If

        ElseIf fila <= 18 Then

            If fila = 15 Then
                posleft = 120
                postop = postop + Me.cmdMozo(fila - 1).Height
            Else
                posleft = posleft + Me.cmdMozo(fila - 1).Width
            End If
        End If

        Me.cmdMozo(cont).Left = posleft
        cmdMozo(cont).Top = postop

        cmdMozo(cont).Visible = True
        cmdMozo(cont).Tag = oRsMozos!Codigo
        cmdMozo(cont).Caption = Trim(oRsMozos!mozo)

        If fila = 18 Then
            fila = 1
            postop = 120
            posleft = 120
        Else
            fila = fila + 1
        End If

        oRsMozos.MoveNext
    Next

    Dim valor As Double

    valor = oRsMozos.RecordCount / 18
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

    If oRsMozos.RecordCount > 18 Then: Me.cmdNext.Enabled = True

End Sub

''Private Sub CargarPedidos()
''    Me.cmdPrev.Enabled = False
''
''    If Me.lblPedido.count <> 1 Then
''
''        For i = 1 To Me.lblPedido.count - 1
''            Unload Me.lblPedido(i)
''        Next
''
''    End If
''
''    Dim cont    As Integer
''
''    Dim posleft As Integer, postop As Integer
''
''    Dim fila    As Integer
''
''    fila = 1
''    posleft = 120
''    postop = 480
''
''    LimpiaParametros oCmdEjec
''    oCmdEjec.CommandText = "SP_DELIVERY_CARGAR"
''
''    Dim oRSp As ADODB.Recordset
''
''    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
''    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
''
''    Set oRSp = oCmdEjec.Execute
''
''    Dim strMSN As String
''
''    For cont = 1 To oRSp.RecordCount
''        Load Me.lblPedido(cont)
''
''        If fila <= 4 Then '1° fila
''            If fila = 1 Then
''                posleft = posleft + Me.cmdPrev.Width
''            Else
''                posleft = posleft + Me.lblPedido(fila - 1).Width
''            End If
''
''            ' MsgBox "james"
''        ElseIf fila <= 9 Then '2° fila
''
''            If fila = 5 Then
''                posleft = 120
''                postop = postop + Me.lblPedido(fila - 1).Height
''            Else
''                posleft = posleft + Me.lblPedido(fila - 1).Width
''            End If
''
''        ElseIf fila <= 13 Then
''
''            If fila = 10 Then
''                posleft = 120
''                postop = postop + Me.lblPedido(fila - 1).Height
''            Else
''                posleft = posleft + Me.lblPedido(fila - 1).Width
''            End If
''
''        End If
''
''        lblPedido(cont).Left = posleft
''        lblPedido(cont).Top = postop
''
''        lblPedido(cont).Visible = True
''        lblPedido(cont).Tag = oRSp!NumSer & "-" & oRSp!NumFac
''
''        strMSN = "Pedido Nro :" & CStr(oRSp!NumFac) + vbCrLf + "Fecha: " & oRSp!fecha + vbCrLf + "Hora Reg.: " & _
''                oRSp!horaREG + vbCrLf + "Hora Salida: " & oRSp!horasalida + vbCrLf + "Hora Lleg.: " & _
''                oRSp!horallegada & vbCrLf & "Cliente: " & Trim(oRSp!cliente) & vbCrLf & "Enviado a: " & _
''                Trim(oRSp!direccion) & vbCrLf & "Telefonos: " & Trim(oRSp!fonos) & vbCrLf & _
''                "Repartidor: " & Trim(oRSp!repartidor)
''
''        lblPedido(cont).Caption = strMSN
''
''        If oRSp!ESTADO = "P" Then
''            lblPedido(cont).BackColor = vbYellow
''        ElseIf oRSp!ESTADO = "E" Then
''
''            lblPedido(cont).BackColor = vbMagenta
''        Else
''
''            lblPedido(cont).BackColor = &H8000000F
''        End If
''
''        If fila = 13 Then
''            fila = 1
''            postop = 480
''            posleft = 120
''        Else
''            fila = fila + 1
''        End If
''
''        oRSp.MoveNext
''    Next
''
''    Dim valor As Double
''
''    valor = oRSp.RecordCount / 13
''    pos = InStr(Trim(str(valor)), ".")
''
''    If pos <> 0 Then
''        If pos = 1 Then
''            ent = Left(CStr(valor), pos)
''        Else
''            ent = Left(CStr(valor), pos - 1)
''        End If
''
''    Else
''        ent = Int(valor)
''    End If
''
''    If pos <> 0 Then
''        pos2 = Right(Trim(str(valor)), Len(Trim(str(valor))) - pos)
''    Else
''        pos2 = 0
''    End If
''
''    If pos2 > 0 Then
''        ent = ent + 1
''    End If
''
''    vPagTotal = ent
''
''    If valor <> 0 Then vPagActual = 1
''
''    If oRSp.RecordCount > 13 Then: Me.cmdNext.Enabled = True
''
''End Sub
