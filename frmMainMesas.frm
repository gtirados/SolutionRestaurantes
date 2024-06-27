VERSION 5.00
Object = "{798C3AED-5101-11D5-9278-0050FC0DD647}#93.1#0"; "CoolButton.ocx"
Begin VB.Form frmMainMesas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mesas"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13365
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
   ScaleHeight     =   8040
   ScaleWidth      =   13365
   Begin VB.Frame Frame2 
      Height          =   8055
      Left            =   2920
      TabIndex        =   1
      Top             =   -50
      Width           =   10410
      Begin CoolButton.CoolCommand cmdMesa 
         Height          =   1215
         Index           =   0
         Left            =   4800
         TabIndex        =   6
         Top             =   4560
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   2143
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
         HighColor1      =   14737632
         HighColor2      =   8421504
         Caption         =   ""
         BackColor1      =   8421504
         BackColor2      =   14737632
         Icon            =   "frmMainMesas.frx":0000
         IconSize        =   3
         IconAlign       =   3
         IconWidth       =   40
         IconHeight      =   40
         TextAlign       =   4
      End
      Begin CoolButton.CoolCommand cmdPrev 
         Height          =   1215
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   2143
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
         Icon            =   "frmMainMesas.frx":0849
         IconSize        =   3
         IconWidth       =   72
         IconHeight      =   72
      End
      Begin CoolButton.CoolCommand cmdNext 
         Height          =   1215
         Left            =   8595
         TabIndex        =   5
         Top             =   6795
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   2143
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
         Icon            =   "frmMainMesas.frx":137A
         IconSize        =   3
         IconWidth       =   72
         IconHeight      =   72
      End
      Begin VB.Label lblZona 
         Alignment       =   2  'Center
         Caption         =   "Zona"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   9855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   -50
      Width           =   2895
      Begin CoolButton.CoolCommand cmdZona 
         Height          =   1095
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1931
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
         Icon            =   "frmMainMesas.frx":1E9B
         IconSize        =   3
         IconAlign       =   1
         IconWidth       =   64
         IconHeight      =   64
         TextAlign       =   2
         IcoDistFromEdge =   0
      End
      Begin CoolButton.CoolCommand cmdPrevZona 
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1931
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   1
         checkcaption    =   "Arriba"
         HighlightStyle  =   1
         Caption         =   "Arriba"
         BackColor1      =   16744576
         BackColor2      =   16761024
         IconSize        =   3
         IconAlign       =   1
         IconWidth       =   64
         IconHeight      =   64
         IcoDistFromEdge =   0
      End
      Begin CoolButton.CoolCommand cmdNextZona 
         Height          =   1095
         Left            =   120
         TabIndex        =   8
         Top             =   6720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1931
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   1
         checkcaption    =   "Abajo"
         HighlightStyle  =   1
         Caption         =   "Abajo"
         BackColor1      =   16744576
         BackColor2      =   16761024
         IconSize        =   3
         IconAlign       =   1
         IconWidth       =   64
         IconHeight      =   64
         IcoDistFromEdge =   0
      End
   End
End
Attribute VB_Name = "frmMainMesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vPagActual As Integer, vPagTotal As Integer
Private vPagActualZ As Integer, vPagTotalZ As Integer
Private vIDZona As Integer

Private Sub cmdMesa_Click(Index As Integer)
   If Split(Me.cmdMesa(Index).Tag, "|")(1) = "L" Then 'MESA LIBRE
        frmComanda.gDefecto = True
        frmComanda.vPrimero = True
        frmComanda.VNuevo = True
        frmComanda.gMozo = frmMainMozos.gCodMozo
        frmComanda.vMesa = Split(Me.cmdMesa(Index).Tag, "|")(0)
        frmComanda.lblmesa.Caption = Me.cmdMesa(Index).Caption
        frmComanda.lblMozo.Caption = frmMainMozos.gMozo
        frmComanda.Show vbModal
    Else 'MESA OCUPADA
        frmComanda.VNuevo = False
        frmComanda.gDefecto = True
        frmComanda.vPrimero = False
        frmComanda.vEstado = "O"
        frmComanda.vMesa = Split(Me.cmdMesa(Index).Tag, "|")(0)
        frmComanda.vCodZona = vIDZona
        'frmcomanda.vCodPlato = Me.lblNomMesa(Index).Tag
        frmComanda.lblmesa.Caption = Me.cmdMesa(Index).Caption
        frmComanda.gMozo = frmMainMozos.gCodMozo 'nuevo
                
        frmComanda.Show vbModal
    End If
                
    'EVALUAR EL ESTADO
                
    ''
End Sub

Private Sub cmdNext_Click()

    Dim ini, fin, f As Integer

    If vPagActual = 1 Then
        ini = 1
        fin = 34
    ElseIf vPagActual = vPagTotal Then
        Exit Sub
    Else
        ini = (34 * vPagActual) - 33
        fin = 34 * vPagActual
    End If

    For f = ini To fin
        Me.cmdMesa(f).Visible = False
    Next

    If vPagActual < vPagTotal Then
        vPagActual = vPagActual + 1

        If vPagActual = vPagTotal Then: Me.cmdNext.Enabled = False
    
        Me.cmdPrev.Enabled = True
    End If
End Sub





Private Sub cmdNextZona_Click()

    Dim ini, fin, f As Integer

    If vPagActualZ = 1 Then
        ini = 1
        fin = 5
    ElseIf vPagActualZ = vPagTotalZ Then
        Exit Sub
    Else
        ini = (5 * vPagActualZ) - 4
        fin = 5 * vPagActualZ
    End If

    For f = ini To fin
        Me.cmdZona(f).Visible = False
    Next

    If vPagActualZ < vPagTotalZ Then
        vPagActualZ = vPagActualZ + 1

        If vPagActualZ = vPagTotalZ Then: Me.cmdNextZona.Enabled = False
    
        Me.cmdPrevZona.Enabled = True
    End If
End Sub

Private Sub cmdPrev_Click()

    Dim ini, fin, f As Integer

    If vPagActual = 2 Then
        ini = 1
        fin = 34
    ElseIf vPagActual = 1 Then

        Exit Sub

    Else
        FF = vPagActual - 1
        ini = (34 * FF) - 33
        fin = 34 * FF
    End If

    For f = ini To fin
        Me.cmdMesa(f).Visible = True
    Next

    If vPagActual > 1 Then
        vPagActual = vPagActual - 1

        If vPagActual = 1 Then: Me.cmdPrev.Enabled = False
    
        Me.cmdNext.Enabled = True
    End If
End Sub

Private Sub cmdPrevZona_Click()

    Dim ini, fin, f As Integer

    If vPagActualZ = 2 Then
        ini = 1
        fin = 5
    ElseIf vPagActualZ = 1 Then

        Exit Sub

    Else
        FF = vPagActual - 1
        ini = (5 * FF) - 4
        fin = 5 * FF
    End If

    For f = ini To fin
        Me.cmdZona(f).Visible = True
    Next

    If vPagActualZ > 1 Then
        vPagActualZ = vPagActualZ - 1

        If vPagActualZ = 1 Then: Me.cmdPrevZona.Enabled = False
    
        Me.cmdNextZona.Enabled = True
    End If
End Sub

Private Sub cmdZona_Click(Index As Integer)
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPCARGARMESAS_LIBRES"

    Dim ORSmESAS As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODZON", adInteger, adParamInput, , Me.cmdZona(Index).Tag)
                
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODMOZO", adInteger, adParamInput, , frmMainMozos.gCodMozo)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
                
    Set ORSmESAS = oCmdEjec.Execute
    Me.lblZona.Caption = Me.cmdZona(Index).Caption
    vIDZona = Me.cmdZona(Index).Tag

    Dim cont    As Integer

    Dim posleft As Integer, postop As Integer

    Dim fila    As Integer

    fila = 1
    posleft = 120
    postop = 720

    For cont = 1 To Me.cmdMesa.count - 1
        Unload Me.cmdMesa(cont)
   
    Next

    For cont = 1 To ORSmESAS.RecordCount

        Load Me.cmdMesa(cont)
 
        If fila <= 5 Then '1° fila
            If fila = 1 Then
                posleft = posleft + Me.cmdPrev.Width
            Else
                posleft = posleft + Me.cmdMesa(fila - 1).Width
            End If

            ' MsgBox "james"
        ElseIf fila <= 11 Then '2° fila

            If fila = 6 Then
                posleft = 120
                postop = postop + Me.cmdMesa(fila - 1).Height
            Else
                posleft = posleft + Me.cmdMesa(fila - 1).Width
            End If

        ElseIf fila <= 17 Then

            If fila = 12 Then
                posleft = 120
                postop = postop + Me.cmdMesa(fila - 1).Height
            Else
                posleft = posleft + Me.cmdMesa(fila - 1).Width
            End If

        ElseIf fila <= 23 Then

            If fila = 18 Then
                posleft = 120
                postop = postop + Me.cmdMesa(fila - 1).Height
            Else
                posleft = posleft + Me.cmdMesa(fila - 1).Width
            End If

        ElseIf fila <= 29 Then

            If fila = 24 Then
                posleft = 120
                postop = postop + Me.cmdMesa(fila - 1).Height
            Else
                posleft = posleft + Me.cmdMesa(fila - 1).Width
            End If

        ElseIf fila <= 35 Then

            If fila = 30 Then
                posleft = 120
                postop = postop + Me.cmdMesa(fila - 1).Height
            Else
                posleft = posleft + Me.cmdMesa(fila - 1).Width
            End If
        End If

        Me.cmdMesa(cont).Left = posleft
        cmdMesa(cont).Top = postop

        cmdMesa(cont).Visible = True
        cmdMesa(cont).Tag = ORSmESAS!codmesa & "|" & ORSmESAS!ESTADO

        If ORSmESAS!ESTADO = "O" Then
            'cmdMesa(cont).BackColor = vbRed
            cmdMesa(cont).BackColor1 = &HC0&
            cmdMesa(cont).BackColor2 = &HFF&
            cmdMesa(cont).HighColor1 = &HFF&
            cmdMesa(cont).HighColor2 = &HC0&
        ElseIf ORSmESAS!ESTADO = "E" Then
            'cmdMesa(cont).BackColor = vbYellow
            cmdMesa(cont).BackColor1 = &HFFFF&
            cmdMesa(cont).BackColor2 = &H80FFFF
            cmdMesa(cont).HighColor1 = &H80FFFF
            cmdMesa(cont).HighColor2 = &HFFFF&
        End If

        cmdMesa(cont).Caption = Trim(ORSmESAS!mesa) & vbCrLf & ORSmESAS!CLIENTE
        cmdMesa(cont).Enabled = True

        '  MsgBox "2"
        If fila = 34 Then
            fila = 1
            postop = 720
            posleft = 120
        Else
            fila = fila + 1
        End If

        ORSmESAS.MoveNext
    Next

    Dim valor As Double

    valor = ORSmESAS.RecordCount / 13
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

    If ORSmESAS.RecordCount > 18 Then: Me.cmdNext.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
CargarZonas
End Sub

Private Sub CargarZonas()

    Dim oRsZonas As ADODB.Recordset

    LimpiaParametros oCmdEjec
    
    oCmdEjec.CommandText = "SpListarZonas"

    Set oRsZonas = oCmdEjec.Execute(, LK_CODCIA)

   
       '        cmdZona(i).Visible = True
    '        Me.cmdZona(i).Top = postop
    '        Me.cmdZona(i).Left = Me.cmdZona(0).Left
    '        Me.cmdZona(i).Caption = oRsZonas!denomina
    '        Me.cmdZona(i).Tag = oRsZonas!Codigo
    '        postop = postop + Me.cmdZona(i).Height
    '        oRsZonas.MoveNext
    
    
    

    Dim cont    As Integer

    Dim postop As Integer

    Dim fila    As Integer

    fila = 1
    postop = 1320

    For cont = 1 To Me.cmdZona.count - 1
        Unload Me.cmdZona(cont)
    Next

    For cont = 1 To oRsZonas.RecordCount

        Load Me.cmdZona(cont)



        Me.cmdZona(cont).Left = Me.cmdPrevZona.Left
        cmdZona(cont).Top = postop
        cmdZona(cont).Visible = True
        cmdZona(cont).Tag = oRsZonas!Codigo
        cmdZona(cont).Caption = Trim(oRsZonas!denomina)
        cmdZona(cont).Enabled = True
 postop = postop + Me.cmdZona(fila - 1).Height
        '  MsgBox "2"
        If fila = 5 Then
            fila = 1
            postop = 1320
        Else
            fila = fila + 1
        End If

        oRsZonas.MoveNext
    Next

    Dim valor As Double

    valor = oRsZonas.RecordCount / 5
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

    vPagTotalZ = ent

    If valor <> 0 Then vPagActualZ = 1

    If oRsZonas.RecordCount > 5 Then: Me.cmdNextZona.Enabled = True
    
    '    Dim oRsZonas As ADODB.Recordset
    '
    '    LimpiaParametros oCmdEjec
    '    oCmdEjec.CommandText = "SpListarZonas"
    '
    '    Set oRsZonas = oCmdEjec.Execute(, LK_CODCIA)
    '
    '    Dim postop As Integer
    '
    '    postop = 1320
    '
    '    Dim i As Integer
    '
    '    For i = 1 To oRsZonas.RecordCount
    '        Load Me.cmdZona(i)
    '
    '        cmdZona(i).Visible = True
    '        Me.cmdZona(i).Top = postop
    '        Me.cmdZona(i).Left = Me.cmdZona(0).Left
    '        Me.cmdZona(i).Caption = oRsZonas!denomina
    '        Me.cmdZona(i).Tag = oRsZonas!Codigo
    '        postop = postop + Me.cmdZona(i).Height
    '        oRsZonas.MoveNext
    '    Next

End Sub

