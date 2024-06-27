VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmimport 
   BackColor       =   &H00FAEFDA&
   Caption         =   "Datos para Importar"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "frmimport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Importar Datos de Registro de Venta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4365
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5790
      Begin VB.TextBox tventa 
         Height          =   285
         Left            =   3840
         TabIndex        =   14
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Igv 
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox vventa 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtanulados 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtcodibol 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtarchi 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Text            =   "A:\"
         Top             =   600
         Width           =   3375
      End
      Begin ComctlLib.ProgressBar pb 
         Height          =   210
         Left            =   270
         TabIndex        =   3
         Top             =   3885
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   370
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton cmdimport 
         Caption         =   "Importar Datos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3615
         TabIndex        =   2
         Top             =   3750
         Width           =   2055
      End
      Begin VB.Label lcodanul 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   2520
         Width           =   3855
      End
      Begin VB.Label lcodclie 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label lblvar 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Cta. Total  Venta :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   3840
         TabIndex        =   17
         Top             =   3000
         Width           =   1485
      End
      Begin VB.Label lblvar 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Cta.IGV.:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   1920
         TabIndex        =   16
         Top             =   3000
         Width           =   720
      End
      Begin VB.Label lblvar 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Cta. Valor Venta :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   3000
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Para los anulados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   2100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Para las Boletas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   1965
      End
      Begin VB.Label lblcia 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   720
         TabIndex        =   7
         Top             =   1200
         Width           =   4785
      End
      Begin VB.Label lblvar 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblvar 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   645
      End
      Begin VB.Label lblvar 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Origen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Solution - Gestión Contable"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   0
      TabIndex        =   20
      Top             =   4395
      Width           =   5820
   End
End
Attribute VB_Name = "frmimport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xl As Object
Dim MOVI  As rdoQuery
Dim regmovi As rdoResultset
Dim CLI_MOVI  As rdoQuery
Dim climovi As rdoResultset
Dim CLI_SAL  As rdoQuery
Dim clisal As rdoResultset
Dim PSMOV_VOU As rdoQuery
Dim VOU_MOV As rdoResultset




Private Sub cmdimport_Click()
Dim QCODCLIE  As Currency
If Val(txtcodibol.Text) = 0 Then
' MsgBox "Ingresar codigo de las boletas"
' Exit Sub
End If

If Val(txtanulados.Text) = 0 Then
' MsgBox "Ingresar codigo de los anulados"
' Exit Sub
End If

SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
PUB_CUENTA = Trim(vventa.Text)
PUB_CODCIA = LK_CODCIA
LEER_COM_LLAVE
If com_llave.EOF Then
  MsgBox "Cuenta No Existe", 48, Pub_Titulo
   Exit Sub
Else
  If com_llave!com_nivel <> 3 Then
    MsgBox "Cuenta No es Analitica", 48, Pub_Titulo
    Exit Sub
  End If
End If

SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
PUB_CUENTA = Trim(Igv.Text)
LEER_COM_LLAVE
If com_llave.EOF Then
  MsgBox "Cuenta No Existe", 48, Pub_Titulo
   Exit Sub
Else
  If com_llave!com_nivel <> 3 Then
    MsgBox "Cuenta No es Analitica", 48, Pub_Titulo
    Exit Sub
  End If
End If
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
PUB_CUENTA = Trim(tventa.Text)
LEER_COM_LLAVE
If com_llave.EOF Then
  MsgBox "Cuenta No Existe", 48, Pub_Titulo
   Exit Sub
Else
  If com_llave!com_nivel <> 3 Then
    MsgBox "Cuenta No es Analitica", 48, Pub_Titulo
    Exit Sub
  End If
End If


regmovi.Requery
F1 = 1
GoSub WEXCEL

PSMOV_VOU.rdoParameters(0) = LK_CODCIA
PSMOV_VOU.rdoParameters(1) = LK_FECHA_COP1
PSMOV_VOU.rdoParameters(2) = LK_FECHA_COP2
PSMOV_VOU.rdoParameters(3) = LK_NRO_MES
PSMOV_VOU.rdoParameters(4) = 2
VOU_MOV.Requery
If VOU_MOV.EOF Then
   ws_nro_voucher = 0
Else
   ws_nro_voucher = VOU_MOV!MOV_NRO_VOUCHER
End If
Do Until Val(xl.Cells(F1, 2)) = 0
 ws_nro_voucher = ws_nro_voucher + 1
 PUB_CP = "C"
 If Trim(xl.Cells(F1, 11)) = "E" Then
'  QCODCLIE = Val(txtanulados.Text)
' GoTo dale
 End If
 QCODCLIE = 0 ' Val(txtcodibol.Text)
 If Val(xl.Cells(F1, 2)) = 1 Then
    SQ_OPER = 4
    pu_cp = "C"
    PUB_RUC = Trim(xl.Cells(F1, 9))
    pu_codcia = LK_CODCIA
    LEER_CLI_LLAVE
    If cli_ruc.EOF Then
      GoSub AGREGA_CLI
    Else
      QCODCLIE = Val(cli_ruc!CLI_CODCLIE)
    End If
 End If
 If Val(xl.Cells(F1, 2)) = 3 Then
    SQ_OPER = 1
    pu_cp = "C"
    pu_codclie = Trim(xl.Cells(F1, 9))
    pu_codcia = LK_CODCIA
    LEER_CLI_LLAVE
    If cli_llave.EOF Then
      GoSub AGREGA_CLI
    Else
      QCODCLIE = Val(cli_llave!CLI_CODCLIE)
    End If
 End If
 
dale:
 WS_NRO_MOV = 0
 PUB_CUENTA = Trim(tventa.Text)
 PUB_IMPORTE = xl.Cells(F1, 8)
 If QCODCLIE = Val(txtanulados.Text) Then PUB_IMPORTE = 0
 w_dh = "D"
 GoSub REGISTRA
 PUB_CP = " "
 QCODCLIE = 0
 WS_NRO_MOV = 1
 PUB_CUENTA = Trim(vventa.Text)
 PUB_IMPORTE = xl.Cells(F1, 6)
 If QCODCLIE = Val(txtanulados.Text) Then PUB_IMPORTE = 0
 w_dh = "H"
 GoSub REGISTRA
 WS_NRO_MOV = 2
 PUB_CUENTA = Trim(Igv.Text)
 PUB_IMPORTE = xl.Cells(F1, 7)
 If QCODCLIE = Val(txtanulados.Text) Then PUB_IMPORTE = 0
 w_dh = "H"
 GoSub REGISTRA
 
 F1 = F1 + 1
Loop
MsgBox "Proceso Terminado", 48, Pub_Titulo
xl.Application.Visible = True
Set xl = Nothing



Exit Sub

REGISTRA:
   regmovi.AddNew
   regmovi!MOV_NRO_MOV = WS_NRO_MOV
   regmovi!MOV_CODCIA = LK_CODCIA
   regmovi!MOV_NRO_VOUCHER = ws_nro_voucher
   regmovi!MOV_TIPMOV = 2
   regmovi!MOV_FECHA = LK_FECHA_COP1
   regmovi!MOV_GLOSA = "Por las Ventas del mes"
   regmovi!MOV_MONEDA = Trim(xl.Cells(F1, 5))
   regmovi!MOV_CODCTA = PUB_CUENTA
   regmovi!MOV_DH = w_dh
   regmovi!MOV_IMPORTE = PUB_IMPORTE
   regmovi!MOV_SUNAT = Val(xl.Cells(F1, 2))
   regmovi!MOV_serie = Val(xl.Cells(F1, 3))
   regmovi!MOV_numfac = Val(xl.Cells(F1, 4))

   regmovi!MOV_codclie = QCODCLIE
   
   regmovi!MOV_CP = PUB_CP
   
   regmovi!MOV_FBG = Format((xl.Cells(F1, 2)), "00")
   regmovi!MOV_MARCA = "X"
   regmovi!MOV_DETALLE = "Por las Ventas del mes"
   regmovi!MOV_FBG_C = " "
   regmovi!MOV_numfac_c = 0
   regmovi!MOV_serie_c = 0

   regmovi!MOV_nro_MES = LK_NRO_MES
   regmovi!MOV_fecha_EMI = (xl.Cells(F1, 1))
   regmovi!MOV_PLANTILLA = 1
   regmovi!MOV_FLAG_TC = ""
   regmovi!MOV_TIPO_CAMBIO = 0
   regmovi!MOV_FLAG_DES = " "
   regmovi!MOV_CODUSU = LK_CODUSU
   regmovi.Update
Return


AGREGA_CLI:
QCODCLIE = GENERA_CODI()
       climovi.AddNew
       climovi!CLI_CP = "C"
       climovi!CLI_CODCLIE = QCODCLIE
       climovi!cli_SALDO = 0
       climovi!CLI_DET_TOT = "D"
       climovi!CLI_MONEDA = "S"
          SQ_OPER = 5
          pu_codclie = QCODCLIE
          pu_cp = "C"
          pu_codcia = LK_CODCIA
          LEER_CLI_LLAVE
          If cls_llave.EOF Then
                cls_llave.AddNew
                cls_llave!CLS_CODCIA = LK_CODCIA
                cls_llave!CLS_CODCLIE = QCODCLIE
                cls_llave!CLS_CP = "C"
                cls_llave!CLS_DEB00 = 0
                cls_llave!CLS_HAB00 = 0
                cls_llave!CLS_DEB01 = 0
                cls_llave!CLS_HAB01 = 0
                cls_llave!CLS_DEB02 = 0
                cls_llave!CLS_HAB02 = 0
                cls_llave!CLS_DEB03 = 0
                cls_llave!CLS_HAB03 = 0
                cls_llave!CLS_DEB04 = 0
                cls_llave!CLS_HAB04 = 0
                cls_llave!CLS_DEB05 = 0
                cls_llave!CLS_HAB05 = 0
                cls_llave!CLS_DEB06 = 0
                cls_llave!CLS_HAB06 = 0
                cls_llave!CLS_DEB07 = 0
                cls_llave!CLS_HAB07 = 0
                cls_llave!CLS_DEB08 = 0
                cls_llave!CLS_HAB08 = 0
                cls_llave!CLS_DEB09 = 0
                cls_llave!CLS_HAB09 = 0
                cls_llave!CLS_DEB10 = 0
                cls_llave!CLS_HAB10 = 0
                cls_llave!CLS_DEB11 = 0
                cls_llave!CLS_HAB11 = 0
                cls_llave!CLS_DEB12 = 0
                cls_llave!CLS_HAB12 = 0
                cls_llave.Update
        End If

    climovi!CLI_CODCIA = LK_CODCIA
    climovi!CLI_NOMBRE_ESPOSO = Trim(xl.Cells(F1, 10))
    climovi!CLI_NOMBRE_ESPOSA = ""
    climovi!CLI_NOMBRE_EMPRESA = ""
    climovi!CLI_123 = 1
    climovi!cli_nombre = Trim(xl.Cells(F1, 10))
    climovi!CLI_CASA_DIREC = ""
    climovi!CLI_CASA_NUM = 0
    climovi!CLI_CASA_ZONA = 1
    climovi!CLI_LUGAR_CASA = 1
    climovi!CLI_LUGAR_TRAB = 1
    climovi!CLI_CASA_SUBZONA = 1
    climovi!CLI_ZONA_NEW = 1
    climovi!CLI_TRAB_DIREC = ""
    climovi!CLI_TRAB_NUM = 0
    climovi!cli_TRAB_ZONA = 1
    climovi!cli_TRAB_SUBZONA = 1
    climovi!cli_ruc_esposo = Trim(xl.Cells(F1, 9))
    climovi!cli_ruc_esposA = ""
    climovi!CLI_RUC_EMPRESA = ""
    climovi!CLI_CASA1 = ""
    climovi!CLI_CASA2 = ""
    climovi!CLI_REGPUB1 = ""
    climovi!CLI_REGPUB2 = ""
    climovi!CLI_AUTOAVALUO = ""
    climovi!CLI_AUTO1 = ""
    climovi!CLI_AUTO2 = ""
    climovi!CLI_PRENDA = ""
    climovi!CLI_TELEF1 = ""
    climovi!CLI_TELEF2 = ""
    climovi!CLI_OTRO_CONTR = 0
    climovi!CLI_LETRA = 0
    climovi!CLI_GRUPO = 0
    climovi!CLI_SUBGRUPO = 0
    climovi!CLI_nucleo = ""
    climovi!CLI_estado = "A"
    climovi!CLI_programado = ""
    climovi!CLI_PORDESCTO = 0
    climovi!cli_fecha_fac = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    climovi!cli_DIAS_FAC = 0
    climovi!cli_DIAS_CRED = 0
    climovi!CLI_CUENTA_CONTAB = ""
    climovi!CLI_CUENTA_CONTAB2 = ""
    climovi!CLI_DET_TOT = ""
    climovi!cli_limcre = 0
    climovi.Update

Return





Exit Sub
WEXCEL:
 
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  xl.Workbooks.Open Trim(txtarchi.Text), 0, False, 4
  
Return

End Sub

Private Sub Form_Load()
CenterMe frmimport
lblcia.Caption = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & Chr(13) & "Mes : " & Format(LK_NRO_MES, "00")

pub_cadena = "SELECT * FROM MOVICONT WHERE MOV_CODCIA = ? "
Set MOVI = CN.CreateQuery("", pub_cadena)
MOVI(0) = 0
Set regmovi = MOVI.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CODCIA = ? AnD CLI_CP = ?  "
Set CLI_MOVI = CN.CreateQuery("", pub_cadena)
CLI_MOVI(0) = 0
CLI_MOVI(1) = 0
Set climovi = CLI_MOVI.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT MOV_NRO_VOUCHER  FROM MOVICONT WHERE MOV_CODCIA = ? AND (MOV_FECHA >= ? AND MOV_FECHA <=?) AND MOV_NRO_MES = ? AND MOV_TIPMOV = ?   ORDER BY MOV_NRO_VOUCHER DESC "
Set PSMOV_VOU = CN.CreateQuery("", pub_cadena)
PSMOV_VOU.MaxRows = 1
PSMOV_VOU(0) = 0
PSMOV_VOU(1) = LK_FECHA_DIA
PSMOV_VOU(2) = LK_FECHA_DIA
PSMOV_VOU(3) = 0
PSMOV_VOU(4) = 0
Set VOU_MOV = PSMOV_VOU.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)


End Sub

Public Function GENERA_CODI() As Double
Dim NUMCAD, FIJO As String
Dim DIGI As String * 2
Dim i, VINT1, VINT2, VINT3, VINT4 As Double
Dim VSTR1, VSTR2, VSTR3, VSTR4 As String
Dim VFIJO As Double
Dim VVARI As Integer
Dim STRpub_cadena As String
Dim INTpub_cadena As Double
pu_cp = "C"
pu_codclie = 0
SQ_OPER = 2
pu_codcia = LK_CODCIA
LEER_CLI_LLAVE

If cli_mayor.EOF Then
    NUMCAD = "1"
    If LK_EMP = "PAR" Or LK_EMP = "CAM" Or LK_EMP = "PIU" Then
        INTpub_cadena = Val(NUMCAD)
        GoTo GEN
    Else
       INTpub_cadena = Val(NUMCAD)
    End If
Else
    cli_mayor.MoveLast
    NUMCAD = cli_mayor!CLI_CODCLIE
    If LK_EMP = "PAR" Or LK_EMP = "CAM" Or LK_EMP = "PIU" Then
      INTpub_cadena = Val(NUMCAD) + 1
      GoTo GEN
    Else
      COD_ORIGINAL = INTpub_cadena
    End If
End If

VINT2 = 0
NUMCAD = Trim(NUMCAD)
VINT1 = Len(NUMCAD)
If NUMCAD = "1" Or NUMCAD = "2" Or NUMCAD = "0" Then
  VINT2 = 1
  VINT1 = 2
End If
If VINT1 > 1 Then
    VSTR4 = Val(Mid(NUMCAD, 1, VINT1 - 2)) + 1
End If
For i = 1 To VINT1 - 2
   VSTR1 = Mid(VSTR4, i, 1)
   VINT2 = VINT2 + Val(VSTR1)
Next i
VINT3 = VINT2 * 9

VSTR3 = Right(CStr(VINT3), 2)
If Len(VSTR3) = 1 Then
  VSTR3 = "0" & VSTR3
End If
FIJO = VSTR4
STRpub_cadena = FIJO & VSTR3
INTpub_cadena = Val(STRpub_cadena)
GEN:
GENERA_CODI = INTpub_cadena

End Function

Private Sub lblcodibol_Click()

End Sub

Private Sub txtanulados_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
End If

End Sub

Private Sub txtanulados_LostFocus()
SQ_OPER = 1
    pu_cp = "C"
    pu_codclie = Val(txtanulados.Text)
    pu_codcia = LK_CODCIA
    LEER_CLI_LLAVE
    If cli_llave.EOF Then
      MsgBox "Cliente No Existe", 48, Pub_Titulo
      Exit Sub
    End If
    lcodanul.Caption = Trim(cli_llave!cli_nombre)
End Sub

Private Sub txtcodibol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
End If
End Sub

Private Sub txtcodibol_LostFocus()
    SQ_OPER = 1
    pu_cp = "C"
    pu_codclie = Val(txtcodibol.Text)
    pu_codcia = LK_CODCIA
    LEER_CLI_LLAVE
    If cli_llave.EOF Then
      MsgBox "Cliente No Existe", 48, Pub_Titulo
      Exit Sub
    End If
    lcodclie.Caption = Trim(cli_llave!cli_nombre)

End Sub
