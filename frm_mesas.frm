VERSION 5.00
Begin VB.Form frm_mesas 
   Caption         =   "Mantenimiento de Mesas"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   ControlBox      =   0   'False
   Icon            =   "frm_mesas.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   8325
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10515
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frm_mesas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1650
      Width           =   1185
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10560
      Picture         =   "frm_mesas.frx":0F2C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   1185
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Modificar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10515
      Picture         =   "frm_mesas.frx":1CEE
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   570
      Width           =   1185
   End
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10515
      Picture         =   "frm_mesas.frx":2B88
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4890
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10515
      Picture         =   "frm_mesas.frx":33FE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3810
      Width           =   1185
   End
   Begin VB.PictureBox ListView1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   6480
      ScaleHeight     =   435
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   7080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame F1 
      Height          =   6615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10095
      Begin VB.ComboBox cmbzona 
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Top             =   1440
         Width           =   4815
      End
      Begin VB.TextBox txtnombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   1
         Top             =   840
         Width           =   4815
      End
      Begin VB.TextBox Txt_key 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Zona :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Timer PARPADEA 
      Interval        =   100
      Left            =   120
      Top             =   4800
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4914&
      BorderStyle     =   1  'Fixed Single
      Height          =   7095
      Index           =   5
      Left            =   10320
      TabIndex        =   8
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label LblMensaje 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   900
   End
End
Attribute VB_Name = "frm_mesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pasa As Boolean
Dim loc_key As Integer
Dim CU As Integer

Dim PS_TRAONE As rdoQuery
Dim TRANSPORTEONE As rdoResultset

Public Function GENERA_VEN() As Integer
Dim valor As Integer
Dim ven_loc As rdoResultset
Dim PSVEN_LOC  As rdoQuery
pub_cadena = "SELECT MES_CODMES FROM MESAS WHERE MES_CODCIA  = ?  ORDER BY MES_CODMES"
Set PSVEN_LOC = CN.CreateQuery("", pub_cadena)
PSVEN_LOC(0) = 0
Set ven_loc = PSVEN_LOC.OpenResultset(rdOpenKeyset, rdConcurValues)
PSVEN_LOC(0) = LK_CODCIA
ven_loc.Requery
If ven_loc.EOF Then
 valor = 0
Else
 ven_loc.MoveLast
 valor = ven_loc!MES_CODMES
End If
GENERA_VEN = valor + 1

End Function

Public Sub GRABAR_VEN()
Dim NAMETRA As String
If Left(cmdModificar.Caption, 2) = "&G" Then
   ven_llave.Edit
Else
   ven_llave.AddNew
       
End If

ven_llave!VEM_CODVEN = Val(FrmVen.Txt_key.Text)
ven_llave!VEM_NOMBRE = FrmVen.txtnombre.Text
ven_llave!vem_codcia = LK_CODCIA
ven_llave!VEM_SERIE_G = Val(FrmVen.serie_g.Text)
ven_llave!VEM_NUMFAC_G_INI = Val(FrmVen.numfac_g.Text)
ven_llave!VEM_SERIE_B = Val(FrmVen.Serie_b.Text)
ven_llave!VEM_NUMFAC_B_INI = Val(FrmVen.numfac_b.Text)
ven_llave!VEM_SERIE_F = Val(FrmVen.serie_f.Text)

ven_llave!VEM_NUMFAC_F_INI = Val(FrmVen.numfac_f.Text)
ven_llave!VEM_NUMFAC_G_FIN = Val(FrmVen.numfac_g_f.Text)
ven_llave!VEM_NUMFAC_B_FIN = Val(FrmVen.numfac_b_f.Text)
ven_llave!VEM_NUMFAC_F_FIN = Val(FrmVen.numfac_f_f.Text)


ven_llave!VEM_SERIE_P = Val(FrmVen.serie_p.Text)
ven_llave!VEM_NUMFAC_P_INI = Val(FrmVen.numfac_p.Text)
ven_llave!VEM_NUMFAC_P_FIN = Val(FrmVen.numfac_p_f.Text)
ven_llave!VEM_FLAG_P = " "
If Check1.Value = 1 Then
  ven_llave!VEM_FLAG_P = "A"
End If

ven_llave!VEM_SERIE_N = Val(FrmVen.serie_nc.Text)
ven_llave!VEM_SERIE_D = Val(FrmVen.serie_nd.Text)
ven_llave!VEM_NUMFAC_N_INI = Val(FrmVen.numfac_nc.Text)
ven_llave!VEM_NUMFAC_D_INI = Val(FrmVen.numfac_nd.Text)
ven_llave!VEM_NUMFAC_N_FIN = Val(FrmVen.numfac_nc_f.Text)
ven_llave!VEM_NUMFAC_D_FIN = Val(FrmVen.numfac_nd_f.Text)
ven_llave!VEM_FLAG_N = " "
If chenc.Value = 1 Then
  ven_llave!VEM_FLAG_N = "A"
End If
ven_llave!VEM_FLAG_D = " "
If chend.Value = 1 Then
  ven_llave!VEM_FLAG_D = "A"
End If

ven_llave!VEM_FECHA_ING = txtfechaing.Text
ven_llave!VEM_DIRECCION = FrmVen.txtdireccion.Text
ven_llave!VEM_TELE_CASA = FrmVen.txttelecasa.Text
ven_llave!VEM_TELE_CELU = FrmVen.txttelecelu.Text
ven_llave!VEM_SERIE_R = Val(FrmVen.remi.Text)
ven_llave!VEM_FLAG_G = " "
ven_llave!VEM_FLAG_B = " "
ven_llave!VEM_FLAG_F = " "
If cheguia.Value = 1 Then
  ven_llave!VEM_FLAG_G = "A"
End If
If cheboleta.Value = 1 Then
  ven_llave!VEM_FLAG_B = "A"
End If
If chefactura.Value = 1 Then
  ven_llave!VEM_FLAG_F = "A"
End If
ven_llave("VEM_TRNKEY") = Val(Right(cmbtransporte.Text, 10))
NAMETRA = IIf(cmbtransporte.Text = "", " ", cmbtransporte.Text)
ven_llave("VEM_TRANSPORTISTA") = Mid(NAMETRA, 1, 50)
ven_llave.Update
SQ_OPER = 2
PUB_CODCIA = LK_CODCIA
PUB_CODVEN = Val(FrmVen.Txt_key.Text)
LEER_PAR_LLAVE
If pac_llave.EOF Then
pac_llave.AddNew
Else
pac_llave.Edit
End If
pac_llave!pac_codcia = LK_CODCIA
pac_llave!pac_codven = PUB_CODVEN
pac_llave!PAC_ARCHI_F = txtfac.Text
pac_llave!PAC_ARCHI_B = txtbol.Text
pac_llave!PAC_ARCHI_G = txtguia.Text
pac_llave!PAC_ARCHI_GUIA = txtgr.Text
pac_llave!PAC_ARCHI_NC = txtnc.Text
pac_llave!PAC_ARCHI_ND = txtnd.Text
pac_llave!PAC_FLAG_CIA = " "
pac_llave.Update

End Sub
Public Sub MENSAJE_VEN(TEXTO As String)
  LblMensaje.Caption = TEXTO
  PARPADEA.Enabled = True
End Sub

Public Sub LLENA_VEN(ban As Integer)
Dim i As Integer
If ban = 0 Then
       If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
         Else
          Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
       End If
       PUB_CODVEN = Val(Txt_key.Text)
       pu_codcia = LK_CODCIA
       PUB_CODCIA = LK_CODCIA
       SQ_OPER = 1
       LEER_VEN_LLAVE
End If

FrmVen.Txt_key.Text = Trim(Nulo_Valors(ven_llave!VEM_CODVEN))
FrmVen.txtnombre.Text = Trim(Nulo_Valors(ven_llave!VEM_NOMBRE))
FrmVen.serie_g.Text = Trim(Nulo_Valors(ven_llave!VEM_SERIE_G))

FrmVen.serie_nc.Text = Trim(Nulo_Valors(ven_llave!VEM_SERIE_N))
FrmVen.serie_nd.Text = Trim(Nulo_Valors(ven_llave!VEM_SERIE_D))
FrmVen.numfac_nc.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_N_INI))
FrmVen.numfac_nd.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_D_INI))
FrmVen.numfac_nc_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_N_FIN))
FrmVen.numfac_nd_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_D_FIN))

FrmVen.numfac_g.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_G_INI))

FrmVen.Serie_b.Text = Trim(Nulo_Valors(ven_llave!VEM_SERIE_B))
FrmVen.serie_p.Text = Trim(Nulo_Valors(ven_llave!VEM_SERIE_P))
FrmVen.numfac_b.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_B_INI))
FrmVen.serie_f.Text = Trim(Nulo_Valors(ven_llave!VEM_SERIE_F))
FrmVen.numfac_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_F_INI))
FrmVen.numfac_p.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_P_INI))
FrmVen.numfac_g_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_G_FIN))
FrmVen.numfac_b_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_B_FIN))
FrmVen.numfac_f_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_F_FIN))
FrmVen.numfac_p_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_P_FIN))
If Not IsNull(ven_llave!VEM_FECHA_ING) Then
  txtfechaing.Text = Format(Nulo_Valors(ven_llave!VEM_FECHA_ING), "dd/mm/yyyy")
End If
txtfechaing.Mask = "##/##/####"
FrmVen.txtdireccion.Text = Trim(Nulo_Valors(ven_llave!VEM_DIRECCION))
FrmVen.txttelecasa.Text = Trim(Nulo_Valors(ven_llave!VEM_TELE_CASA))
FrmVen.txttelecelu.Text = Trim(Nulo_Valors(ven_llave!VEM_TELE_CELU))
FrmVen.remi.Text = Nulo_Valor0(ven_llave!VEM_SERIE_R)
FindInCmb Nulo_Valor0(ven_llave!VEM_TRNKEY)
cheguia.Value = 0
cheboleta.Value = 0
chefactura.Value = 0
Check1.Value = 0
chenc.Value = 0
chend.Value = 0
If UCase(Nulo_Valors(ven_llave!VEM_FLAG_G)) = "A" Then
  cheguia.Value = 1
End If
If UCase(Nulo_Valors(ven_llave!VEM_FLAG_B)) = "A" Then
  cheboleta.Value = 1
End If
If UCase(Nulo_Valors(ven_llave!VEM_FLAG_F)) = "A" Then
  chefactura.Value = 1
End If
If UCase(Nulo_Valors(ven_llave!VEM_FLAG_P)) = "A" Then
  Check1.Value = 1
End If

If UCase(Nulo_Valors(ven_llave!VEM_FLAG_N)) = "A" Then
  chenc.Value = 1
End If
If UCase(Nulo_Valors(ven_llave!VEM_FLAG_D)) = "A" Then
  chend.Value = 1
End If
SQ_OPER = 2
PUB_CODCIA = LK_CODCIA
PUB_CODVEN = Val(ven_llave!VEM_CODVEN)
LEER_PAR_LLAVE
If Not pac_llave.EOF Then
 txtfac.Text = Trim(pac_llave!PAC_ARCHI_F)
 txtbol.Text = Trim(pac_llave!PAC_ARCHI_B)
 txtguia.Text = Trim(pac_llave!PAC_ARCHI_G)
 txtgr.Text = Trim(pac_llave!PAC_ARCHI_GUIA)
 txtnc.Text = Trim(pac_llave!PAC_ARCHI_NC)
 txtnd.Text = Trim(pac_llave!PAC_ARCHI_ND)
End If


End Sub
Public Sub LIMPIA_VEN()
Txt_key.Text = ""
txtnombre.Text = ""
cmbzona.ListIndex = -1

End Sub

Private Sub cheboleta_Click()
If Serie_b.Enabled Then
 Serie_b.SetFocus
End If
End Sub

Private Sub chefactura_Click()
If serie_f.Enabled Then
 serie_f.SetFocus
End If
End Sub

Private Sub cheguia_Click()
If serie_g.Enabled Then
 serie_g.SetFocus
End If
End Sub

Private Sub cmdagregar_Click()
'On Error GoTo ESCAPA
If Left(cmdAgregar.Caption, 2) = "&A" Then
    cmdAgregar.Caption = "&Grabar"
    cmdCancelar.Enabled = True
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    LIMPIA_VEN
    DESBLOQUEA_TEXT txtnombre, Txt_key, cmbzona
    
   
    frm_mesas.Txt_key = GENERA_VEN
    
    frm_mesas.txtnombre.SetFocus
    'AGREGAMOS EN BLANCO
Else
   If frm_mesas.txtnombre.Text = "" Or Len(frm_mesas.txtnombre.Text) = 0 Then
       MsgBox "Ingrese Nombre de Vendedor ..!!!", 48, Pub_Titulo
       Azul txtnombre, txtnombre
       Exit Sub
   End If
   
   '"SI GRABA.."
    SQ_OPER = 1
    PUB_CODVEN = Val(frm_mesas.Txt_key.Text)
    pu_codcia = LK_CODCIA
    LEER_VEN_LLAVE
    If Not ven_llave.EOF Then
       MsgBox "Registro ,  EXISTE ... ", 48, Pub_Titulo
       Azul FrmVen.Txt_key, Txt_key
       Exit Sub
    End If
   Screen.MousePointer = 11
   GRABAR_VEN
   MENSAJE_VEN "Bancos , AGREGADO... "
   cmdAgregar.Caption = "&Agregar"
   cmdEliminar.Enabled = True
   cmdModificar.Enabled = True
   LIMPIA_VEN
   BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
   BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtdireccion, txttelecasa, txttelecelu, Check1
   BLOQUEA_TEXT serie_nc, numfac_nc, numfac_nc_f, chenc, serie_nd, numfac_nd, numfac_nd_f, chend, cmbtransporte
   remi.Enabled = False
   txtfechaing.Enabled = False
   Txt_key.Locked = False
   Txt_key.SetFocus
   Screen.MousePointer = 0
      
End If
   
End Sub

Private Sub cmdAgregar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    
End If

End Sub

Private Sub cmdcancelar_Click()
If Left(cmdAgregar.Caption, 2) = "&A" And Left(cmdModificar.Caption, 2) = "&M" Then
    LIMPIA_VEN
    Txt_key.Locked = False
    MENSAJE_VEN "Proceso Cancelado... !!!    "
    Txt_key.Enabled = True
    Txt_key.SetFocus
     Exit Sub
End If
     Screen.MousePointer = 11
     If Left(cmdModificar.Caption, 2) = "&G" Then
        cmdModificar.Caption = "&Modificar"
        LLENA_VEN 1
        BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
        BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtdireccion, txttelecasa, txttelecelu, Check1
        BLOQUEA_TEXT serie_nc, numfac_nc, numfac_nc_f, chenc, serie_nd, numfac_nd, numfac_nd_f, chend, cmbtransporte
        remi.Enabled = False
        txtfechaing.Enabled = False
        
        Txt_key.Locked = True
     Else
        cmdAgregar.Caption = "&Agregar"
        LIMPIA_VEN
        BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
        BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtdireccion, txttelecasa, txttelecelu, Check1
        BLOQUEA_TEXT serie_nc, numfac_nc, numfac_nc_f, chenc, serie_nd, numfac_nd, numfac_nd_f, chend, cmbtransporte
        remi.Enabled = False
        txtfechaing.Enabled = False
        Txt_key.Locked = False
     End If
     cmdCerrar.Caption = "&Cerrar"
     cmdCancelar.Enabled = True
     cmdAgregar.Enabled = True
     cmdModificar.Enabled = True
     cmdEliminar.Enabled = True
     Txt_key.Enabled = True
     MENSAJE_VEN "Proceso Cancelado... !!!    "
     Txt_key.SetFocus
     Screen.MousePointer = 0

End Sub

Private Sub cmdCerrar_Click()
ws_conta = 0
Unload FrmVen

End Sub

Private Sub cmdCerrar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    FrmVen.Txt_key.SetFocus
End If

End Sub

Private Sub cmdEliminar_Click()
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset

If Len(Txt_key) = 0 Or Len(txtnombre) = 0 Then
   MENSAJE_VEN "NO a seleccionado NADA ... !"
   Exit Sub
End If
  pub_cadena = "SELECT FAR_CODVEN FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODVEN = ? "
  Set PS_REP01 = CN.CreateQuery("", pub_cadena)
  PS_REP01(0) = 0
  PS_REP01(1) = 0
  PS_REP01.MaxRows = 1
  Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = ven_llave!VEM_CODVEN
  llave_rep01.Requery
  If Not llave_rep01.EOF Then
     Screen.MousePointer = 0
     MsgBox "NO se Puede Eliminar ...  Vendedor  TIENE H I S T O R I A.. ", 48, Pub_Titulo
     Exit Sub
  End If
  
  pub_mensaje = " ¿Desea Eliminar el Registro... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbYes Then   ' El usuario eligió
    Screen.MousePointer = 11
    ven_llave.Delete
    Txt_key.Text = ""
    Txt_key.Locked = False
    LIMPIA_VEN
    MENSAJE_VEN "Registro   ELIMINADO ... "
    Screen.MousePointer = 0
   Exit Sub
  End If
  Screen.MousePointer = 0
End Sub

Private Sub CmdModificar_Click()
If Len(Txt_key) = 0 Then
   MENSAJE_VEN "NO a seleccionado NADA ... !"
   Exit Sub
End If
If Left(cmdModificar.Caption, 2) = "&M" Then
    cmdModificar.Caption = "&Grabar"
    cmdAgregar.Enabled = False
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = True
    Txt_key.Locked = True
    DESBLOQUEA_TEXT txtnombre, Txt_key, cmbzona
    txtnombre.SetFocus
Else
    '*Grabar las modificaciones
    If txtnombre.Text = "" Or Len(txtnombre.Text) = 0 Then
         MsgBox " Nombre Invalido ....", 48, Pub_Titulo
         Exit Sub
    End If
    WSFECHA = ES_FECHAS(txtfechaing)
    If WSFECHA = "1" Then
      MsgBox " Fecha Invalidad ...", 48, Pub_Titulo
      Azul2 txtfechaing, txtfechaing
      Exit Sub
    End If
    txtfechaing.Text = Format(WSFECHA, "dd/mm/yyyy")
     Screen.MousePointer = 11
     GRABAR_VEN
     MENSAJE_VEN "Registro , MODIFICADO... "
     cmdModificar.Caption = "&Modificar"
     cmdCancelar.Enabled = True
     cmdAgregar.Enabled = True
     cmdEliminar.Enabled = True
     Txt_key.Locked = True
     BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
     BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtdireccion, txttelecasa, txttelecelu, Check1
     BLOQUEA_TEXT serie_nc, numfac_nc, numfac_nc_f, chenc, serie_nd, numfac_nd, numfac_nd_f, chend, cmbtransporte
     remi.Enabled = False
     txtfechaing.Enabled = False
     Screen.MousePointer = 0
End If

End Sub


Private Sub Form_Load()
Unload FORMGEN
If LK_CODCIA = "04" Then
'  FrmVen.Caption = "&Chofer / Solic."
'  F1.Caption = "&Chofer / Solic."
Else
'  FrmVen.Caption = "&Vendedor"
 ' F1.Caption = "Vendedor"
End If

loc_key = 0
LIMPIA_VEN
BLOQUEA_TEXT txtnombre, Txt_key, cmbzona


Txt_key.Enabled = True


Llenamesas

pub_cadena = "SELECT * FROM MESAS WHERE MES_CODMES = ? ORDER BY MES_NOMBRE"
Set PS_TRAONE = CN.CreateQuery("", pub_cadena)
PS_TRAONE(0) = 0
Set TRANSPORTEONE = PS_TRAONE.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

End Sub

Private Sub Form_Unload(Cancel As Integer)
ws_conta = 0
End Sub

Public Sub BLOQUEA_TEXT(Optional o1, Optional o2, Optional o3, Optional o4, Optional o5, Optional o6, Optional o7, Optional o8, Optional o9, Optional o10, Optional o11, Optional o12)
'** BLOQUEA TEXTBOX  CANTIDAD DE OBJECTOS **
If Not IsMissing(o1) Then
 o1.Enabled = False
' o1.BackColor = QBColor(7)
End If
If Not IsMissing(o2) Then
 o2.Enabled = False
 'o2.BackColor = QBColor(7)
End If
If Not IsMissing(o3) Then
 o3.Enabled = False
 'o3.BackColor = QBColor(7)
End If
If Not IsMissing(o4) Then
 o4.Enabled = False
 'o4.BackColor = QBColor(7)
End If
If Not IsMissing(o5) Then
 o5.Enabled = False
 'o5.BackColor = QBColor(7)
End If
If Not IsMissing(o6) Then
 o6.Enabled = False
 'o6.BackColor = QBColor(7)
End If
If Not IsMissing(o7) Then
 o7.Enabled = False
 'o7.BackColor = QBColor(7)
End If
If Not IsMissing(o8) Then
 o8.Enabled = False
 'o8.BackColor = QBColor(7)
End If
If Not IsMissing(o9) Then
 o9.Enabled = False
 'o9.BackColor = QBColor(7)
End If
If Not IsMissing(o10) Then
 o10.Enabled = False
 'o10.BackColor = QBColor(7)
End If
If Not IsMissing(o11) Then
 o11.Enabled = False
 'o11.BackColor = QBColor(7)
End If
If Not IsMissing(o12) Then
 o12.Enabled = False
 'o12.BackColor = QBColor(7)
End If

End Sub
Public Sub DESBLOQUEA_TEXT(Optional o1, Optional o2, Optional o3, Optional o4, Optional o5, Optional o6, Optional o7, Optional o8, Optional o9, Optional o10, Optional o11, Optional o12)
'** BLOQUEA TEXTBOX  CANTIDAD DE OBJECTOS **
If Not IsMissing(o1) Then
 o1.Enabled = True
' o1.BackColor = QBColor(15)
End If
If Not IsMissing(o2) Then
 o2.Enabled = True
' o2.BackColor = QBColor(15)
End If
If Not IsMissing(o3) Then
 o3.Enabled = True
' o3.BackColor = QBColor(15)
End If
If Not IsMissing(o4) Then
 o4.Enabled = True
' o4.BackColor = QBColor(15)
End If
If Not IsMissing(o5) Then
 o5.Enabled = True
' o5.BackColor = QBColor(15)
End If
If Not IsMissing(o6) Then
 o6.Enabled = True
' o6.BackColor = QBColor(15)
End If
If Not IsMissing(o7) Then
 o7.Enabled = True
' o7.BackColor = QBColor(15)
End If
If Not IsMissing(o8) Then
 o8.Enabled = True
' o8.BackColor = QBColor(15)
End If
If Not IsMissing(o9) Then
 o9.Enabled = True
' o9.BackColor = QBColor(15)
End If
If Not IsMissing(o10) Then
 o10.Enabled = True
' o10.BackColor = QBColor(15)
End If
If Not IsMissing(o11) Then
 o11.Enabled = True
' o11.BackColor = QBColor(15)
End If
If Not IsMissing(o12) Then
 o12.Enabled = True
' o12.BackColor = QBColor(15)
End If

End Sub

Private Sub ListView1_GotFocus()
If loc_key <> 0 Then
 Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
 ListView1.ListItems.Item(loc_key).Selected = True
 ListView1.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView1.SelectedItem.Index
 Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 Txt_key.Text = ""
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
txt_key_KeyPress 13
End Sub

Private Sub numfac_b_f_GotFocus()
Azul numfac_b_f, numfac_b_f
End Sub

Private Sub numfac_b_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 serie_f.SetFocus
End If

End Sub

Private Sub numfac_b_GotFocus()
Azul numfac_b, numfac_b
End Sub

Private Sub numfac_b_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_b_f.SetFocus
End If

End Sub

Private Sub numfac_f_f_GotFocus()
Azul numfac_f_f, numfac_f_f
End Sub

Private Sub numfac_f_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then
  Exit Sub
End If
If cmdModificar.Enabled Then
   cmdModificar.SetFocus
Else
   cmdAgregar.SetFocus
End If

End Sub

Private Sub numfac_f_GotFocus()
Azul numfac_f, numfac_f
End Sub

Private Sub numfac_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  numfac_f_f.SetFocus
End If
End Sub

Private Sub numfac_g_f_GotFocus()
Azul numfac_g_f, numfac_g_f
End Sub

Private Sub numfac_g_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 Serie_b.SetFocus
End If

End Sub

Private Sub numfac_g_GotFocus()
Azul numfac_g, numfac_g
End Sub

Private Sub numfac_g_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_g_f.SetFocus
End If

End Sub

Private Sub remi_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
End Sub

Private Sub Serie_b_GotFocus()
Azul Serie_b, Serie_b
End Sub

Private Sub Serie_b_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_b.SetFocus
End If

End Sub

Private Sub serie_f_GotFocus()
Azul serie_f, serie_f
End Sub

Private Sub serie_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_f.SetFocus
End If

End Sub

Private Sub serie_g_GotFocus()
Azul serie_g, serie_g
End Sub

Private Sub serie_g_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_g.SetFocus
End If
End Sub

Private Sub txt_key_GotFocus()
 Azul Txt_key, Txt_key
End Sub
Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And Txt_key.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  Txt_key.SelStart = Len(Txt_key.Text)
  DoEvents
fin:

End Sub
Private Sub txt_key_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem
If KeyAscii = 27 And Trim(txtnombre.Text) = "" Then
 Txt_key.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
pu_codclie = Val(Txt_key.Text)
If Len(Txt_key.Text) = 0 Or Txt_key.Locked Then
   Exit Sub
End If
If pu_codclie <> 0 And IsNumeric(Txt_key.Text) = True Then
   loc_key = 0
   On Error GoTo mucho
   PUB_CODVEN = Val(Txt_key.Text)
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_VEN_LLAVE
   On Error GoTo 0
   If ven_llave.EOF Then
     Azul Txt_key, Txt_key
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Txt_key.SetFocus
     GoTo fin
   End If
   ListView1.Visible = False
   cmdCancelar.Enabled = True
   LLENA_VEN 0
   Txt_key.Locked = True
   cmdModificar.SetFocus
   Screen.MousePointer = 0
   Exit Sub
Else
   If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(ListView1.ListItems.Item(loc_key).Text)
   If Trim(UCase(Txt_key.Text)) = Left(valor, Len(Trim(Txt_key.Text))) Then
   Else
      Exit Sub
   End If
   ListView1.Visible = False
   cmdCancelar.Enabled = True
   LLENA_VEN 0
    Txt_key.Locked = True
   cmdCancelar.Enabled = True
    cmdModificar.SetFocus
End If
dale:
mucho:
ListView1.Visible = False
fin:
End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
If Len(Txt_key.Text) = 0 Or Txt_key.Locked = True Or IsNumeric(Txt_key.Text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(Txt_key.Text) = 1 Then
    var = Asc(Txt_key.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 9
    archi = "SELECT * FROM VEMAEST WHERE  VEM_CODCIA = '" & LK_CODCIA & "' AND VEM_NOMBRE BETWEEN '" & Txt_key.Text & "' AND  '" & var & "' ORDER BY VEM_NOMBRE"
    PROC_LISVIEW ListView1
    loc_key = 1
    If ListView1.Visible = False Then
        loc_key = 0
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
'If ListView1.Visible Then
'  Set itmFound = ListView1.FindItem(LTrim(Txt_key.Text), lvwText, , lvwPartial)
'  If itmFound Is Nothing Then
'  Else
'   itmFound.EnsureVisible
'   itmFound.Selected = True
'   loc_key = itmFound.Tag
'   If loc_key + 8 > ListView1.ListItems.count Then
'      ListView1.ListItems.Item(ListView1.ListItems.count).EnsureVisible
'   Else
'     ListView1.ListItems.Item(loc_key + 8).EnsureVisible
'   End If
'   DoEvents
'  End If
'  Exit Sub
'End If
End Sub

Private Sub PARPADEA_Timer()
 CU = CU + 1
 LblMensaje.Visible = True 'Not LblMensaje.Visible
 If CU > 8 Then
   CU = 0
   PARPADEA.Enabled = False
   LblMensaje.Visible = False
 End If
End Sub

Private Sub txtdireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul2 txtfechaing, txtfechaing
End If
End Sub

Private Sub txtfechaing_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul txttelecasa, txttelecasa
End If
End Sub

Private Sub txtnombre_GotFocus()
Azul txtnombre, txtnombre
End Sub

Private Sub txtnombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul txtnombre, txtnombre
End If
End Sub

Public Function ES_FECHAS(CAMPOFECHA As MaskEdBox) As String
Dim wfecha As String
ES_FECHAS = "0"
If CAMPOFECHA = "00/00/0000" Then
 Exit Function
End If
If Right(CAMPOFECHA.Text, 2) = "__" Then
  wfecha = Left(CAMPOFECHA.Text, 8)
Else
  wfecha = Trim(CAMPOFECHA.Text)
End If
If Not IsDate(wfecha) Then
  ES_FECHAS = "1"
  Exit Function
End If
ES_FECHAS = wfecha
End Function

Private Sub txttelecasa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul txttelecelu, txttelecelu
End If
End Sub

Private Sub txttelecelu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If F2.Visible Then
    Azul serie_g, serie_g
  Else
   If cmdModificar.Enabled Then
    cmdModificar.SetFocus
   Else
    cmdAgregar.SetFocus
   End If
  End If

End If
End Sub

Private Sub Llenamesas()
Dim PS_MES As rdoQuery
Dim MESAS As rdoResultset
Dim SQL As String
SQL = "SELECT * FROM MESAS ORDER BY MES_NOMBRE"
Set PS_MES = CN.CreateQuery("", SQL)
Set MESAS = PS_MES.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
MESAS.Requery
cmbzona.Clear
Do Until MESAS.EOF
    cmbtransporte.AddItem Trim(MESAS!MES_NOMBRE) & String(80, " ") ' & TRANSPORTE!TRN_KEY
    MESAS.MoveNext
Loop

End Sub
Private Function FindInCmb(ByVal s_transporte As String) As Boolean
Dim i As Long
Dim aux_f As String

    cmbtransporte.ListIndex = -1
    For i = 0 To cmbtransporte.ListCount - 1
     aux_f = cmbtransporte.List(i)
     aux_f = Trim$(Right$(aux_f, 10))
     If Trim(aux_f) = Trim(s_transporte) Then
      cmbtransporte.ListIndex = i
      Exit For
     End If
    Next
   
End Function
