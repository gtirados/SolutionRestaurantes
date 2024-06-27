VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTransportista 
   Caption         =   "Datos de Transportista"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   11505
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView ListView1 
      Height          =   1260
      Left            =   6735
      TabIndex        =   20
      Top             =   2565
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   2223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   16118252
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox txtCertificado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6000
      MaxLength       =   12
      TabIndex        =   31
      Top             =   3480
      Width           =   1575
   End
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
      Left            =   10710
      Picture         =   "transportista.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1590
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
      Left            =   10710
      Picture         =   "transportista.frx":0AEA
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2580
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
      Left            =   10710
      Picture         =   "transportista.frx":18AC
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   510
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
      Left            =   10710
      Picture         =   "transportista.frx":2746
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4830
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
      Left            =   10710
      Picture         =   "transportista.frx":2FBC
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3750
      Width           =   1185
   End
   Begin VB.TextBox txtRuc 
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
      Left            =   3000
      MaxLength       =   12
      TabIndex        =   3
      Top             =   2610
      Width           =   1575
   End
   Begin VB.TextBox txtFindRegistro 
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
      Left            =   3015
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
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
      Left            =   3015
      MaxLength       =   40
      TabIndex        =   1
      Top             =   1680
      Width           =   4815
   End
   Begin VB.TextBox txtdireccion 
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
      Left            =   3015
      MaxLength       =   40
      TabIndex        =   2
      Top             =   2160
      Width           =   4815
   End
   Begin VB.TextBox txtDni 
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
      Left            =   3015
      MaxLength       =   10
      TabIndex        =   4
      Top             =   3135
      Width           =   1335
   End
   Begin VB.TextBox txtPlaca 
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
      Left            =   3015
      MaxLength       =   10
      TabIndex        =   5
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtdnichofer 
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
      Left            =   3000
      MaxLength       =   12
      TabIndex        =   8
      Top             =   5415
      Width           =   1335
   End
   Begin VB.TextBox txtbrevete 
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
      Left            =   3000
      MaxLength       =   40
      TabIndex        =   9
      Top             =   5895
      Width           =   4815
   End
   Begin VB.TextBox txtchofer 
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
      Left            =   3030
      MaxLength       =   40
      TabIndex        =   6
      Top             =   4455
      Width           =   4815
   End
   Begin VB.TextBox txtdirechofer 
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
      Left            =   3030
      MaxLength       =   40
      TabIndex        =   7
      Top             =   4935
      Width           =   4815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   17
      Left            =   5160
      TabIndex        =   32
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Datos de Transportista"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   4230
      TabIndex        =   30
      Top             =   825
      Width           =   2265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Datos de Chofer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   4365
      TabIndex        =   28
      Top             =   4155
      Width           =   1725
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4914&
      BorderStyle     =   1  'Fixed Single
      Height          =   7290
      Index           =   5
      Left            =   10515
      TabIndex        =   26
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   0
      Left            =   1815
      TabIndex        =   19
      Top             =   1200
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   1
      Left            =   1815
      TabIndex        =   18
      Top             =   1680
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   11
      Left            =   1815
      TabIndex        =   17
      Top             =   2160
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   12
      Left            =   1815
      TabIndex        =   16
      Top             =   2640
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   13
      Left            =   1815
      TabIndex        =   15
      Top             =   3120
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Placa :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   14
      Left            =   1815
      TabIndex        =   14
      Top             =   3600
      Width           =   540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DNI :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   13
      Top             =   5415
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Brevete :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   12
      Top             =   5895
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Chofer : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   1830
      TabIndex        =   11
      Top             =   4455
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   1830
      TabIndex        =   10
      Top             =   4935
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2385
      Index           =   6
      Left            =   1425
      TabIndex        =   27
      Top             =   4080
      Width           =   7530
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3300
      Index           =   8
      Left            =   1410
      TabIndex        =   29
      Top             =   705
      Width           =   7530
   End
End
Attribute VB_Name = "frmTransportista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim liIndexRegAct As Integer
Dim PSTRA As rdoQuery
Dim RSTRA As rdoResultset
Dim PSALLTRA As rdoQuery
Dim RSALLTRA As rdoResultset
Dim CODTRANSP As Integer
Dim loc_key  As Integer
Private loListItem As MSComctlLib.ListItem

Public Function GENERA_TRA() As Integer
Dim valor As Integer
Dim ven_loc As rdoResultset
Dim PSVEN_LOC  As rdoQuery
pub_cadena = "SELECT TRN_KEY FROM transporte WHERE TRN_CODCIA='" & LK_CODCIA & "' ORDER BY TRN_KEY"
Set PSVEN_LOC = CN.CreateQuery("", pub_cadena)
Set ven_loc = PSVEN_LOC.OpenResultset(rdOpenKeyset, rdConcurValues)
ven_loc.Requery
If ven_loc.EOF Then
 valor = 0
Else
 ven_loc.MoveLast
 valor = ven_loc!TRN_KEY
End If
GENERA_TRA = valor + 1

End Function

Public Sub GRABAR_TRA()
If Left(cmdModificar.Caption, 2) = "&G" Then
   RSTRA.Edit
Else
   RSTRA.AddNew
End If

    RSTRA!TRN_KEY = Val(txtFindRegistro.Text)
    RSTRA!trn_codcia = LK_CODCIA
    RSTRA!TRN_NOMBRE = txtnombre.Text
    RSTRA!TRN_DIRECCION = txtdireccion.Text
    RSTRA!TRN_RUC = Trim(txtRuc.Text)
    RSTRA!TRN_DNI = Trim(txtDni.Text)
    RSTRA!TRN_PLACA = txtPlaca.Text
    RSTRA!TRN_CHOFER = txtchofer.Text
    RSTRA!TRN_DIR_CHOFER = txtdirechofer.Text
    RSTRA!TRN_BREVETE = txtbrevete.Text
    RSTRA!TRN_DNI_CHOFER = txtdnichofer.Text
    'RSTRA!TRN_ARREN = txtArrenNom.Text
    'RSTRA!TRN_RUC_ARREN = txtArrenRUC.Text
    RSTRA!TRN_MARCA = txtCertificado.Text
    RSTRA.Update
End Sub
Public Sub LLENA_TRA(ban As Integer)
Dim i As Integer
    If ban = 0 Then
       If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
         Else
          txtFindRegistro.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
       End If
       CODTRANSP = Val(txtFindRegistro.Text)
       PUB_CODCIA = LK_CODCIA
       PSTRA(0) = CODTRANSP
       RSTRA.Requery
    End If
    If Not RSTRA.EOF Then
        txtnombre.Text = Nulo_Valors(RSTRA!TRN_NOMBRE)
        txtdireccion.Text = Nulo_Valors(RSTRA!TRN_DIRECCION)
        txtRuc.Text = Nulo_Valors(RSTRA!TRN_RUC)
        txtDni.Text = Nulo_Valors(RSTRA!TRN_DNI)
        txtPlaca.Text = Nulo_Valors(RSTRA!TRN_PLACA)
        txtchofer.Text = Nulo_Valors(RSTRA!TRN_CHOFER)
        txtdirechofer.Text = Nulo_Valors(RSTRA!TRN_DIR_CHOFER)
        txtbrevete.Text = Nulo_Valors(RSTRA!TRN_BREVETE)
        txtdnichofer.Text = Nulo_Valors(RSTRA!TRN_DNI_CHOFER)
        txtCertificado.Text = Nulo_Valors(RSTRA!TRN_MARCA)
        'txtArrenNom.Text = Nulo_Valors(RSTRA!TRN_ARREN)
        'txtArrenRUC.Text = Nulo_Valors(RSTRA!TRN_RUC_ARREN)
    Else
        MsgBox "Registro no Existe...", vbInformation, Pub_Titulo
    End If
End Sub
Public Sub LIMPIA_TRA()
    txtFindRegistro.Text = ""
    txtnombre.Text = ""
    txtdireccion.Text = ""
    txtRuc.Text = ""
    txtDni.Text = ""
    txtPlaca.Text = ""
    txtchofer.Text = ""
    txtdirechofer.Text = ""
    txtbrevete.Text = ""
    txtdnichofer.Text = ""
    txtCertificado.Text = ""
    'txtArrenNom.Text = ""
    'txtArrenRUC.Text = ""
End Sub

Private Sub cmdagregar_Click()
'On Error GoTo ESCAPA
If Left(cmdAgregar.Caption, 2) = "&A" Then
    cmdAgregar.Caption = "&Grabar"
    cmdCancelar.Enabled = True
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    LIMPIA_TRA
    FrmVen.DESBLOQUEA_TEXT txtnombre, txtdireccion, txtRuc, txtDni, txtPlaca, txtchofer, txtdirechofer, txtbrevete, txtdnichofer, txtCertificado ', txtArrenNom, txtArrenRUC
    txtFindRegistro = GENERA_TRA
    txtnombre.SetFocus
Else
   If txtnombre.Text = "" Or Len(txtnombre.Text) = 0 Then
       MsgBox "Ingrese Nombre de Vendedor ..!!!", 48, Pub_Titulo
       Azul txtnombre, txtnombre
       Exit Sub
   End If
   
   '"SI GRABA.."
    SQ_OPER = 1
    CODTRANSP = Val(txtFindRegistro.Text)
    PSTRA(0) = CODTRANSP
    RSTRA.Requery
    If Not RSTRA.EOF Then
       MsgBox "Registro ,  EXISTE ... ", 48, Pub_Titulo
       Azul txtFindRegistro, txtFindRegistro
       Exit Sub
    End If
   Screen.MousePointer = 11
   GRABAR_TRA
   
   cmdAgregar.Caption = "&Agregar"
   cmdEliminar.Enabled = True
   cmdModificar.Enabled = True
   LIMPIA_TRA
   FrmVen.BLOQUEA_TEXT txtnombre, txtdireccion, txtRuc, txtDni, txtPlaca, txtchofer, txtdirechofer, txtbrevete, txtdnichofer, txtCertificado ', txtArrenNom, txtArrenRUC
      
   txtFindRegistro.Locked = False
   txtFindRegistro.SetFocus
   Screen.MousePointer = 0
      
End If
   
End Sub

Private Sub cmdcancelar_Click()
If Left(cmdAgregar.Caption, 2) = "&A" And Left(cmdModificar.Caption, 2) = "&M" Then
    LIMPIA_TRA
    txtFindRegistro.Locked = False
    txtFindRegistro.Enabled = True
    txtFindRegistro.SetFocus
     Exit Sub
End If
     Screen.MousePointer = 11
     If Left(cmdModificar.Caption, 2) = "&G" Then
        cmdModificar.Caption = "&Modificar"
        LLENA_TRA 1
        FrmVen.BLOQUEA_TEXT txtnombre, txtdireccion, txtRuc, txtDni, txtPlaca, txtchofer, txtdirechofer, txtbrevete, txtdnichofer, txtCertificado
        txtFindRegistro.Locked = True
     Else
        cmdAgregar.Caption = "&Agregar"
        LIMPIA_TRA
        FrmVen.BLOQUEA_TEXT txtnombre, txtdireccion, txtRuc, txtDni, txtPlaca, txtchofer, txtdirechofer, txtbrevete, txtdnichofer, txtCertificado
        txtFindRegistro.Locked = False
     End If
     cmdCerrar.Caption = "&Cerrar"
     cmdCancelar.Enabled = True
     cmdAgregar.Enabled = True
     cmdModificar.Enabled = True
     cmdEliminar.Enabled = True
     txtFindRegistro.Enabled = True
     txtFindRegistro.SetFocus
     Screen.MousePointer = 0

End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset

'If Len(txtFindRegistro) = 0 Or Len(txtnombre) = 0 Then
'   MENSAJE_VEN "NO a seleccionado NADA ... !"
'   Exit Sub
'End If
'  pub_cadena = "SELECT FAR_CODVEN FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODVEN = ? "
'  Set PS_REP01 = CN.CreateQuery("", pub_cadena)
'  PS_REP01(0) = 0
'  PS_REP01(1) = 0
'  PS_REP01.MaxRows = 1
'  Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
'  PS_REP01(0) = LK_CODCIA
'  PS_REP01(1) = RSTRA!vem_codven
'  llave_rep01.Requery
'  If Not llave_rep01.EOF Then
'     Screen.MousePointer = 0
'     MsgBox "NO se Puede Eliminar ...  Vendedor  TIENE H I S T O R I A.. ", 48, Pub_Titulo
'     Exit Sub
'  End If
'
  pub_mensaje = " ¿Desea Eliminar el Registro... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbYes Then   ' El usuario eligió
    Screen.MousePointer = 11
    RSTRA.Delete
    txtFindRegistro.Text = ""
    txtFindRegistro.Locked = False
    LIMPIA_TRA
    Screen.MousePointer = 0
   Exit Sub
  End If
  Screen.MousePointer = 0
End Sub

Private Sub CmdModificar_Click()
If Len(txtFindRegistro) = 0 Then
   Exit Sub
End If
If Left(cmdModificar.Caption, 2) = "&M" Then
    cmdModificar.Caption = "&Grabar"
    cmdAgregar.Enabled = False
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = True
    txtFindRegistro.Locked = True
    FrmVen.DESBLOQUEA_TEXT txtnombre, txtdireccion, txtRuc, txtDni, txtPlaca, txtchofer, txtdirechofer, txtbrevete, txtdnichofer, txtCertificado
    txtnombre.SetFocus
Else
    '*Grabar las modificaciones
    If txtnombre.Text = "" Or Len(txtnombre.Text) = 0 Then
         MsgBox " Nombre Invalido ....", 48, Pub_Titulo
         Exit Sub
    End If
     Screen.MousePointer = 11
     GRABAR_TRA
     cmdModificar.Caption = "&Modificar"
     cmdCancelar.Enabled = True
     cmdAgregar.Enabled = True
     cmdEliminar.Enabled = True
     txtFindRegistro.Locked = True
     FrmVen.BLOQUEA_TEXT txtnombre, txtdireccion, txtRuc, txtDni, txtPlaca, txtchofer, txtdirechofer, txtbrevete, txtdnichofer, txtCertificado
     Screen.MousePointer = 0
End If

End Sub
Private Sub Form_Load()
        pub_cadena = "SELECT * FROM TRANSPORTE WHERE TRN_KEY = ? and TRN_CODCIA= ?"
    Set PSTRA = CN.CreateQuery("", pub_cadena)
    PSTRA.rdoParameters(0) = 0
    PSTRA.rdoParameters(1) = ""
    Set RSTRA = PSTRA.OpenResultset(rdOpenKeyset, rdConcurValues)
    PSTRA(0) = 0
    PSTRA.rdoParameters(1) = LK_CODCIA
    loc_key = 0
    LIMPIA_TRA
    FrmVen.BLOQUEA_TEXT txtnombre, txtdireccion, txtRuc, txtDni, txtPlaca, txtchofer, txtdirechofer, txtbrevete, txtdnichofer, txtCertificado
    txtFindRegistro.Enabled = True

    ListView1.ColumnHeaders.Add , , "Nombres", 5000
    ListView1.ColumnHeaders.Add , , "Codigo", 1200
End Sub


''=========================PARA BUSCADOR
Private Sub LlenaList()
Dim i As Long
Dim J As Integer
    ListView1.ListItems.Clear
    Do While Not RSALLTRA.EOF
        i = i + 1
        Set loListItem = ListView1.ListItems.Add(, "k" + CStr(RSALLTRA("TRN_KEY")), RSALLTRA("TRN_NOMBRE"))
        loListItem.ListSubItems.Add Key:="Nombre" & CStr(J), Text:=RSALLTRA("TRN_KEY")
        RSALLTRA.MoveNext
    Loop
    Set RSALLTRA = Nothing
    If i > 0 Then
        liIndexRegAct = 1
        Call FormatObjects(1)
    Else
        FormatObjects 2
    End If
End Sub
Private Sub ListView1_DblClick()
    txtFindRegistro.Text = ListView1.SelectedItem.ListSubItems(1).Text
    Call FormatObjects(2)
    LLENA_TRA 0
    cmdCancelar.Enabled = True
    txtFindRegistro.Locked = True
    cmdModificar.SetFocus
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Item.EnsureVisible
    Item.Selected = True
    liIndexRegAct = Item.Index
    txtFindRegistro.Text = Item.Text
    txtFindRegistro.SetFocus
End Sub

Private Sub txtFindRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
Dim liNumItems As Integer

    liNumItems = ListView1.ListItems.count
    If liNumItems = 0 Then Exit Sub
    If KeyCode = 40 Then 'ABAJO
        liIndexRegAct = liIndexRegAct + 1
        If liIndexRegAct > liNumItems Then liIndexRegAct = liNumItems
    ElseIf KeyCode = 38 Then 'ARRIBA
        liIndexRegAct = liIndexRegAct - 1
        If liIndexRegAct <= 0 Then liIndexRegAct = 1
    ElseIf KeyCode = 33 Then '33 RE PAG
        liIndexRegAct = liIndexRegAct - 10
        If liIndexRegAct <= 0 Then liIndexRegAct = 1
    ElseIf KeyCode = 34 Then ' 34 AV PAG
        liIndexRegAct = liIndexRegAct + 10
        If liIndexRegAct > liNumItems Then liIndexRegAct = liNumItems
    End If
If liIndexRegAct = 0 Then Exit Sub
    ListView1.ListItems.Item(liIndexRegAct).EnsureVisible
    ListView1.ListItems.Item(liIndexRegAct).Selected = True
    If KeyCode = 33 Or KeyCode = 34 Or KeyCode = 38 Or KeyCode = 40 Then txtFindRegistro.Text = ListView1.ListItems(liIndexRegAct).Text
    If KeyCode = 13 Then Call ListView1_DblClick
End Sub

Private Sub txtFindRegistro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        ListView1.ListItems.Clear
        liIndexRegAct = 0
        txtFindRegistro = ""
        Call FormatObjects(2)
    End If
    If KeyCode >= 37 And KeyCode <= 40 Or (KeyCode = 33 Or KeyCode = 34) Then Exit Sub
    If Len(txtFindRegistro) = 1 Then
        
        pub_cadena = "SELECT TRN_KEY, TRN_NOMBRE FROM TRANSPORTE WHERE TRN_NOMBRE LIKE '" & Trim(txtFindRegistro) & "%' AND TRN_CODCIA = '" & LK_CODCIA & "'"
        Set PSALLTRA = CN.CreateQuery("", pub_cadena)
        Set RSALLTRA = PSALLTRA.OpenResultset(rdOpenKeyset, rdConcurValues)
        RSALLTRA.Requery
        LlenaList
    End If
    If Len(txtFindRegistro) > 1 Then FindItem Trim(txtFindRegistro)
End Sub

Private Sub FindItem(ByVal sDato As String)
Dim intSelectedOption As Integer
Dim strFindMe As String

   intSelectedOption = lvwText
   Set loListItem = ListView1.FindItem(sDato, intSelectedOption, , lvwPartial)
   If loListItem Is Nothing Then
        Exit Sub
   Else
        loListItem.EnsureVisible
        loListItem.Selected = True
        liIndexRegAct = loListItem.Index
   End If
End Sub
Private Sub FormatObjects(ByVal lFlag As Integer)
    If lFlag = 1 Then
        ListView1.Width = 6000
        ListView1.Height = 3000
        ListView1.Left = txtFindRegistro.Left
        ListView1.Top = txtFindRegistro.Top + txtFindRegistro.Height
        ListView1.Visible = True
    ElseIf lFlag = 2 Then
        ListView1.Visible = False
    End If
End Sub

