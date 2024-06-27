VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmtransporte 
   Caption         =   "DATOS GENERALES DE TRANSPORTISTAS"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   Icon            =   "frmtransporte.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView ListView1 
      Height          =   495
      Left            =   6480
      TabIndex        =   26
      Top             =   6000
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   128
      BackColor       =   14737632
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "DATOS DE TRANSPORTISTA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6375
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   8415
      Begin VB.TextBox TXTPLACA 
         Height          =   285
         Left            =   1290
         TabIndex        =   28
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos de Chofer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2175
         Left            =   360
         TabIndex        =   17
         Top             =   2760
         Width           =   5415
         Begin VB.TextBox TXTNOMBRECH 
            Height          =   285
            Left            =   1260
            TabIndex        =   21
            Top             =   480
            Width           =   3975
         End
         Begin VB.TextBox TXTDIRCH 
            Height          =   285
            Left            =   1260
            TabIndex        =   20
            Top             =   840
            Width           =   3975
         End
         Begin VB.TextBox TXTBREVETECH 
            Height          =   285
            Left            =   1260
            TabIndex        =   19
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox TXTDNICH 
            Height          =   285
            Left            =   1260
            TabIndex        =   18
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label lbletiqueta 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Chofer :"
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
            Index           =   5
            Left            =   525
            TabIndex        =   25
            Top             =   480
            Width           =   600
         End
         Begin VB.Label lbletiqueta 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Direccion :"
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
            Index           =   6
            Left            =   375
            TabIndex        =   24
            Top             =   840
            Width           =   750
         End
         Begin VB.Label lbletiqueta 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Brevete :"
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
            Index           =   7
            Left            =   450
            TabIndex        =   23
            Top             =   1200
            Width           =   675
         End
         Begin VB.Label lbletiqueta 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "D.N.I. :"
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
            Index           =   8
            Left            =   570
            TabIndex        =   22
            Top             =   1560
            Width           =   555
         End
      End
      Begin VB.TextBox TXTDNI 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox TXTRUC 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox TXTDIRECCION 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox TXTNOMBRE 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox TXTKEY 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lbletiqueta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Placa. :"
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
         Index           =   9
         Left            =   615
         TabIndex        =   27
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label lbletiqueta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "D.N.I. :"
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
         Index           =   4
         Left            =   630
         TabIndex        =   11
         Top             =   2040
         Width           =   555
      End
      Begin VB.Label lbletiqueta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "R.U.C. :"
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
         Index           =   3
         Left            =   585
         TabIndex        =   10
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lbletiqueta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Direccion :"
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
         Index           =   2
         Left            =   435
         TabIndex        =   9
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label lbletiqueta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
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
         Left            =   525
         TabIndex        =   8
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lbletiqueta 
         Alignment       =   1  'Right Justify
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
         Left            =   585
         TabIndex        =   7
         Top             =   600
         Width           =   600
      End
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   625
      Left            =   9120
      Picture         =   "frmtransporte.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1300
   End
   Begin VB.CommandButton CmdModificar 
      Caption         =   "&Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   625
      Left            =   9120
      Picture         =   "frmtransporte.frx":044E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1300
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   625
      Left            =   9120
      Picture         =   "frmtransporte.frx":0598
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   1300
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   625
      Left            =   9120
      Picture         =   "frmtransporte.frx":06E2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   1300
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   625
      Left            =   9120
      Picture         =   "frmtransporte.frx":0B24
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   1300
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4914&
      BorderStyle     =   1  'Fixed Single
      Height          =   7095
      Index           =   5
      Left            =   9000
      TabIndex        =   5
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmtransporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub GRABAR_TRA()
    If Left(CmdModificar.Caption, 2) = "&G" Then
       ven_llave.Edit
    Else
       ven_llave.AddNew
    End If
    
    ven_llave!TRN_KEY = Val(TXTKEY.Text)
    ven_llave!TRN_CODCIA = LKCODCIA
    ven_llave!TRN_NOMBRE = TXTNOMBRE.Text
    ven_llave!TRN_DIRECCION = TXTDIRECCION.Text
    ven_llave!TRN_RUC = TXTRUC.Text
    ven_llave!TRN_DNI = TXTDNI.Text
    ven_llave!TRN_PLACA = TXTPLACA.Text
    ven_llave!TRN_CHOFER = TXTNOMBRECH.Text
    ven_llave!TRN_DIR_CHOFER = TXTDIRCH.Text
    ven_llave!TRN_BREVETE = TXTBREVETECH.Text
    ven_llave!TRN_DNI_CHOFER = TXTDNICH.Text
    ven_llave.Update
End Sub

Public Sub LLENA_TRA(ban As Integer)
Dim I As Integer

    If ban = 0 Then
       If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
       Else
          TXTKEY.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
       End If
       PUB_CODVEN = Val(TXTKEY.Text)
       pu_codcia = LK_CODCIA
       PUB_CODCIA = LK_CODCIA
       SQ_OPER = 1
       LEER_VEN_LLAVE
    End If

    TXTKEY.Text = ven_llave!TRN_KEY
    TXTNOMBRE.Text = ven_llave!TRN_NOMBRE
    TXTDIRECCION.Text = ven_llave!TRN_DIRECCION
    TXTRUC.Text = ven_llave!TRN_RUC
    TXTDNI.Text = ven_llave!TRN_DNI
    TXTPLACA.Text = ven_llave!TRN_PLACA
    TXTNOMBRECH.Text = ven_llave!TRN_CHOFER
    TXTDIRCH.Text = ven_llave!TRN_DIR_CHOFER
    TXTBREVETECH.Text = ven_llave!TRN_BREVETE
    TXTDNICH.Text = ven_llave!TRN_DNI_CHOFER

End Sub

Public Function GENERA_TRA() As Integer
Dim VALOR As Integer
Dim ven_loc As rdoResultset
Dim PSVEN_LOC  As rdoQuery

    pub_cadena = "SELECT TRN_KEY FROM TRANSPORTE WHERE TRN_CODCIA  = ?  ORDER BY TRN_KEY"
    Set PSVEN_LOC = CN.CreateQuery("", pub_cadena)
    PSVEN_LOC(0) = 0
    Set ven_loc = PSVEN_LOC.OpenResultset(rdOpenKeyset, rdConcurValues)
    PSVEN_LOC(0) = LK_CODCIA
    ven_loc.Requery
    If ven_loc.EOF Then
     VALOR = 0
    Else
     ven_loc.MoveLast
     VALOR = ven_loc!VEM_CODVEN
    End If
    GENERA_TRA = VALOR + 1
End Function

Private Sub LIMPIA_TRA()
    TXTKEY.Text = ""
    TXTNOMBRE.Text = ""
    TXTDIRECCION.Text = ""
    TXTRUC.Text = ""
    TXTDNI.Text = ""
    TXTPLACA.Text = ""
    TXTNOMBRECH.Text = ""
    TXTDIRCH.Text = ""
    TXTBREVETECH.Text = ""
    TXTDNICH.Text = ""
End Sub
Private Sub cmdagregar_Click()
'On Error GoTo ESCAPA
If Left(CmdAgregar.Caption, 2) = "&A" Then
    CmdAgregar.Caption = "&Grabar"
    cmdCancelar.Enabled = True
    CmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    LIMPIA_TRA
    DESBLOQUEA_TEXT TXTKEY, TXTNOMBRE, TXTDIRECCION, TXTRUC, TXTDNI, TXTPLACA, TXTNOMBRECH, TXTDIRCH, TXTBREVETECH, TXTDNICH
    TXTKEY = GENERA_TRA
    TXTNOMBRE.SetFocus
Else
   If TXTNOMBRE.Text = "" Or Len(TXTNOMBRE.Text) = 0 Then
       MsgBox "Ingrese Nombre de Transportista ..!!!", 48, Pub_Titulo
       Azul TXTNOMBRE, TXTNOMBRE
       Exit Sub
   End If
   '"SI GRABA.."
    SQ_OPER = 1
    PUB_CODVEN = Val(FrmVen.Txt_key.Text)
    pu_codcia = LK_CODCIA
    LEER_TRA_LLAVE
    If Not ven_llave.EOF Then
       MsgBox "Registro ,  EXISTE ... ", 48, Pub_Titulo
       Azul FrmVen.Txt_key, Txt_key
       Exit Sub
    End If
   Screen.MousePointer = 11
   GRABAR_VEN
   MENSAJE_VEN "Bancos , AGREGADO... "
   CmdAgregar.Caption = "&Agregar"
   cmdEliminar.Enabled = True
   CmdModificar.Enabled = True
   LIMPIA_VEN
   BLOQUEA_TEXT TXTNOMBRE, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
   BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, TXTDIRECCION, txttelecasa, txttelecelu, Check1
   BLOQUEA_TEXT serie_nc, numfac_nc, numfac_nc_f, chenc, serie_nd, numfac_nd, numfac_nd_f, chend, cmbtransporte
   remi.Enabled = False
   txtfechaing.Enabled = False
   Txt_key.Locked = False
   Txt_key.SetFocus
   Screen.MousePointer = 0
      
End If
End Sub
