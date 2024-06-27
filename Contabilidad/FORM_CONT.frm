VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FORM_CONT 
   BackColor       =   &H00FAEFDA&
   Caption         =   "Maestro de Plan de Cuentas"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   8925
   WindowState     =   2  'Maximized
   Begin ComctlLib.TreeView TreeView1 
      Height          =   7035
      Left            =   195
      TabIndex        =   29
      Top             =   225
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   12409
      _Version        =   327682
      HideSelection   =   0   'False
      Style           =   6
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraorden 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Orden para el Cierre de las Cuentas"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   360
      TabIndex        =   38
      Top             =   8160
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton bajar 
         Caption         =   "Bajar"
         Height          =   375
         Left            =   2760
         TabIndex        =   41
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton subir 
         Caption         =   "Subir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   40
         Top             =   360
         Width           =   615
      End
      Begin VB.ListBox ctaorden 
         Height          =   1035
         Left            =   240
         TabIndex        =   39
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Plan Contable :"
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
      Height          =   7425
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   11775
      Begin VB.CheckBox chkAfectacion 
         BackColor       =   &H00F8DED7&
         Caption         =   "Afectación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10260
         TabIndex        =   52
         Top             =   870
         Width           =   1290
      End
      Begin VB.CommandButton Cmdcopiar 
         Caption         =   "C&opiar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5250
         TabIndex        =   50
         Top             =   6015
         Width           =   1125
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Consolidar Compañias :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1080
         Left            =   5100
         TabIndex        =   47
         Top             =   0
         Width           =   4455
         Begin VB.ListBox LISCIA 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   735
            Left            =   120
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   48
            Top             =   240
            Visible         =   0   'False
            Width           =   4215
         End
      End
      Begin VB.CommandButton ctacierre 
         Caption         =   "&Ctas. Cierre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6915
         TabIndex        =   46
         Top             =   6015
         Width           =   1125
      End
      Begin VB.CheckBox salacu 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Ver Saldo Acumulado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5145
         TabIndex        =   45
         Top             =   2670
         Width           =   2250
      End
      Begin VB.CommandButton comsal 
         Caption         =   "Crear Cuenta en ComSal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9330
         TabIndex        =   44
         Top             =   6495
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.CheckBox checostos 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Cuenta es Centro de Costos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8640
         TabIndex        =   43
         Top             =   3075
         Width           =   2820
      End
      Begin VB.Frame fradestino 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Ctas. Destino"
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
         Height          =   1800
         Left            =   5100
         TabIndex        =   22
         Top             =   3405
         Visible         =   0   'False
         Width           =   6480
         Begin VB.TextBox txtautodp5 
            Height          =   285
            Left            =   3960
            TabIndex        =   17
            Top             =   1395
            Width           =   930
         End
         Begin VB.TextBox txtautodp4 
            Height          =   285
            Left            =   3000
            TabIndex        =   16
            Top             =   1395
            Width           =   930
         End
         Begin VB.TextBox txtautodp3 
            Height          =   285
            Left            =   2040
            TabIndex        =   15
            Top             =   1395
            Width           =   930
         End
         Begin VB.TextBox txtautodp2 
            Height          =   285
            Left            =   1080
            TabIndex        =   14
            Top             =   1395
            Width           =   930
         End
         Begin VB.TextBox txtautodp 
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   1395
            Width           =   930
         End
         Begin VB.TextBox txtautod5 
            Height          =   285
            Left            =   3960
            TabIndex        =   12
            Top             =   810
            Width           =   930
         End
         Begin VB.TextBox txtautod4 
            Height          =   285
            Left            =   3000
            TabIndex        =   11
            Top             =   810
            Width           =   930
         End
         Begin VB.TextBox txtautod3 
            Height          =   285
            Left            =   2040
            TabIndex        =   10
            Top             =   810
            Width           =   930
         End
         Begin VB.TextBox txtautod2 
            Height          =   285
            Left            =   1080
            TabIndex        =   9
            Top             =   810
            Width           =   930
         End
         Begin VB.TextBox txtautoh 
            Height          =   285
            Left            =   2025
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtautod 
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   810
            Width           =   930
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "(%)           1                   2                   3                   4                   5"
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
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   1155
            Width           =   4815
         End
         Begin VB.Label LABEL2 
            BackStyle       =   0  'Transparent
            Caption         =   "Debe :       1                   2                   3                   4                   5"
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
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   555
            Width           =   4815
         End
         Begin VB.Label LABEL1 
            BackStyle       =   0  'Transparent
            Caption         =   "Haber :"
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
            Height          =   255
            Left            =   1410
            TabIndex        =   23
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Orden Cta de Cierre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7125
         TabIndex        =   37
         Top             =   6495
         Width           =   2025
      End
      Begin VB.CommandButton CmdCerrar 
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
         Height          =   375
         Left            =   10245
         TabIndex        =   20
         Top             =   6015
         Width           =   1095
      End
      Begin VB.Frame fraOp 
         BackColor       =   &H00FAEFDA&
         ForeColor       =   &H00FAEFDA&
         Height          =   735
         Left            =   5100
         TabIndex        =   33
         Top             =   5220
         Width           =   6480
         Begin VB.CommandButton cmdcancelar 
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
            Height          =   375
            Left            =   5145
            TabIndex        =   51
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Eliminar 
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
            Height          =   375
            Left            =   3490
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Insertar 
            Caption         =   "&Insertar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1835
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton grabar 
            Caption         =   "&Grabar"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame cuentas 
         BackColor       =   &H00FAEFDA&
         Height          =   855
         Left            =   5100
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   6495
         Begin VB.TextBox txtnombre 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   2
            Top             =   360
            Width           =   4455
         End
         Begin VB.TextBox txtindice 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            TabIndex        =   1
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtcuenta 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
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
            Left            =   1815
            TabIndex        =   31
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Contable"
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
            Left            =   120
            TabIndex        =   30
            Top             =   120
            Width           =   1395
         End
      End
      Begin ComctlLib.ProgressBar PB 
         Height          =   195
         Left            =   7815
         TabIndex        =   34
         Top             =   2700
         Visible         =   0   'False
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Frame fraTipo 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Tipo de Cuenta :"
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
         Height          =   615
         Left            =   5100
         TabIndex        =   32
         Top             =   1950
         Visible         =   0   'False
         Width           =   3135
         Begin VB.ComboBox tipo 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame SIGNOS 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Signos :"
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
         Height          =   615
         Left            =   8340
         TabIndex        =   26
         Top             =   1950
         Visible         =   0   'False
         Width           =   3255
         Begin VB.ComboBox SIGNOH 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "FORM_CONT.frx":0000
            Left            =   2265
            List            =   "FORM_CONT.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox SIGNOD 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "FORM_CONT.frx":0015
            Left            =   705
            List            =   "FORM_CONT.frx":001F
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Haber :"
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
            Height          =   255
            Left            =   1560
            TabIndex        =   28
            Top             =   270
            Width           =   615
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Debe :"
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
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   270
            Width           =   495
         End
      End
      Begin VB.Label lblbarraos 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Solution - Gestion Contable"
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
         Left            =   5010
         TabIndex        =   49
         Top             =   6945
         Width           =   6735
      End
      Begin VB.Label saldos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6960
         TabIndex        =   35
         Top             =   3030
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo de Cuenta.:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5475
         TabIndex        =   36
         Top             =   3030
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FORM_CONT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim com_cont As rdoResultset
Dim PSCOM_CONT As rdoQuery
Dim PSCOM_MAYOR2 As rdoQuery
Dim com_mayor2 As rdoResultset
Dim NIVEL_ACT As Integer
Dim CARAC
Dim fila As Integer
Dim wCOM_NIVEL(6) As Integer
Dim NIVEL_MAX As Integer
Dim ww_nivel As Integer
Dim qq_cuenta As String * 12
Dim qq_indice As String * 12
Dim Posi_Ultimo As Integer
Dim Modo_Acceso As String

Public Function BUSCA_AUTO(valor As String) As Boolean
Dim ba As rdoResultset
Dim cade As String
  cade = "SELECT * FROM COMAEST WHERE COM_CUENTA = '" & Trim(valor) & "' AND COM_CODCIA = '" & LK_CODCIA & "' ORDER BY COM_CUENTA"
  Set ba = CN.OpenResultset(cade, rdOpenKeyset, rdConcurValues)
   ba.Requery
   BUSCA_AUTO = False
   If ba.EOF Then
     Exit Function
   End If
   If ba!com_flag_afectacion = "1" Then
     BUSCA_AUTO = True
   End If

End Function

Public Sub VERI_NIVEL(valor As String)
Dim cade As String
Dim wnivel As Integer
Dim i, windex As Integer
Dim wBusca As String

LOC_CTA_SUP = ""
NIVEL_ACT = 0
cade = Trim(valor)
wnivel = Len(cade)
windex = 0
For i = 1 To 6
  If wCOM_NIVEL(i) = wnivel Then
     windex = i
     Exit For
  End If
Next i
If windex = 0 Then
    MsgBox "Nro. de digitos invalidos, Verificar ..", 48, Pub_Titulo
    NIVEL_ACT = 0
   Exit Sub
End If
If windex <> 1 Then
'SQ_OPER = 1
'PUB_CUENTA = grid_cont.TextMatrix(fila - 1, 2)
'LEER_COM_LLAVE
'If com_llave.EOF Then
'   MsgBox "Cuenta Superior NO Existe, Verificar ...", 48, Pub_Titulo
'   Exit Sub
'End If
End If
NIVEL_ACT = windex
End Sub


Private Sub checostos_Click()
If Posi_Ultimo = -1 Then Exit Sub
If Not grabar.Enabled Then grabar.Enabled = True
If Insertar.Enabled Then
 Insertar.Enabled = False
 Eliminar.Enabled = False
 Cmdcopiar.Enabled = False
 Modo_Acceso = "E"
End If

End Sub

Private Sub chkAfectacion_Click()
If ValidarAfectacion Then
    'Exit Sub
End If
If Not grabar.Enabled Then grabar.Enabled = True
If Insertar.Enabled Then
 Insertar.Enabled = False
 Eliminar.Enabled = False
 Cmdcopiar.Enabled = False
 Modo_Acceso = "E"
End If
End Sub

Private Sub cmdcancelar_Click()
Dim cuenta As Integer
If TreeView1.Nodes.Count = 0 Then
 Exit Sub
End If
If Left(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Text, 1) <> "N" Then
 FORM_CONT.TreeView1.SetFocus
 TreeView1_NodeClick TreeView1.Nodes.Item(TreeView1.SelectedItem.Index)
 Insertar.Enabled = True
 Eliminar.Enabled = True
 Cmdcopiar.Enabled = True
 Exit Sub
End If
TreeView1.Enabled = True ' desbloquea la estructura .
TreeView1.Nodes.Remove TreeView1.SelectedItem.Index
Insertar.Enabled = True
Eliminar.Enabled = True
Cmdcopiar.Enabled = True
FORM_CONT.TreeView1.SetFocus
TreeView1_NodeClick TreeView1.Nodes.Item(TreeView1.SelectedItem.Index)

End Sub

Private Sub cmdcerrar_Click()
Unload FORM_CONT
End Sub

Private Sub cmdCopiar_Click()
Dim WCUENTA As String
If TreeView1.Nodes.Count = 0 Then
 Exit Sub
End If
If Trim(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag) = "RAIZ" Then
 MsgBox " No Procede, en este nivel.", 48, Pub_Titulo
 Exit Sub
End If
WCUENTA = Mid(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key, 2, Len(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key))
COPIA_ESTRU WCUENTA

End Sub


Private Sub comsal_Click()

Dim PS_REP01  As rdoQuery
Dim llave_rep01 As rdoResultset
Screen.MousePointer = 11
pub_cadena = "SELECT * FROM COMSAL WHERE COS_CODCIA = ? AND COS_CUENTA = ? "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
SQ_OPER = 2
PUB_CUENTA = 0
PUB_CODCIA = LK_CODCIA
LEER_COM_LLAVE

Do Until com_mayor.EOF
  SQ_OPER = 3
  PUB_CUENTA = com_mayor!com_cuenta
  PUB_CODCIA = LK_CODCIA
  LEER_COM_LLAVE
  If cos_llave.EOF Then
  DoEvents
  cos_llave.AddNew
  cos_llave!COS_CODCIA = LK_CODCIA
  cos_llave!COS_CUENTA = com_mayor!com_cuenta
  cos_llave!COS_NRO_ANO = Val(Format(LK_FECHA_COP1, "yyyy"))
  
  cos_llave!COS_DEB00 = 0
  cos_llave!COS_HAB00 = 0
  cos_llave!COS_DEB01 = 0
  cos_llave!COS_HAB01 = 0
  cos_llave!COS_DEB02 = 0
  cos_llave!COS_HAB02 = 0
  cos_llave!COS_DEB03 = 0
  cos_llave!COS_HAB03 = 0
  cos_llave!COS_DEB04 = 0
  cos_llave!COS_HAB04 = 0
  cos_llave!COS_DEB05 = 0
  cos_llave!COS_HAB05 = 0
  cos_llave!COS_DEB06 = 0
  cos_llave!COS_HAB06 = 0
  cos_llave!COS_DEB07 = 0
  cos_llave!COS_HAB07 = 0
  cos_llave!COS_DEB08 = 0
  cos_llave!COS_HAB08 = 0
  cos_llave!COS_DEB09 = 0
  cos_llave!COS_HAB09 = 0
  cos_llave!COS_DEB10 = 0
  cos_llave!COS_HAB10 = 0
  cos_llave!COS_DEB11 = 0
  cos_llave!COS_HAB11 = 0
  cos_llave!COS_DEB12 = 0
  cos_llave!COS_HAB12 = 0
  cos_llave.Update
End If

com_mayor.MoveNext
Loop

Screen.MousePointer = 0
MsgBox "TERMINO"

End Sub

Private Sub ctacierre_Click()
PUB_TIPREG = "-55" ' editar cuentas de cierre
PUB_CODCIA = LK_CODCIA
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
Load FrmDatArti
FrmDatArti.Caption = "Cuentas de Cierre Contable "
FrmDatArti.Show 1

End Sub

Private Sub Eliminar_Click()
Dim WS_CUENTA As String * 12
If TreeView1.Nodes.Count = 0 Or NIVEL_ACT = 0 Then
 Exit Sub
End If
WS_CUENTA = Mid(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key, 2, Len(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key))
If TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Children = 0 Then
 pub_mensaje = "Desea Eliminar la Cuenta  " + TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Text + " !!! ...   " + Chr(13) + " ¿Desea Continuar... ?"
 Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
 If Pub_Respuesta = vbNo Then
    Exit Sub
 End If
Else
  MsgBox "NO Procede, existe sub Cuentas", 48, Pub_Titulo
  Exit Sub
End If

PUB_CUENTA = Trim(WS_CUENTA)
PUB_CODCIA = LK_CODCIA
LEER_COM_LLAVE
If com_llave.EOF Then
 MsgBox "Intente Nuevamente . . .", 48, Pub_Titulo
 Exit Sub
End If
com_llave.Delete
TreeView1.Nodes.Remove TreeView1.SelectedItem.Index
FORM_CONT.TreeView1.SetFocus
TreeView1_NodeClick TreeView1.Nodes.Item(TreeView1.SelectedItem.Index)
End Sub
Private Sub Form_Activate()
'If TreeView1.Visible Then TreeView1.SetFocus
End Sub

Private Sub Form_Load()
Dim i
Dim ws_tipo_cta As Integer
Dim WS_SIGNO_D, WS_SIGNO_H As Integer
Dim nodX As Node
Dim wscodcia As String * 2
On Error GoTo SIGUE:
Modo_REICIO = ""
'Dim i As Integer
Modo_Acceso = ""
LIMPIA_DATOS
Posi_Ultimo = -1

If Not cop_llave.EOF Then
For i = 1 To 6
  If cop_llave.rdoColumns(i) <> 0 Then
     wCOM_NIVEL(i) = cop_llave.rdoColumns(i)
     NIVEL_MAX = i
  End If
Next i
Else
 MsgBox "Definir parametros para el plan contable.", 48, Pub_Titulo
 Exit Sub
End If

tipo.Clear
PUB_TIPREG = 16
PUB_CODCIA = "00"
SQ_OPER = 2
LEER_TAB_LLAVE
Do Until tab_mayor.EOF
  tipo.AddItem Nulo_Valors(tab_mayor!tab_nomcorto) & String(50, " ") & tab_mayor!TAB_NUMTAB
  tab_mayor.MoveNext
Loop
tipo.ListIndex = 0
txtnombre.Text = ""
txtautod.Text = ""
txtautoh.Text = ""
txtindice.Text = ""
txtcuenta.Text = ""

pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? ORDER BY COM_CUENTA"
Set PSCOM_CONT = CN.CreateQuery("", pub_cadena)
PSCOM_CONT(0) = 0
Set com_cont = PSCOM_CONT.OpenResultset(rdOpenKeyset, rdConcurValues)
wscodcia = LK_CODCIA
If LK_EMP_PTO = "A" Then
 wscodcia = "00"
End If
PSCOM_CONT(0) = wscodcia
com_cont.Requery
fila = 0
'If com_cont.EOF Then
'Exit Sub
'Else
Set nodX = TreeView1.Nodes.Add(, tvwChild, "TITULO", "Cuentas Principales")
TreeView1.Nodes.Item(1).Tag = "RAIZ"
TreeView1.Indentation = 400
'End If

Do Until com_cont.EOF
   If Trim(com_cont!com_cuenta) = "" Then GoTo otrito
   If com_cont!com_nivel = 1 Then
      WS_SIGNO_D = com_cont!com_signo_d
      WS_SIGNO_H = com_cont!com_signo_h
      ws_tipo_cta = com_cont!com_tipo_cta
   End If
   If WS_SIGNO_D = com_cont!com_signo_d And WS_SIGNO_H = com_cont!com_signo_h And ws_tipo_cta = com_cont!com_tipo_cta Then
   Else
      com_cont.Edit
      com_cont!com_signo_d = WS_SIGNO_D
      com_cont!com_signo_h = WS_SIGNO_H
      com_cont!com_tipo_cta = ws_tipo_cta
      com_cont.Update
   End If
   If Len(Trim(com_cont!com_cuenta)) <> wCOM_NIVEL(com_cont!com_nivel) Then
      i = 1
      Do Until i > 6
         If wCOM_NIVEL(i) = Len(Trim(com_cont!com_cuenta)) Then Exit Do
         i = i + 1
      Loop
      com_cont.Edit
      com_cont!com_nivel = i
      com_cont.Update
   End If
   'If Trim(com_cont!com_cuenta) = "25110" Then Stop
   fila = fila + 1
   If com_cont!com_nivel = 1 Then
      Set nodX = TreeView1.Nodes.Add("TITULO", tvwChild, "A" + Left(com_cont!com_cuenta, wCOM_NIVEL(1)), Trim(com_cont!com_cuenta) + " " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!com_nivel = 2 Then
      Set nodX = TreeView1.Nodes.Add("A" + Left(com_cont!com_cuenta, wCOM_NIVEL(1)), tvwChild, "B" + Left(com_cont!com_cuenta, wCOM_NIVEL(2)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!com_nivel = 3 Then
      Set nodX = TreeView1.Nodes.Add("B" + Left(com_cont!com_cuenta, wCOM_NIVEL(2)), tvwChild, "C" + Left(com_cont!com_cuenta, wCOM_NIVEL(3)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!com_nivel = 4 Then
      Set nodX = TreeView1.Nodes.Add("C" + Left(com_cont!com_cuenta, wCOM_NIVEL(3)), tvwChild, "D" + Left(com_cont!com_cuenta, wCOM_NIVEL(4)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!com_nivel = 5 Then
      Set nodX = TreeView1.Nodes.Add("D" + Left(com_cont!com_cuenta, wCOM_NIVEL(4)), tvwChild, "E" + Left(com_cont!com_cuenta, wCOM_NIVEL(5)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!com_nivel = 6 Then
      Set nodX = TreeView1.Nodes.Add("E" + Left(com_cont!com_cuenta, wCOM_NIVEL(5)), tvwChild, "F" + Left(com_cont!com_cuenta, wCOM_NIVEL(6)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   End If
   'If Trim(com_cont!COM_NIVEL) = "3" Then Stop
   If Trim(com_cont!com_nivel) = "" Then Stop
   TreeView1.Nodes.Item(fila + 1).Tag = Str(com_cont!com_nivel)
''   If Trim(TreeView1.Nodes.Item(fila + 1).Tag) = "" Then Stop
   TreeView1.Nodes.Item(fila + 1).Sorted = True
   WS_CUENTA = com_cont!com_cuenta
   txtcuenta.Text = com_cont!com_cuenta
   txtindice.Text = ""
   If com_cont!com_nivel <> 1 Then
      txtcuenta.Text = Left(Trim(WS_CUENTA), wCOM_NIVEL(com_cont!com_nivel - 1))
      txtindice.Text = Right(Trim(WS_CUENTA), (wCOM_NIVEL(com_cont!com_nivel) - wCOM_NIVEL(com_cont!com_nivel - 1)))
   End If
otrito:
   com_cont.MoveNext
Loop
Posi_Ultimo = 0
NIVEL_ACT = 0
TreeView1.TabIndex = 0
If LK_CODUSU = "ADMIN" Then
 comsal.Visible = True
End If
If Trim(LK_ART_CIAS) <> "" Then
  LISCIA.Visible = True
  LISCIA.Clear
  xcuenta = 0
  For fila = 1 To 30 Step 2
    PUB_CODCIA = Mid(Trim(LK_ART_CIAS), fila, 2)
    If Trim(PUB_CODCIA) = "" Then Exit For
    xcuenta = xcuenta + 1
    PSPAR_MULTI(0) = PUB_CODCIA
    par_multi.Requery
    LISCIA.AddItem PUB_CODCIA & " - " & Trim(par_multi!PAR_NOMBRE)
  Next fila
  For fila = 0 To LISCIA.ListCount - 1
   LISCIA.ListIndex = fila
   If Left(LISCIA.Text, 2) = LK_CODCIA Then LISCIA.Selected(fila) = True
  Next fila
End If



Exit Sub
SIGUE:
 Resume Next
End Sub

Private Sub grabar_Click()
Dim xcuenta As Integer
Dim ING
Dim wsdescri As String
ING = 0
If txtnombre.Text = "" Then
 MsgBox "Cuenta Necesita su descripción ", 48, Pub_Titulo
 txtnombre.SetFocus
 Exit Sub
End If
If txtindice.Visible And txtindice.Text = "" Then
 MsgBox "Cuenta Necesita Sub Indice . ", 48, Pub_Titulo
 txtindice.SetFocus
 Exit Sub
End If
If txtindice.Visible And txtindice.MaxLength <> Len(Trim(txtindice.Text)) Then
 MsgBox "Indice Invalido , Longitud no es correcto. ", 48, Pub_Titulo
 Azul txtindice, txtindice
 Exit Sub
End If
VERI_NIVEL (Trim(txtcuenta.Text) & Trim(txtindice.Text))
If NIVEL_ACT = 0 Then
  MsgBox "Nivel 0 Intente Nuevamente .. Carge las Cuentas. . .", 48, Pub_Titulo
  Exit Sub
End If
'SI EXISTEN LAS CUENTAS AUTOMATICAS ..
If Trim(txtautod.Text) <> "" Then
 If Not BUSCA_AUTO(txtautod.Text) Then
   MsgBox "Cuenta Automatica del Debe No Existe o es Invalida ..", 48, Pub_Titulo
   Azul txtautod, txtautod
   Exit Sub
 End If
End If
If Trim(txtautoh.Text) <> "" Then
 If Not BUSCA_AUTO(txtautoh.Text) Then
   MsgBox "Cuenta Automatica del Haber No Existe o es Invalida ..", 48, Pub_Titulo
   Azul txtautoh, txtautoh
   Exit Sub
 End If
End If
If Trim(txtautod2.Text) <> "" Then
 If Not BUSCA_AUTO(txtautod2.Text) Then
   MsgBox "Cuenta Automatica del Debe No Existe o es Invalida ..", 48, Pub_Titulo
   Azul txtautod2, txtautod2
   Exit Sub
 End If
End If
If Trim(txtautod3.Text) <> "" Then
 If Not BUSCA_AUTO(txtautod3.Text) Then
   MsgBox "Cuenta Automatica del Debe No Existe o es Invalida ..", 48, Pub_Titulo
   Azul txtautod3, txtautod3
   Exit Sub
 End If
End If
If Trim(txtautod4.Text) <> "" Then
 If Not BUSCA_AUTO(txtautod4.Text) Then
   MsgBox "Cuenta Automatica del Debe No Existe o es Invalida ..", 48, Pub_Titulo
   Azul txtautod4, txtautod4
   Exit Sub
 End If
End If
If Trim(txtautod5.Text) <> "" Then
 If Not BUSCA_AUTO(txtautod5.Text) Then
   MsgBox "Cuenta Automatica del Debe No Existe o es Invalida ..", 48, Pub_Titulo
   Azul txtautod5, txtautod5
   Exit Sub
 End If
End If



If Trim(tipo.Text) = "" Then
   MsgBox "Seleccione Tipo de Cuenta ", 48, Pub_Titulo
   If tipo.Enabled Then tipo.SetFocus
   Exit Sub
End If
If Trim(SIGNOH.Text) = "" Then
   MsgBox "Seleccione Signo del Haber ", 48, Pub_Titulo
   SIGNOH.SetFocus
   Exit Sub
End If
If Trim(SIGNOD.Text) = "" Then
   MsgBox "Seleccione Signo del debe ", 48, Pub_Titulo
   SIGNOD.SetFocus
   Exit Sub
End If
If consis() = False Then
  Exit Sub
End If

If NIVEL_ACT = 1 Then
 PUB_CUENTA = Trim(txtcuenta.Text)
Else
 PUB_CUENTA = Trim(txtcuenta.Text) & Trim(txtindice.Text)
End If
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_COM_LLAVE
If Not com_llave.EOF And Modo_Acceso = "I" Then
 MsgBox "Cuenta contable ya Existe . . .", 48, Pub_Titulo
 If NIVEL_ACT = 1 Then
  Azul txtcuenta, txtcuenta
 Else
  Azul txtindice, txtindice
 End If
 Exit Sub
End If
pb.Max = 3
pb.Min = 0
pb.Value = 0
pb.Visible = True
DoEvents
grabar.Enabled = False
pb.Value = pb.Value + 1
If com_llave.EOF Then
    PUB_CODCIA = LK_CODCIA
    If Trim(LK_ART_CIAS) <> "" And LK_EMP <> "PIU" Then
      xcuenta = 1
      For fila = 1 To 30
          ws_codcia = Mid(Trim(LK_ART_CIAS), xcuenta, 2)
          If Trim(ws_codcia) = "" Then Exit For
          PUB_CODCIA = ws_codcia
          GoSub GRABAR_CON
          CREAR_SALDOS PUB_CUENTA
          xcuenta = xcuenta + 2
      Next fila
    Else
      GoSub GRABAR_CON
      CREAR_SALDOS PUB_CUENTA
    End If
    Modo_REICIO = "A"
    TreeView1.Enabled = True ' desbloquea la estructura .
Else
    ING = 1
      If Trim(LK_ART_CIAS) <> "" And LK_EMP <> "PIU" Then
      xcuenta = 1
      For fila = 1 To 30
          ws_codcia = Mid(Trim(LK_ART_CIAS), xcuenta, 2)
          If Trim(ws_codcia) = "" Then Exit For
          PUB_CODCIA = ws_codcia
          SQ_OPER = 1
          LEER_COM_LLAVE
          GoSub ACTU
          CREAR_SALDOS PUB_CUENTA
          xcuenta = xcuenta + 2
      Next fila
    Else
      GoSub ACTU
      CREAR_SALDOS PUB_CUENTA
    End If
    
    If NIVEL_ACT = 1 And Modo_Acceso = "E" Then ACTUALIZA_CUENTAS txtcuenta.Text, Val(Right(tipo.Text, 3)), Val(SIGNOD.Text), Val(SIGNOH.Text)
End If
pb.Value = pb.Value + 1
If Modo_Acceso = "E" Then
wsdescri = PUB_CUENTA + "  " + Trim(txtnombre.Text)
TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Text = wsdescri
ElseIf Modo_Acceso = "I" Then
 wsdescri = PUB_CUENTA + "  " + Trim(txtnombre.Text)
 TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Text = wsdescri
 TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key = TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key + PUB_CUENTA
 TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag = NIVEL_ACT
 TreeView1.Nodes.Item(Posi_Ultimo).Sorted = True
End If
TreeView1_NodeClick TreeView1.Nodes.Item(TreeView1.SelectedItem.Index)
pb.Value = pb.Value + 1
pb.Visible = False
If NIVEL_ACT = 1 And Modo_Acceso = "E" Then MsgBox "Actualizados la Cuenta y los niveles siguientes. ", 48, Pub_Titulo
Eliminar.Enabled = True
Cmdcopiar.Enabled = True
Insertar.Enabled = True
TreeView1.SetFocus
Exit Sub

fin:

Exit Sub

GRABAR_CON:
com_llave.AddNew
    com_llave!COM_CODCIA = PUB_CODCIA
    com_llave!com_cuenta = PUB_CUENTA
    com_llave!com_DESCRIPCION = Trim(txtnombre.Text)
    com_llave!com_nivel = NIVEL_ACT
    com_llave!com_cuenta_sup = Trim(txtcuenta.Text)
    com_llave!com_cuenta_AUTO_H = Trim(txtautoh.Text)
    com_llave!com_cuenta_AUTOM_D = Trim(txtautod.Text)
    com_llave!COM_CUENTA_AUTOM_D2 = Trim(txtautod2.Text)
    com_llave!COM_CUENTA_AUTOM_D3 = Trim(txtautod3.Text)
    com_llave!COM_CUENTA_AUTOM_D4 = Trim(txtautod4.Text)
    com_llave!COM_CUENTA_AUTOM_D5 = Trim(txtautod5.Text)
    com_llave!COM_POR_AUTOM_D = Val(txtautodp.Text)
    com_llave!COM_POR_AUTOM_D2 = Val(txtautodp2.Text)
    com_llave!COM_POR_AUTOM_D3 = Val(txtautodp3.Text)
    com_llave!COM_POR_AUTOM_D4 = Val(txtautodp4.Text)
    com_llave!COM_POR_AUTOM_D5 = Val(txtautodp5.Text)
    com_llave!com_flag_afectacion = chkAfectacion.Value
    If NIVEL_ACT = NIVEL_MAX Then
       com_llave!com_flag_afectacion = "1"
    Else
       com_llave!com_flag_afectacion = "0"
    End If
    com_llave!com_ESTADO = ""
    com_llave!COM_DEB_ANO = 0
    com_llave!COM_HAB_ANO = 0
    com_llave!COM_DEB_MES = 0
    com_llave!COM_HAB_MES = 0
    com_llave!com_signo_d = 0
    com_llave!com_signo_h = 0
    com_llave!com_ACT_PAS = 0
    com_llave!com_signo_d = SIGNOD.Text
    com_llave!com_signo_h = SIGNOH.Text
    com_llave!com_tipo_cta = Right(tipo, 3)
    If checostos.Value = 0 Then
       com_llave!COM_CENTRO_COSTOS = " "
    Else
       com_llave!COM_CENTRO_COSTOS = "A"
    End If
   
    com_llave.Update

Return

ACTU:
com_llave.Edit
com_llave!com_DESCRIPCION = txtnombre.Text
com_llave!com_cuenta_AUTOM_D = txtautod.Text
com_llave!com_cuenta_AUTO_H = txtautoh.Text
com_llave!com_tipo_cta = Right(tipo, 3)
com_llave!com_signo_d = SIGNOD.Text
com_llave!com_signo_h = SIGNOH.Text
com_llave!COM_CUENTA_AUTOM_D2 = Trim(txtautod2.Text)
com_llave!COM_CUENTA_AUTOM_D3 = Trim(txtautod3.Text)
com_llave!COM_CUENTA_AUTOM_D4 = Trim(txtautod4.Text)
com_llave!COM_CUENTA_AUTOM_D5 = Trim(txtautod5.Text)
com_llave!com_flag_afectacion = chkAfectacion.Value
com_llave!COM_POR_AUTOM_D = Val(txtautodp.Text)
com_llave!COM_POR_AUTOM_D2 = Val(txtautodp2.Text)
com_llave!COM_POR_AUTOM_D3 = Val(txtautodp3.Text)
com_llave!COM_POR_AUTOM_D4 = Val(txtautodp4.Text)
com_llave!COM_POR_AUTOM_D5 = Val(txtautodp5.Text)
If checostos.Value = 0 Then
   com_llave!COM_CENTRO_COSTOS = " "
Else
   com_llave!COM_CENTRO_COSTOS = "A"
End If
com_llave.Update
Return
   
End Sub
Private Sub Insertar_Click()
Dim WCUENTA As String
If TreeView1.Nodes.Count = 0 Then
 Exit Sub
End If
Posi_Ultimo = TreeView1.SelectedItem.Index
Modo_Acceso = "I"
NIVEL_ACT = NIVEL_ACT + 1
If NIVEL_ACT > NIVEL_MAX Then
 MsgBox "No Procede, llego al Ultimo Nivel. ", 48, Pub_Titulo
 Exit Sub
End If
If chkAfectacion.Value = 1 Then
    MsgBox "No Procede esta Cuenta tiene Afectación. Modifiquela y vuelva a intentar", vbInformation, Pub_Titulo
    Exit Sub
End If
Insertar.Enabled = False
grabar.Enabled = True
Eliminar.Enabled = False
Cmdcopiar.Enabled = False
WCUENTA = Mid(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key, 2, Len(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key))
If Trim(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag) = "RAIZ" Then
 TreeView1.Nodes.Add "TITULO", tvwChild, "A", "Nueva Cuenta..."
End If
If Val(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag) = 1 Then
   TreeView1.Nodes.Add TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key, tvwChild, "B", "Nueva Cuenta..."
ElseIf Val(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag) = 2 Then
   TreeView1.Nodes.Add TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key, tvwChild, "C", "Nueva Cuenta..."
ElseIf Val(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag) = 3 Then
   TreeView1.Nodes.Add TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key, tvwChild, "D", "Nueva Cuenta..."
ElseIf Val(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag) = 4 Then
   TreeView1.Nodes.Add TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key, tvwChild, "E", "Nueva Cuenta..."
ElseIf Val(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag) = 5 Then
   TreeView1.Nodes.Add TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key, tvwChild, "F", "Nueva Cuenta..."
ElseIf Val(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag) = 6 Then
   TreeView1.Nodes.Add TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key, tvwChild, "G", "Nueva Cuenta..."
End If
If WCUENTA = "ITULO" Then
Else
 txtcuenta.Text = Left(WCUENTA, wCOM_NIVEL(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag))
End If
Set TreeView1.SelectedItem = TreeView1.Nodes.Item(TreeView1.Nodes.Count)
TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).EnsureVisible
If NIVEL_ACT = 1 Then
  fradestino.Visible = True
  SIGNOS.Visible = True
  cuentas.Visible = True
  fraTipo.Visible = True
  tipo.Enabled = True
  SIGNOD.Enabled = True
  SIGNOH.Enabled = True
  'txtautod.Visible = False
  'txtautoh.Visible = False
  txtindice.Visible = False
  txtcuenta.Visible = True
  txtcuenta.Enabled = True
  VAR1 = wCOM_NIVEL(NIVEL_ACT)
  txtcuenta.MaxLength = VAR1
  txtcuenta.Text = ""
  txtcuenta.SetFocus
Else
    'txtautod.Visible = False
    'txtautoh.Visible = False
    txtcuenta.Enabled = False
    txtindice.Enabled = True
    txtindice.Visible = True
    VAR1 = wCOM_NIVEL(NIVEL_ACT) - Len(txtcuenta.Text)
    txtindice.MaxLength = VAR1
    txtindice.Text = ""
    txtindice.SetFocus
End If
txtnombre.Text = "Nueva Cuenta . . . "
If NIVEL_ACT <> 1 Then
  tipo.Enabled = False
  SIGNOD.Enabled = False
  SIGNOH.Enabled = False
End If
If NIVEL_ACT = NIVEL_MAX Then
  tipo.Enabled = False
  fradestino.Visible = True
  'txtautod.Visible = True
  'txtautoh.Visible = True
  SIGNOH.Visible = True
  SIGNOD.Enabled = False
  SIGNOH.Enabled = False
End If
TreeView1.Enabled = False ' bloquea la estructura .
Exit Sub
End Sub

Private Sub LISCIA_Click()
If FORM_CONT.TreeView1.Visible Then
  FORM_CONT.TreeView1.SetFocus
 ' TreeView1_NodeClick TreeView1.Nodes.Item(TreeView1.SelectedItem.Index)
End If

End Sub

Private Sub salacu_Click()
TreeView1.SetFocus
TreeView1_NodeClick TreeView1.Nodes.Item(TreeView1.SelectedItem.Index)

End Sub

Private Sub SIGNOD_Click()
If Posi_Ultimo = -1 Then Exit Sub
If Not grabar.Enabled Then grabar.Enabled = True
If Insertar.Enabled Then
 Insertar.Enabled = False
 Eliminar.Enabled = False
 Cmdcopiar.Enabled = False
 Modo_Acceso = "E"
End If

End Sub

Private Sub SIGNOD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If SIGNOH.Visible And SIGNOH.Enabled Then SIGNOH.SetFocus: Exit Sub
End If

End Sub

Private Sub SIGNOH_Click()
If Posi_Ultimo = -1 Then Exit Sub
If Not grabar.Enabled Then grabar.Enabled = True
If Insertar.Enabled Then
 Insertar.Enabled = False
 Eliminar.Enabled = False
 Cmdcopiar.Enabled = False
 Modo_Acceso = "E"
End If

End Sub

Private Sub SIGNOH_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If txtautod.Visible And txtautod.Enabled Then txtautod.SetFocus: Exit Sub
  If txtautoh.Visible And txtautoh.Enabled Then txtautoh.SetFocus: Exit Sub
  If grabar.Visible And grabar.Enabled Then grabar.SetFocus: Exit Sub
End If

End Sub

Private Sub tipo_Click()
If Posi_Ultimo = -1 Then Exit Sub
If Not grabar.Enabled Then grabar.Enabled = True
If Insertar.Enabled Then
 Insertar.Enabled = False
 Eliminar.Enabled = False
 Cmdcopiar.Enabled = False
 Modo_Acceso = "E"
End If
End Sub

Private Sub tipo_GotFocus()
CARAC = tipo.Text

End Sub

Private Sub tipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If SIGNOD.Visible And SIGNOD.Enabled Then SIGNOD.SetFocus: Exit Sub
  If txtautod.Visible And txtautod.Enabled Then txtautod.SetFocus: Exit Sub
End If

End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
Cancel = True   ' Se cancela la operación
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
'On Error GoTo SALE
Dim WCUENTA As String
Dim windex As String
If LLENA_CIASEL(LISCIA) = 9 Then Exit Sub
grabar.Enabled = False
LIMPIA_DATOS
If Trim(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag) = "RAIZ" Then
 NIVEL_ACT = 0
 GoTo fin
End If
SQ_OPER = 1
PUB_CUENTA = ""
If InStr(1, Node.Text, " ") <> 0 Then
 PUB_CUENTA = Left(Node.Text, InStr(1, Node.Text, " ") - 1)
End If
PUB_CODCIA = LK_CODCIA
LEER_COM_LLAVE
If com_llave.EOF Then
 GoTo fin:
 Exit Sub
End If
txtcuenta.MaxLength = 0
txtindice.MaxLength = 0
NIVEL_ACT = com_llave!com_nivel
WCUENTA = Mid(Node.Key, 2, Len(Node.Key))
txtcuenta.Enabled = False
txtindice.Visible = False
txtcuenta.Text = WCUENTA
Posi_Ultimo = -1
txtnombre.Text = Trim(com_llave!com_DESCRIPCION)
tipo.ListIndex = com_llave!com_tipo_cta - 1

If com_llave!com_signo_d = 1 Then
  SIGNOD.ListIndex = 0
Else
  SIGNOD.ListIndex = 1
End If
If com_llave!com_signo_h = 1 Then
  SIGNOH.ListIndex = 0
Else
  SIGNOH.ListIndex = 1
End If

If Val(com_llave!com_nivel) = 1 Then ' devuelve el Nivel
    tipo.Enabled = True
    SIGNOS.Enabled = True
    SIGNOD.Enabled = True
    SIGNOH.Enabled = True
Else
   If tipo.Enabled = True Then tipo.Enabled = False
   SIGNOD.Enabled = False
   SIGNOH.Enabled = False
End If
chkAfectacion.Value = com_llave!com_flag_afectacion
If Val(com_llave!com_flag_afectacion) = 1 Then ' cuando el nivel es el maximo
   'If Not txtautod.Visible Then txtautod.Visible = True
   'If Not txtautoh.Visible Then txtautoh.Visible = True
   fradestino.Visible = True
   checostos.Visible = True
   txtautoh.Text = Trim(Nulo_Valors(com_llave!com_cuenta_AUTO_H))
   txtautod.Text = Trim(Nulo_Valors(com_llave!com_cuenta_AUTOM_D))
   txtautod2.Text = Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D2)
   txtautod3.Text = Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D3)
   txtautod4.Text = Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D4)
   txtautod5.Text = Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D5)
   txtautodp.Text = Nulo_Valor0(com_llave!COM_POR_AUTOM_D)
   txtautodp2.Text = Nulo_Valor0(com_llave!COM_POR_AUTOM_D2)
   txtautodp3.Text = Nulo_Valor0(com_llave!COM_POR_AUTOM_D3)
   txtautodp4.Text = Nulo_Valor0(com_llave!COM_POR_AUTOM_D4)
   txtautodp5.Text = Nulo_Valor0(com_llave!COM_POR_AUTOM_D5)
   If Nulo_Valors(com_llave!COM_CENTRO_COSTOS) = "A" Then
       checostos.Value = 1
   Else
       checostos.Value = 0
   End If
Else
   fradestino.Visible = False
   checostos.Visible = False
   fradestino.Visible = False
End If
Posi_Ultimo = 0
fin:
If Val(Node.Tag) = 0 Then ' No ha Seleccionado nada
  fradestino.Visible = False
  SIGNOS.Visible = False
  cuentas.Visible = False
  fraTipo.Visible = False
   Exit Sub
Else
  'fradestino.Visible = True
  SIGNOS.Visible = True
  cuentas.Visible = True
  fraTipo.Visible = True
End If

If salacu.Value = 0 Then
 
 JALA_SALDO PUB_CUENTA, 0
Else
 JALA_SALDO PUB_CUENTA, 1
End If
'saldos.Caption = Format(((Val(com_llave!COM_DEB_ANO) + Val(com_llave!COM_DEB_MES)) * com_llave!com_SIGNO_D) + ((Val(com_llave!COM_HAB_ANO) + Val(com_llave!COM_HAB_MES)) * com_llave!com_SIGNO_H), "#,###.00")
saldos.Caption = Format(((PUB_IMPORTE_DEB) * com_llave!com_signo_d) + ((PUB_IMPORTE_HAB) * com_llave!com_signo_h), "#,###.00")

TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Sorted = True
Insertar.Enabled = True
Eliminar.Enabled = True
Cmdcopiar.Enabled = True
Exit Sub
SALE:
MsgBox "Verificar tipo de Cuentas", 48, Pub_Titulo
Resume Next
End Sub

Private Sub txtautod_Change()
If Posi_Ultimo = -1 Then Exit Sub
If Not grabar.Enabled Then grabar.Enabled = True
If Insertar.Enabled Then
 Insertar.Enabled = False
 Eliminar.Enabled = False
 Cmdcopiar.Enabled = False
 Modo_Acceso = "E"
End If
End Sub

Private Sub txtautod_GotFocus()
Azul txtautod, txtautod
CARAC = txtautod.Text

End Sub

Private Sub txtautod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If grabar.Visible And grabar.Enabled Then grabar.SetFocus: Exit Sub
  'If txtautoh.Visible And txtautoh.Enabled Then txtautoh.SetFocus: Exit Sub
End If

End Sub

Private Sub txtautod2_Change()
If Posi_Ultimo = -1 Then Exit Sub
If Not grabar.Enabled Then grabar.Enabled = True
If Insertar.Enabled Then
 Insertar.Enabled = False
 Eliminar.Enabled = False
 Cmdcopiar.Enabled = False
 Modo_Acceso = "E"
End If

End Sub

Private Sub txtautod2_GotFocus()
Azul txtautod2, txtautod2
CARAC = txtautod2.Text

End Sub

Private Sub txtautod3_Change()
If Posi_Ultimo = -1 Then Exit Sub
If Not grabar.Enabled Then grabar.Enabled = True
If Insertar.Enabled Then
 Insertar.Enabled = False
 Eliminar.Enabled = False
 Cmdcopiar.Enabled = False
 Modo_Acceso = "E"
End If

End Sub

Private Sub txtautod3_GotFocus()
Azul txtautod3, txtautod3
CARAC = txtautod3.Text

End Sub

Private Sub txtautod4_Change()
If Posi_Ultimo = -1 Then Exit Sub
If Not grabar.Enabled Then grabav.Enabled = True
If Insertar.Enabled Then
 Insertar.Enabled = False
 Eliminar.Enabled = False
 Cmdcopiar.Enabled = False
 Modo_Acceso = "E"
End If

End Sub

Private Sub txtautod4_GotFocus()
Azul txtautod4, txtautod4
CARAC = txtautod4.Text

End Sub

Private Sub txtautod5_Change()
If Posi_Ultimo = -1 Then Exit Sub
If Not grabar.Enabled Then grabar.Enabled = True
If Insertar.Enabled Then
 Insertar.Enabled = False
 Eliminar.Enabled = False
 Cmdcopiar.Enabled = False
 Modo_Acceso = "E"
End If

End Sub

Private Sub txtautod5_GotFocus()
Azul txtautod5, txtautod5
CARAC = txtautod5.Text

End Sub

Private Sub txtautodp_Change()
If Posi_Ultimo = -1 Then Exit Sub
If Not grabar.Enabled Then grabar.Enabled = True
If Insertar.Enabled Then
 Insertar.Enabled = False
 Eliminar.Enabled = False
 Cmdcopiar.Enabled = False
 Modo_Acceso = "E"
End If

End Sub

Private Sub txtautodp_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
End Sub

Private Sub txtautodp2_Change()
If Posi_Ultimo = -1 Then Exit Sub
If Not grabar.Enabled Then grabar.Enabled = True
If Insertar.Enabled Then
 Insertar.Enabled = False
 Eliminar.Enabled = False
 Cmdcopiar.Enabled = False
 Modo_Acceso = "E"
End If

End Sub

Private Sub txtautodp2_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
End Sub

Private Sub txtautodp3_Change()
If Posi_Ultimo = -1 Then Exit Sub
If Not grabar.Enabled Then grabar.Enabled = True
If Insertar.Enabled Then
 Insertar.Enabled = False
 Eliminar.Enabled = False
 Cmdcopiar.Enabled = False
 Modo_Acceso = "E"
End If

End Sub

Private Sub txtautodp3_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
End Sub

Private Sub txtautodp4_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
End Sub

Private Sub txtautodp5_Change()
If Posi_Ultimo = -1 Then Exit Sub
If Not grabar.Enabled Then grabar.Enabled = True
If Insertar.Enabled Then
 Insertar.Enabled = False
 Eliminar.Enabled = False
 Cmdcopiar.Enabled = False
 Modo_Acceso = "E"
End If

End Sub

Private Sub txtautodp5_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
End Sub

Private Sub txtautoh_Change()
If Posi_Ultimo = -1 Then Exit Sub
If Not grabar.Enabled Then grabar.Enabled = True
If Insertar.Enabled Then
 Insertar.Enabled = False
 Eliminar.Enabled = False
 Cmdcopiar.Enabled = False
 Modo_Acceso = "E"
End If

End Sub

Private Sub txtautoh_GotFocus()
Azul txautoh, txtautoh
CARAC = txtautoh.Text
End Sub

Private Sub txtautoh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul txtautod, txtautod
End If

End Sub

Private Sub txtcuenta_GotFocus()
Azul txtcuenta, txtcuenta
End Sub

Private Sub txtcuenta_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
'grabar.Visible = True
If KeyAscii = 13 Then
  If txtindice.Visible And txtindice.Enabled Then txtindice.SetFocus: Exit Sub
  If txtnombre.Visible And txtnombre.Enabled Then txtnombre.SetFocus: Exit Sub
End If

End Sub

Private Sub txtindice_GotFocus()
Azul txtindice, txtindice

End Sub

Private Sub txtindice_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 If txtnombre.Visible And txtnombre.Enabled Then txtnombre.SetFocus
End If
'grabar.Visible = True
End Sub

Private Sub txtnombre_Change()
If Posi_Ultimo = -1 Then Exit Sub
If Not grabar.Enabled Then grabar.Enabled = True
If Insertar.Enabled Then
 Insertar.Enabled = False
 Eliminar.Enabled = False
 Cmdcopiar.Enabled = False
 Modo_Acceso = "E"
End If
End Sub

Private Sub txtnombre_GotFocus()
Azul txtnombre, txtnombre
CARAC = txtnombre.Text
'grabar.Visible = True
End Sub

Private Sub txtnombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  cmdcancelar_Click
  TreeView1.SetFocus
  Exit Sub
End If
If KeyAscii = 13 Then
   If tipo.Enabled Then tipo.SetFocus: Exit Sub
   If grabar.Enabled Then grabar.SetFocus: Exit Sub
End If

End Sub

Public Sub LIMPIA_DATOS()
Posi_Ultimo = -1
txtcuenta.Text = ""
txtindice.Text = ""
txtautod.Text = ""
txtautod2.Text = ""
txtautod3.Text = ""
txtautod4.Text = ""
txtautod5.Text = ""

txtautoh.Text = ""
txtautodp.Text = ""
txtautodp2.Text = ""
txtautodp3.Text = ""
txtautodp4.Text = ""
txtautodp5.Text = ""
txtnombre.Text = ""
tipo.ListIndex = -1
SIGNOD.ListIndex = -1
SIGNOH.ListIndex = -1
Posi_Ultimo = 0
saldos.Caption = ""
checostos.Value = 0
End Sub

Public Sub LLENA_CUENTAS()
Dim wscodcia As String * 2
wscodcia = LK_CODCIA
If LK_EMP_PTO = "A" Then
 wscodcia = "00"
End If

PSCOM_CONT(0) = wscodcia
com_cont.Requery
fila = 0
If com_cont.EOF Then Exit Sub
TreeView1.Nodes.Clear
Set nodX = TreeView1.Nodes.Add(, tvwChild, "TITULO", "Cuentas Principales")
TreeView1.Nodes.Item(1).Tag = "RAIZ"
TreeView1.Indentation = 400
Do Until com_cont.EOF
   fila = fila + 1
   If com_cont!com_nivel = 1 Then
      Set nodX = TreeView1.Nodes.Add("TITULO", tvwChild, "A" + Left(com_cont!com_cuenta, wCOM_NIVEL(1)), Trim(com_cont!com_cuenta) + " " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!com_nivel = 2 Then
      Set nodX = TreeView1.Nodes.Add("A" + Left(com_cont!com_cuenta, wCOM_NIVEL(1)), tvwChild, "B" + Left(com_cont!com_cuenta, wCOM_NIVEL(2)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!com_nivel = 3 Then
      Set nodX = TreeView1.Nodes.Add("B" + Left(com_cont!com_cuenta, wCOM_NIVEL(2)), tvwChild, "C" + Left(com_cont!com_cuenta, wCOM_NIVEL(3)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!com_nivel = 4 Then
      Set nodX = TreeView1.Nodes.Add("C" + Left(com_cont!com_cuenta, wCOM_NIVEL(3)), tvwChild, "D" + Left(com_cont!com_cuenta, wCOM_NIVEL(4)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!com_nivel = 5 Then
      Set nodX = TreeView1.Nodes.Add("D" + Left(com_cont!com_cuenta, wCOM_NIVEL(4)), tvwChild, "E" + Left(com_cont!com_cuenta, wCOM_NIVEL(5)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!com_nivel = 6 Then
      Set nodX = TreeView1.Nodes.Add("E" + Left(com_cont!com_cuenta, wCOM_NIVEL(5)), tvwChild, "F" + Left(com_cont!com_cuenta, wCOM_NIVEL(6)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   End If
   TreeView1.Nodes.Item(fila + 1).Tag = Str(com_cont!com_nivel)
   WS_CUENTA = com_cont!com_cuenta
   txtcuenta.Text = com_cont!com_cuenta
   txtindice.Text = ""
   If com_cont!com_nivel <> 1 Then
      txtcuenta.Text = Left(Trim(WS_CUENTA), wCOM_NIVEL(com_cont!com_nivel - 1))
      txtindice.Text = Right(Trim(WS_CUENTA), (wCOM_NIVEL(com_cont!com_nivel) - wCOM_NIVEL(com_cont!com_nivel - 1)))
   End If
   com_cont.MoveNext
Loop
Posi_Ultimo = 0
NIVEL_ACT = 0
Exit Sub

End Sub

Public Sub ACTUALIZA_CUENTAS(WCUENTAS As String, WTIPO As Integer, wsignod As Integer, wsignoh As Integer)
Dim wcuentas2 As String
Dim wscodcia As String * 2

pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? and COM_CUENTA > ?  ORDER BY COM_CUENTA"
Set PSCOM_MAYOR2 = CN.CreateQuery("", pub_cadena)
PSCOM_MAYOR2(0) = 0
PSCOM_MAYOR2(1) = 0
Set com_mayor2 = PSCOM_MAYOR2.OpenResultset(rdOpenKeyset, rdConcurValues)
PSCOM_MAYOR2(0) = LK_CODCIA
PSCOM_MAYOR2(1) = WCUENTAS
com_mayor2.Requery
If com_mayor2.EOF Then
  Exit Sub
End If
Do Until com_mayor2.EOF
   If com_mayor2!com_nivel = 1 Then Exit Do
   com_mayor2.Edit
   com_mayor2!com_tipo_cta = WTIPO
   com_mayor2!com_signo_d = wsignod
   com_mayor2!com_signo_h = wsignoh
   com_mayor2.Update
   com_mayor2.MoveNext
Loop

End Sub

Public Sub COPIA_ESTRU(WCUENTAS As String)
Dim wcuentas2 As String
Dim Wflag As String * 1
Dim WCUENTA As Integer
Dim valor
Dim WULTIMO As Integer
valor = InputBox("Ingrese la Cuenta de Nivel Reemplasante, segun el nivel donde se encuentra  :", Pub_Titulo, "")
If valor = "" Then
  Screen.MousePointer = 0
  Exit Sub
End If
If Not IsNumeric(valor) Then
  MsgBox "Cuenta No Procede. ", 48, Pub_Titulo
  Exit Sub
End If
If NIVEL_ACT = 1 Then
  If Len(valor) <> Len(Left(WCUENTAS, wCOM_NIVEL(NIVEL_ACT))) Then
    MsgBox "Cuenta No es correcta. ", 48, Pub_Titulo
    Exit Sub
  End If
Else
  If Left(valor, wCOM_NIVEL(NIVEL_ACT - 1)) <> Left(WCUENTAS, wCOM_NIVEL(NIVEL_ACT - 1)) Then
    MsgBox "Cuenta No es correcta. ", 48, Pub_Titulo
    Exit Sub
  End If
End If

pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? and COM_CUENTA >= ?  ORDER BY COM_CUENTA"
Set PSCOM_MAYOR2 = CN.CreateQuery("", pub_cadena)
PSCOM_MAYOR2(0) = 0
PSCOM_MAYOR2(1) = 0
Set com_mayor2 = PSCOM_MAYOR2.OpenResultset(rdOpenKeyset, rdConcurValues)
PSCOM_MAYOR2(0) = LK_CODCIA
PSCOM_MAYOR2(1) = WCUENTAS
com_mayor2.Requery
WCUENTA = 0
If com_mayor2.EOF Then
  MsgBox "Intente Nuevamente...", 48, Pub_Titulo
  Exit Sub
End If
'wfijo = Left(valor, wCOM_NIVEL(com_mayor2!COM_NIVEL - 1))
'windice = Mid(com_mayor2!com_cuenta, Len(Trim(wfijo)) + 1, Len(com_mayor2!com_cuenta))
SQ_OPER = 1
PUB_CUENTA = valor
PUB_CODCIA = LK_CODCIA
LEER_COM_LLAVE
If Not com_llave.EOF Then
 MsgBox "Cuenta Existe en Plan Contable .", 48, Pub_Titulo
 Exit Sub
End If
pb.Min = 0
pb.Max = 3
pb.Value = 0
pb.Visible = True
DoEvents
Wflag = ""
WULTIMO = -1
pb.Value = pb.Value + 1
Do Until com_mayor2.EOF
   If NIVEL_ACT = com_mayor2!com_nivel Then WCUENTA = WCUENTA + 1
   If WCUENTA = 2 Then Exit Do
   If NIVEL_ACT > com_mayor2!com_nivel Then Exit Do
   If Trim(Wflag) = "" Then
    Wflag = "A"
    WFIJO = Left(valor, wCOM_NIVEL(NIVEL_ACT))
    windice = ""
   Else
    WFIJO = Left(valor, wCOM_NIVEL(com_mayor2!com_nivel - 1))
    windice = Mid(com_mayor2!com_cuenta, Len(Trim(WFIJO)) + 1, Len(com_mayor2!com_cuenta))
   End If
'  MsgBox wfijo & windice
'  GoTo dale
   com_llave.AddNew
   com_llave!com_cuenta = WFIJO + windice
   com_llave!com_tipo_cta = com_mayor2!com_tipo_cta
   com_llave!COM_CODCIA = LK_CODCIA
   com_llave!com_DESCRIPCION = com_mayor2!com_DESCRIPCION
   com_llave!com_nivel = com_mayor2!com_nivel
   com_llave!com_cuenta_sup = Left(Trim(com_llave!com_cuenta), wCOM_NIVEL(com_llave!com_nivel - 1))
   com_llave!com_cuenta_AUTOM_D = ""
   com_llave!com_cuenta_AUTO_H = ""
   If com_mayor2!com_nivel = NIVEL_MAX Then
       com_llave!com_flag_afectacion = "1"
   Else
       com_llave!com_flag_afectacion = "0"
   End If
   com_llave!com_ESTADO = ""
   com_llave!COM_DEB_ANO = 0
   com_llave!COM_HAB_ANO = 0
   com_llave!COM_DEB_MES = 0
   com_llave!COM_HAB_MES = 0
   com_llave!com_signo_d = com_mayor2!com_signo_d
   com_llave!com_signo_h = com_mayor2!com_signo_h
   com_llave!com_ACT_PAS = 0
   com_llave!com_tipo_cta = com_mayor2!com_tipo_cta
   com_llave.Update
   If com_mayor2!com_nivel = 1 Then
     TreeView1.Nodes.Add "TITULO", tvwChild, "A", "Nueva Cuenta..."
   ElseIf com_mayor2!com_nivel = 2 Then
    TreeView1.Nodes.Add TreeView1.Nodes.Item("A" + Left(WFIJO + windice, wCOM_NIVEL(com_mayor2!com_nivel - 1))).Key, tvwChild, "B", "Nueva Cuenta..."
   ElseIf com_mayor2!com_nivel = 3 Then
    TreeView1.Nodes.Add TreeView1.Nodes.Item("B" + Left(WFIJO + windice, wCOM_NIVEL(com_mayor2!com_nivel - 1))).Key, tvwChild, "C", "Nueva Cuenta..."
   ElseIf com_mayor2!com_nivel = 4 Then
    TreeView1.Nodes.Add TreeView1.Nodes.Item("C" + Left(WFIJO + windice, wCOM_NIVEL(com_mayor2!com_nivel - 1))).Key, tvwChild, "D", "Nueva Cuenta..."
   ElseIf com_mayor2!com_nivel = 5 Then
    TreeView1.Nodes.Add TreeView1.Nodes.Item("D" + Left(WFIJO + windice, wCOM_NIVEL(com_mayor2!com_nivel - 1))).Key, tvwChild, "E", "Nueva Cuenta..."
   ElseIf com_mayor2!com_nivel = 6 Then
    TreeView1.Nodes.Add TreeView1.Nodes.Item("E" + Left(WFIJO + windice, wCOM_NIVEL(com_mayor2!com_nivel - 1))).Key, tvwChild, "F", "Nueva Cuenta..."
   End If
   Set TreeView1.SelectedItem = TreeView1.Nodes.Item(TreeView1.Nodes.Count)
   TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).EnsureVisible
   TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Text = Trim(WFIJO + windice) + "  " + Trim(com_mayor2!com_DESCRIPCION)
   TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key = TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key + Trim(WFIJO + windice)
   TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag = com_mayor2!com_nivel
   If WULTIMO = -1 Then
     WULTIMO = TreeView1.SelectedItem.Index
   End If
   com_mayor2.MoveNext
Loop
pb.Value = pb.Value + 1
TreeView1.Nodes.Item(1).Sorted = True
If WULTIMO <> -1 Then
 Set TreeView1.SelectedItem = TreeView1.Nodes.Item(WULTIMO)
 TreeView1_NodeClick TreeView1.Nodes.Item(TreeView1.SelectedItem.Index)
End If
pb.Value = pb.Value + 1
pf.Visible = False
MsgBox "Proceso de Copiado Terminado.", 48, Pub_Titulo
End Sub

Public Function consis() As Boolean
Dim suma As Currency
If fradestino.Visible = False Then
  consis = True
  Exit Function
End If
If Trim(txtautod.Text) <> "" Then
 ' If Val(txtautodp.Text) = 0 Then
 '   MsgBox "Verificar Falta Ingresar Datos ...", 48, Pub_Titulo
 '   txtautodp.SetFocus
 '   GoTo falso
 ' End If
End If
If Val(txtautodp.Text) <> 0 Then
  If Trim(txtautod.Text) = "" Then
    MsgBox "Verificar Falta Ingresar Datos ...", 48, Pub_Titulo
    txtautod.SetFocus
    GoTo falso
  End If
End If

If Trim(txtautod2.Text) <> "" Then
'  If Val(txtautodp2.Text) = 0 Then
'     MsgBox "Verificar Falta Ingresar Datos ...", 48, Pub_Titulo
'     txtautodp2.SetFocus
'     GoTo falso
'  End If
End If

If Val(txtautodp2.Text) <> 0 Then
  If Trim(txtautod2.Text) = "" Then
    MsgBox "Verificar Falta Ingresar Datos ...", 48, Pub_Titulo
    txtautod2.SetFocus
    GoTo falso
  End If
End If

If Trim(txtautod3.Text) <> "" Then
 ' If Val(txtautodp3.Text) = 0 Then
 '    MsgBox "Verificar Falta Ingresar Datos ...", 48, Pub_Titulo
 '    txtautodp3.SetFocus
 '    GoTo falso
 ' End If
End If
If Val(txtautodp3.Text) <> 0 Then
  If Trim(txtautod3.Text) = "" Then
   MsgBox "Verificar Falta Ingresar Datos ...", 48, Pub_Titulo
    txtautod3.SetFocus
    GoTo falso
  End If
End If

If Trim(txtautod4.Text) <> "" Then
 ' If Val(txtautodp4.Text) = 0 Then
 '    MsgBox "Verificar Falta Ingresar Datos ...", 48, Pub_Titulo
 '    txtautodp4.SetFocus
 '    GoTo falso
 ' End If
End If
If Val(txtautodp4.Text) <> 0 Then
  If Trim(txtautod4.Text) = "" Then
   MsgBox "Verificar Falta Ingresar Datos ...", 48, Pub_Titulo
    txtautod4.SetFocus
    GoTo falso
  End If
End If

If Trim(txtautod5.Text) <> "" Then
 ' If Val(txtautodp5.Text) = 0 Then
 '    MsgBox "Verificar Falta Ingresar Datos ...", 48, Pub_Titulo
 '    GoTo falso
 ' End If
End If
If Val(txtautodp5.Text) <> 0 Then
  If Trim(txtautod5.Text) = "" Then
    MsgBox "Verificar Falta Ingresar Datos ...", 48, Pub_Titulo
    txtautod5.SetFocus
    GoTo falso
  End If
End If

suma = Val(txtautodp.Text) + Val(txtautodp2.Text) + Val(txtautodp3.Text) + Val(txtautodp4.Text) + Val(txtautodp5.Text)
If Trim(txtautoh.Text) <> "" Then
 If suma = 0 Then
   MsgBox "Ingresar porcentajes...", 48, Pub_Titulo
   Azul txtautodp, txtautodp
   GoTo falso
 End If
End If
If suma = 100 Or suma = 0 Then
Else
   MsgBox "Verificar porcentajes NO esta proporcionado ...", 48, Pub_Titulo
   txtautodp.SetFocus
   GoTo falso
End If

consis = True
Exit Function
falso:
consis = False
End Function

Public Sub CREAR_SALDOS(WCUENTA As String)
  SQ_OPER = 3
  PUB_CUENTA = WCUENTA
  LEER_COM_LLAVE
  If cos_llave.EOF Then
    cos_llave.AddNew
    cos_llave!COS_CODCIA = PUB_CODCIA
    cos_llave!COS_CUENTA = WCUENTA
    cos_llave!COS_NRO_ANO = Val(Format(LK_FECHA_COP1, "yyyy"))
    cos_llave!COS_DEB00 = 0
    cos_llave!COS_HAB00 = 0
    cos_llave!COS_DEB01 = 0
    cos_llave!COS_HAB01 = 0
    cos_llave!COS_DEB02 = 0
    cos_llave!COS_HAB02 = 0
    cos_llave!COS_DEB03 = 0
    cos_llave!COS_HAB03 = 0
    cos_llave!COS_DEB04 = 0
    cos_llave!COS_HAB04 = 0
    cos_llave!COS_DEB05 = 0
    cos_llave!COS_HAB05 = 0
    cos_llave!COS_DEB06 = 0
    cos_llave!COS_HAB06 = 0
    cos_llave!COS_DEB07 = 0
    cos_llave!COS_HAB07 = 0
    cos_llave!COS_DEB08 = 0
    cos_llave!COS_HAB08 = 0
    cos_llave!COS_DEB09 = 0
    cos_llave!COS_HAB09 = 0
    cos_llave!COS_DEB10 = 0
    cos_llave!COS_HAB10 = 0
    cos_llave!COS_DEB11 = 0
    cos_llave!COS_HAB11 = 0
    cos_llave!COS_DEB12 = 0
    cos_llave!COS_HAB12 = 0
    cos_llave.Update
End If
End Sub
Private Function ValidarAfectacion() As Boolean
Dim iNivel As Integer

    ValidarAfectacion = True
    If chkAfectacion.Value = 1 Then
        iNivel = NIVEL_ACT + 1
        archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = '" & LK_CODCIA & "' AND COM_CUENTA LIKE '" & Trim(PUB_CUENTA) & "%' AND COM_NIVEL = " & iNivel
        Set PSX = CN.CreateQuery("", archi)
        Set X = PSX.OpenResultset(rdOpenKeyset)
        X.Requery
        If Not X.EOF Then
            ValidarAfectacion = False
            MsgBox "Esta operacion es invalida. Esta cuenta tiene cuentas de Nivel Inferior", vbInformation, Pub_Titulo
            chkAfectacion.Value = 0
            Exit Function
        End If
    End If
End Function
