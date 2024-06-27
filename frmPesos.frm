VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmPesos 
   Caption         =   "Transfer. Import. (Recepción de Información)"
   ClientHeight    =   6900
   ClientLeft      =   555
   ClientTop       =   1845
   ClientWidth     =   5310
   Icon            =   "frmPesos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fracabe 
      Caption         =   "Opcion :"
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   5175
      Begin VB.OptionButton openvio 
         Caption         =   "Transfer. de Ventas"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   17
         Top             =   480
         Width           =   2415
      End
      Begin VB.OptionButton openvio 
         Caption         =   "Transfer Recepción de Mercaderia"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   16
         Top             =   240
         Width           =   3375
      End
      Begin VB.OptionButton openvio 
         Caption         =   "Transfer. Ajustes de Mercaderia"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   15
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label7 
         Caption         =   "Proceso :"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Archivo: "
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lbltransfer 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1080
         TabIndex        =   19
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label lblarchi 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1080
         TabIndex        =   18
         Top             =   960
         Width           =   3975
      End
   End
   Begin ComctlLib.ProgressBar barra 
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Anular datos Importados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      Picture         =   "frmPesos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Solo ver los Articulos."
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3480
      TabIndex        =   6
      Top             =   4200
      Width           =   495
   End
   Begin VB.Data VENTAS 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CommandButton flopy 
      Caption         =   "Importar Transfer."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Picture         =   "frmPesos.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
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
      Height          =   615
      Left            =   1920
      Picture         =   "frmPesos.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label lblpro 
      Caption         =   "Procesando Información ..."
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H008B4914&
      Caption         =   "Solution for Business"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   6600
      Width           =   5295
   End
   Begin VB.Label Label6 
      Caption         =   "Descripción de Articulo en Base Transfer :"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   4695
   End
   Begin VB.Label Label5 
      Caption         =   "Descripción de Articulo en Base del Servidor :"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Digitar SI = es correcto, de lo contrario NO : "
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   3840
      Width           =   5175
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   5295
   End
End
Attribute VB_Name = "frmPesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TF_TIPMOV As Integer

Private Sub Command1_Click()
Unload frmPesos
End Sub

Private Sub Command2_Click()
pub_mensaje = "Anular Pase !!! ...   ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
Screen.MousePointer = 11
pub_cadena = "DELETE FACART WHERE FAR_TIPMOV = " & TF_TIPMOV & " AND FAR_CODCIA = '" & LK_CODCIA & "'"
CN.Execute pub_cadena, rdExecDirect
Screen.MousePointer = 0
MsgBox "Pase de la Información del Transfer. ha sido Anulado ", 48, Pub_Titulo

End Sub

Private Sub flopy_Click()
'On Error GoTo ABORTA_PROCESO
Dim xl As Object
Dim FLAG_ARTI As String * 1
Dim FLAG_CLI As String * 1

Dim PSNUC_REPO As rdoQuery
Dim PSNUC_REPO2 As rdoQuery
Dim PSNUC_REPO3 As rdoQuery

Dim PUB_ALTERNO As String
Dim WS_NOMCLI As String
Dim WS_NOMART As String
Dim AVISO As Integer
Dim AVISO2 As Integer
Dim nuc_repo  As rdoResultset
Dim nuc_repo2  As rdoResultset
Dim nuc_repo3  As rdoResultset

Dim f1 As Integer
Dim flag_activo As Integer
Dim PSART_REPO As rdoQuery
Dim ART_REPO  As rdoResultset
Dim PSfar_menorx As rdoQuery
Dim PSFAR As rdoQuery
Dim far_r As rdoResultset
Dim far_menorx As rdoResultset
Dim WCONFIG As String
If openvio(0).Value Then
 TF_TIPMOV = 26
ElseIf openvio(1).Value Then
 TF_TIPMOV = 27
ElseIf openvio(2).Value Then
 TF_TIPMOV = 28
End If
GEN.Requery
If Trim(Nulo_Valors(GEN!gen_ciaimp)) = "" Then
  MsgBox "El codigo de la Compañia de la Recepcion no Validad... Cia : " & Trim(Nulo_Valors(GEN!gen_ciaimp)), 48, Pub_Titulo
  Exit Sub
End If
If LK_CODCIA <> Trim(Nulo_Valors(GEN!gen_ciaimp)) Then
   MsgBox "El Codigo de esta Compañia no es validad para el proceso... " & Chr(13) & "Compañia validad : " & Trim(Nulo_Valors(GEN!gen_ciaimp)), 48, Pub_Titulo
   GoTo fin
End If

VENTAS.DatabaseName = lblarchi.Caption '  Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\Standar\TAjustes.mdb"
PUB_IMPORTE = 0
WCONFIG = InputBox("- Digitar 1 para fechas con formato ##/##/#### " & Chr(13) & "- Digitar 2 para fechas con formato ##/##/## ", "Cofiguración del TRANSFER", 2)
If WCONFIG = "" Then
  Exit Sub
End If
If val(WCONFIG) = 1 Then
  VENTAS.RecordSource = "SELECT DISTINCT FAR_FBG, FAR_RUC, FAR_DNI, FAR_NOMCLI , FAR_DIREC FROM FACART WHERE FAR_FECHA = '" & Format(LK_FECHA_DIA, "dd/mm/yyyy") & "'"
ElseIf val(WCONFIG) = 2 Then
  VENTAS.RecordSource = "SELECT DISTINCT FAR_FBG, FAR_RUC, FAR_DNI, FAR_NOMCLI , FAR_DIREC FROM FACART WHERE FAR_FECHA = '" & Format(LK_FECHA_DIA, "dd/mm/yy") & "'"
Else
  MsgBox "Fuera de rango ...", 48, Pub_Titulo
  Exit Sub
End If

VENTAS.Refresh
If VENTAS.Recordset.EOF Then
  Screen.MousePointer = 0
  MsgBox "VERIFICAR ...NO HAY DATOS EN EL ARCHIVO : " & Trim(lblarchi.Caption) & Chr(13) & "Posible Soluciones :" & Chr(13) & "El formato de fecha no es correcta  0  la fecha del dia de la compañia no es la misma del archivo transfer. O el archivo esta vacio o dañado.", 48, Pub_Titulo
  VENTAS.RecordSource = "SELECT * FROM FACART "
  VENTAS.Refresh
  If val(WCONFIG) = 1 Then
    If Not VENTAS.Recordset.EOF Then MsgBox "Datos el Transfer son del día : " & VENTAS.Recordset!FAR_fecha & Chr(13) & "y la fecha de la compañia es del dia : " & Format(LK_FECHA_DIA, "dd/mm/yyyy"), vbInformation, Pub_Titulo
  Else
    If Not VENTAS.Recordset.EOF Then MsgBox "Datos el Transfer son del día : " & VENTAS.Recordset!FAR_fecha & Chr(13) & "y la fecha de la compañia es del dia : " & Format(LK_FECHA_DIA, "dd/mm/yy"), vbInformation, Pub_Titulo
  End If
  Exit Sub
End If
GoSub WEXCEL

Label1.Visible = True
DoEvents
pub_cadena = "SELECT CLI_NOMBRE, CLI_CODCLIE FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_RUC_ESPOSO = ? AND CLI_CP = 'C' "
Set PSNUC_REPO = CN.CreateQuery("", pub_cadena)
PSNUC_REPO(0) = 0
PSNUC_REPO(1) = 0
Set nuc_repo = PSNUC_REPO.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT CLI_NOMBRE, CLI_CODCLIE FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_RUC_ESPOSA = ? "
Set PSNUC_REPO2 = CN.CreateQuery("", pub_cadena)
PSNUC_REPO2(0) = 0
PSNUC_REPO2(1) = 0
Set nuc_repo2 = PSNUC_REPO2.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT CLI_NOMBRE, CLI_CODCLIE FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_NOMBRE = ? "
Set PSNUC_REPO3 = CN.CreateQuery("", pub_cadena)
PSNUC_REPO3(0) = 0
PSNUC_REPO3(1) = 0
Set nuc_repo3 = PSNUC_REPO3.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


If Check1.Value = 1 Then
   '  ** Mustra todos los productos que vienen antes de procesar. **
    pub_cadena = "SELECT * FROM  ARTI WHERE ART_CODCIA = ? ORDER BY ART_ALTERNO"
    Set PSART_REPO = CN.CreateQuery("", pub_cadena)
    PSART_REPO(0) = 0
    Set ART_REPO = PSART_REPO.OpenResultset(rdOpenKeyset, rdConcurValues)
    PSART_REPO.rdoParameters(0) = LK_CODCIA
    ART_REPO.Requery
    Do Until ART_REPO.EOF
      If Trim(PUB_ALTERNO) = Trim(ART_REPO!ART_ALTERNO) Then
        f1 = f1 + 1
        xl.Cells(f1, 2) = ART_REPO!ART_ALTERNO
      End If
      PUB_ALTERNO = ART_REPO!ART_ALTERNO
      ART_REPO.MoveNext
    Loop
    GoTo VAMOS

End If

pub_cadena = "SELECT ART_NOMBRE , ART_ALTERNO, ART_KEY , ART_CP FROM ARTI WHERE ART_CODCIA = ? AND ART_ALTERNO = ? "
Set PSART_REPO = CN.CreateQuery("", pub_cadena)
PSART_REPO(0) = 0
PSART_REPO(1) = 0
Set ART_REPO = PSART_REPO.OpenResultset(rdOpenKeyset, rdConcurValues)
barra.Visible = True
lblpro.Visible = True
barra.Min = 0
barra.max = VENTAS.Recordset.RecordCount + 1
PUB_NUMGUIA = 0
Label1.Caption = "Buscando Clientes inexistentes..."
DoEvents

f1 = 1
FLAG_CLI = ""
f1 = f1 + 1
xl.Cells(f1, 1) = "SE AGREGARON ESTOS CLIENTES :"
 
Do Until VENTAS.Recordset.EOF

   If VENTAS.Recordset!far_fbg = "F" Then
      PSNUC_REPO.rdoParameters(0) = LK_CODCIA
      PSNUC_REPO.rdoParameters(1) = Trim(Nulo_Valors(VENTAS.Recordset!FAR_RUC))
      nuc_repo.Requery
      If nuc_repo.EOF Then
         GoSub GENERA_CLI
         f1 = f1 + 1
         xl.Cells(f1, 1) = Trim(VENTAS.Recordset!FAR_NOMCLI)
         xl.Cells(f1, 2) = "RUC=" & Trim(VENTAS.Recordset!FAR_RUC)
         xl.Cells(f1, 3) = Trim(VENTAS.Recordset!FAR_DIREC)
         FLAG_CLI = "A"
      End If
   End If
   
   If VENTAS.Recordset!far_fbg = "B" Then
      If Trim(Nulo_Valors(VENTAS.Recordset!FAR_DNI)) = "" Then
         PSNUC_REPO3.rdoParameters(0) = LK_CODCIA
         PSNUC_REPO3.rdoParameters(1) = Trim(Nulo_Valors(VENTAS.Recordset!FAR_NOMCLI))
         nuc_repo3.Requery
         If nuc_repo3.EOF Then
            GoSub GENERA_CLI
            f1 = f1 + 1
            xl.Cells(f1, 1) = Trim(VENTAS.Recordset!FAR_NOMCLI)
            xl.Cells(f1, 2) = "DNI=" & Trim(VENTAS.Recordset!FAR_DNI)
            xl.Cells(f1, 3) = Trim(VENTAS.Recordset!FAR_DIREC)
            FLAG_CLI = "A"
            'crear cliente
         End If
      Else
      PSNUC_REPO2.rdoParameters(0) = LK_CODCIA
      PSNUC_REPO2.rdoParameters(1) = Trim(Nulo_Valors(VENTAS.Recordset!FAR_DNI))
      nuc_repo2.Requery
      If nuc_repo2.EOF Then
         GoSub GENERA_CLI
         f1 = f1 + 1
         xl.Cells(f1, 1) = Trim(VENTAS.Recordset!FAR_NOMCLI)
         xl.Cells(f1, 2) = "DNI=" & Trim(VENTAS.Recordset!FAR_DNI)
         xl.Cells(f1, 3) = Trim(VENTAS.Recordset!FAR_DIREC)
         FLAG_CLI = "A"
         ' crear cliente
      End If
   End If
   End If
   
   VENTAS.Recordset.MoveNext
   PUB_NUMGUIA = PUB_NUMGUIA + 1
   barra.Value = PUB_NUMGUIA
Loop

Label1.Caption = "Buscando Articulos inexistentes..."
DoEvents

PUB_NUMGUIA = 0
If val(WCONFIG) = 1 Then
 VENTAS.RecordSource = "SELECT DISTINCT  FAR_ALTERNO, FAR_NOMART FROM FACART WHERE FAR_FECHA= '" & Format(LK_FECHA_DIA, "dd/mm/yyyy") & "'"
Else
 VENTAS.RecordSource = "SELECT DISTINCT  FAR_ALTERNO, FAR_NOMART FROM FACART WHERE FAR_FECHA= '" & Format(LK_FECHA_DIA, "dd/mm/yy") & "'"
End If
VENTAS.Refresh
barra.Min = 0
barra.max = VENTAS.Recordset.RecordCount + 1
If f1 > 1 Then f1 = f1 + 3
FLAG_ARTI = ""
Do Until VENTAS.Recordset.EOF
   PSART_REPO.rdoParameters(0) = LK_CODCIA
   PSART_REPO.rdoParameters(1) = Trim(VENTAS.Recordset!far_ALTERNO)
   ART_REPO.Requery
   If ART_REPO.EOF Then
      f1 = f1 + 1
      xl.Cells(f1, 2) = Trim(VENTAS.Recordset!far_nomart)
      xl.Cells(f1, 3) = Trim(VENTAS.Recordset!far_ALTERNO)
      FLAG_ARTI = "A"
   Else
      If Trim(ART_REPO!art_nombre) <> Trim(VENTAS.Recordset!far_nomart) Then
         If ART_REPO!ART_CP = "X" Then GoTo sal
         Label2.Caption = Trim(ART_REPO!art_nombre)
         Label3.Caption = Trim(VENTAS.Recordset!far_nomart)
         Text1.Text = ""
         Do Until Text1.Text = "SI" Or Text1.Text = "NO"
         ' DEJALO QUE PREGUNTE
            'Frame1.Visible = True
            Label2.Visible = True
            Label3.Visible = True
            Text1.SetFocus
            DoEvents
         Loop
         If Text1.Text = "NO" Then
            f1 = f1 + 1
            xl.Cells(f1, 2) = "REVISAR COD. VIENE:" & VENTAS.Recordset!far_ALTERNO & " " & VENTAS.Recordset!far_nomart
            FLAG_ARTI = "A"
         Else
            ART_REPO.Edit
            ART_REPO!ART_CP = "X"
            ART_REPO.Update
         End If
      End If
   End If
sal:
   VENTAS.Recordset.MoveNext
   PUB_NUMGUIA = PUB_NUMGUIA + 1
   barra.Value = PUB_NUMGUIA
Loop

If FLAG_ARTI = "A" Then
   GoTo VAMOS
End If

If val(WCONFIG) = 1 Then
 VENTAS.RecordSource = "SELECT * FROM FACART WHERE FAR_FECHA = '" & Format(LK_FECHA_DIA, "dd/mm/yyyy") & "'" & " ORDER BY FAR_FBG, FAR_RUC, FAR_DNI"
Else
 VENTAS.RecordSource = "SELECT * FROM FACART WHERE FAR_FECHA = '" & Format(LK_FECHA_DIA, "dd/mm/yy") & "'" & " ORDER BY FAR_FBG, FAR_RUC, FAR_DNI"
End If
VENTAS.Refresh

barra.Min = 0
barra.max = VENTAS.Recordset.RecordCount + 1
PUB_NUMGUIA = 0
Label1.Caption = "Actualizando codigos de Clientes ..."
DoEvents
WS_NOMCLI = ""
f1 = 1
Do Until VENTAS.Recordset.EOF
   PUB_CODCLIE = 0
   If VENTAS.Recordset!far_fbg = "F" Then
      PSNUC_REPO.rdoParameters(0) = LK_CODCIA
      PSNUC_REPO.rdoParameters(1) = Trim(Nulo_Valors(VENTAS.Recordset!FAR_RUC))
      nuc_repo.Requery
      PUB_CODCLIE = nuc_repo!cli_codclie
   End If
   If VENTAS.Recordset!far_fbg = "B" Then
      If Trim(Nulo_Valors(VENTAS.Recordset!FAR_DNI)) = "" Then
         PSNUC_REPO3.rdoParameters(0) = LK_CODCIA
         PSNUC_REPO3.rdoParameters(1) = Trim(Nulo_Valors(VENTAS.Recordset!FAR_NOMCLI))
         nuc_repo3.Requery
         PUB_CODCLIE = nuc_repo3!cli_codclie
      Else
         PSNUC_REPO2.rdoParameters(0) = LK_CODCIA
         PSNUC_REPO2.rdoParameters(1) = Trim(Nulo_Valors(VENTAS.Recordset!FAR_DNI))
         nuc_repo2.Requery
         PUB_CODCLIE = nuc_repo2!cli_codclie
      End If
   Else
      PSNUC_REPO.rdoParameters(0) = LK_CODCIA
      PSNUC_REPO.rdoParameters(1) = Trim(Nulo_Valors(VENTAS.Recordset!FAR_RUC))
      nuc_repo.Requery
      If Not nuc_repo.EOF Then
        PUB_CODCLIE = nuc_repo!cli_codclie
      End If
      
   End If
   
      
   VENTAS.Recordset.Edit
   VENTAS.Recordset!far_codclie = PUB_CODCLIE
   VENTAS.Recordset.Update
   VENTAS.Recordset.MoveNext
   PUB_NUMGUIA = PUB_NUMGUIA + 1
   barra.Value = PUB_NUMGUIA
Loop

If val(WCONFIG) = 1 Then
  VENTAS.RecordSource = "SELECT * FROM FACART WHERE FAR_FECHA = '" & Format(LK_FECHA_DIA, "dd/mm/yyyy") & "'" & " ORDER BY FAR_NOMART"
Else
  VENTAS.RecordSource = "SELECT * FROM FACART WHERE FAR_FECHA = '" & Format(LK_FECHA_DIA, "dd/mm/yy") & "'" & " ORDER BY FAR_NOMART"
End If
VENTAS.Refresh

barra.Min = 0
barra.max = VENTAS.Recordset.RecordCount + 1
PUB_NUMGUIA = 0
Label1.Caption = "Actualizando codigos de Articulos ..."
DoEvents
WS_NOMCLI = ""
f1 = 1
Do Until VENTAS.Recordset.EOF
       SQ_OPER = 3
       pu_alterno = VENTAS.Recordset!far_ALTERNO
       pu_codcia = LK_CODCIA
       LEER_ART_LLAVE
       If art_llave_alt.EOF Then MsgBox "El Codigo del Articulo No existe en base de datos : " & (VENTAS.Recordset!far_ALTERNO) & " " & Trim(VENTAS.Recordset!far_nomart)
       VENTAS.Recordset.Edit
       VENTAS.Recordset!far_codart = art_llave_alt!ART_KEY
       VENTAS.Recordset.Update
   
   VENTAS.Recordset.MoveNext
   PUB_NUMGUIA = PUB_NUMGUIA + 1
   barra.Value = PUB_NUMGUIA
Loop


pub_cadena = "SELECT * FROM facart WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_FBG , FAR_NUMSER, FAR_NUMFAC"
Set PSfar_menorx = CN.CreateQuery("", pub_cadena)
PSfar_menorx(0) = 0
PSfar_menorx(1) = 0
Set far_menorx = PSfar_menorx.OpenResultset(rdOpenKeyset, rdConcurValues)
PSfar_menorx.rdoParameters(0) = LK_CODCIA
PSfar_menorx.rdoParameters(1) = TF_TIPMOV
far_menorx.Requery
Do Until far_menorx.EOF
       If far_menorx!far_transito = "X" And far_menorx!FAR_fecha < LK_FECHA_DIA Then
          far_menorx.Delete
          GoTo mas
       End If
       If far_menorx!far_transito = "X" And far_menorx!FAR_fecha = LK_FECHA_DIA Then
          AVISO = 1
       End If
       If far_menorx!far_transito = "T" And far_menorx!FAR_fecha = LK_FECHA_DIA Then
          AVISO2 = 1
       End If
mas:
      far_menorx.MoveNext
Loop

If AVISO2 = 1 And AVISO = 1 Then
   MsgBox "No procede falta terminar proceso... Procese la Transferencia. ", 48, Pub_Titulo
   GoTo fin
End If

If AVISO = 1 Then
   MsgBox "Ya proceso los datos del día ", 48, Pub_Titulo
   GoTo fin
End If


pub_cadena = "DELETE FACART  WHERE FAR_TIPMOV = " & TF_TIPMOV & " AND FAR_CODCIA = '" & LK_CODCIA & "'"
CN.Execute pub_cadena, rdExecDirect

far_menorx.Requery




Screen.MousePointer = 11
DoEvents
If val(WCONFIG) = 1 Then
 VENTAS.RecordSource = "SELECT * FROM FACART WHERE FAR_FECHA =  '" & Format(LK_FECHA_DIA, "dd/mm/yyyy") & " '   ORDER BY FAR_FBG, FAR_NUMSER, FAR_NUMFAC"
Else
 VENTAS.RecordSource = "SELECT * FROM FACART WHERE FAR_FECHA =  '" & Format(LK_FECHA_DIA, "dd/mm/yy") & " '   ORDER BY FAR_FBG, FAR_NUMSER, FAR_NUMFAC"
End If
VENTAS.Refresh


barra.Min = 0
barra.max = VENTAS.Recordset.RecordCount + 1
PUB_NUMGUIA = 0

PUB_NUMFAC = VENTAS.Recordset!far_numfac
PUB_NUMSER = VENTAS.Recordset!far_numser
PUB_FBG = VENTAS.Recordset!far_fbg

CN.Execute "Begin Transaction", rdExecDirect
PUB_NUMGUIA = 0
  Do Until VENTAS.Recordset.EOF
repite:
      If PUB_NUMFAC <> VENTAS.Recordset!far_numfac Then PUB_NUMFAC = PUB_NUMFAC + 1
         
      If PUB_FBG = VENTAS.Recordset!far_fbg And PUB_NUMSER = val(VENTAS.Recordset!far_numser) Then
         If PUB_NUMFAC <> VENTAS.Recordset!far_numfac Then
             If LK_CODTRA = 2401 Then
                MsgBox "Documento. Anulado..." & PUB_FBG & "/. " & PUB_NUMSER & " - " & PUB_NUMFAC, 48, Pub_Titulo
             End If
             GoTo repite
         End If
      Else
         PUB_NUMFAC = VENTAS.Recordset!far_numfac
         PUB_NUMSER = VENTAS.Recordset!far_numser
         PUB_FBG = VENTAS.Recordset!far_fbg
      End If
      
      GoSub registra
      PUB_NUMGUIA = PUB_NUMGUIA + 1
      barra.Value = PUB_NUMGUIA
      VENTAS.Recordset.MoveNext
  Loop
  
CN.Execute "Commit Transaction", rdExecDirect

Screen.MousePointer = 0
MsgBox "PROCESO TERMINADO.. Ahora puede activar el proceso automatico", 48, Pub_Titulo

VAMOS:
xl.DisplayAlerts = False
xl.APPLICATION.Visible = True
Set xl = Nothing
Screen.MousePointer = 0

Unload frmPesos

Exit Sub
  

registra:
      far_menorx.AddNew
      far_menorx!FAR_TIPMOV = TF_TIPMOV
      far_menorx!FAR_CODCIA = VENTAS.Recordset!FAR_CODCIA
      far_menorx!far_numser = VENTAS.Recordset!far_numser
      far_menorx!far_numfac = VENTAS.Recordset!far_numfac
      far_menorx!FAR_NUMSEC = VENTAS.Recordset!FAR_NUMSEC
      far_menorx!far_numguia = VENTAS.Recordset!far_codart
      far_menorx!far_codart = 0
      far_menorx!FAR_STOCK = 0
      far_menorx!far_cantidad = VENTAS.Recordset!far_cantidad
      far_menorx!FAR_fecha = VENTAS.Recordset!FAR_fecha
      far_menorx!far_signo_arm = VENTAS.Recordset!far_signo_arm
      far_menorx!far_signo_car = VENTAS.Recordset!far_signo_car
      far_menorx!far_codclie = 0
      far_menorx!FAR_NUMDOC = VENTAS.Recordset!far_codclie
      far_menorx!FAR_cp = "C"
      far_menorx!far_transito = "T"
      far_menorx!far_estado = "N"
      far_menorx!FAR_COSPRO = 0
      far_menorx!FAR_PRECIO = VENTAS.Recordset!FAR_PRECIO
      far_menorx!far_fbg = VENTAS.Recordset!far_fbg
      far_menorx!far_IMPTO = VENTAS.Recordset!far_IMPTO
      far_menorx!FAR_SUBTOTAL = VENTAS.Recordset!FAR_SUBTOTAL
      far_menorx!FAR_DESCTO = 0
      far_menorx!FAR_TOT_DESCTO = 0
      far_menorx!FAR_GASTOS = 0
      far_menorx!FAR_BRUTO = VENTAS.Recordset!FAR_BRUTO
      far_menorx!FAR_pordescto1 = VENTAS.Recordset!FAR_pordescto1
      far_menorx!FAR_tipo_cambio = 0
      far_menorx!FAR_TOT_DESCTO = 0
      far_menorx!FAR_DESCTO = 0
      far_menorx!FAR_DIAS = VENTAS.Recordset!FAR_DIAS
      far_menorx!FAR_TIPDOC = VENTAS.Recordset!FAR_TIPDOC
      far_menorx!FAR_fecha = VENTAS.Recordset!FAR_fecha
      far_menorx!FAR_NUMSER_C = VENTAS.Recordset!far_serguia
      far_menorx!FAR_NUMFAC_C = VENTAS.Recordset!far_numguia
      far_menorx!FAR_NUMOPER = PUB_NUMGUIA
      far_menorx!far_descri = VENTAS.Recordset!far_descri
      far_menorx!far_PESO = VENTAS.Recordset!far_PESO
      
      far_menorx!FAR_PEDSER = 0
      far_menorx!FAR_PEDFAC = 0
      
      far_menorx!FAR_STOCK = 0
      far_menorx!FAR_TOT_DESCTO = 0
      far_menorx!FAR_STOCK = 0
      far_menorx!FAR_GASTOS = 0
      far_menorx!FAR_equiv = 0
      far_menorx!FAR_tipo_cambio = 0
      far_menorx!far_precio_neto = 0
      far_menorx!FAR_MONEDA = VENTAS.Recordset!FAR_MONEDA
      
      far_menorx!far_otra_cia = " "
      far_menorx!FAR_cp = "C"
      far_menorx!FAR_fecha_compra = LK_FECHA_DIA
      far_menorx!FAR_CODVEN = VENTAS.Recordset!FAR_CODVEN
      far_menorx!FAR_NUM_LOTE = VENTAS.Recordset!FAR_SECUENCIA
      far_menorx.Update
Return

  
WEXCEL:
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  PUB_CLAVE = ""
  xl.Workbooks.Open Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\STANDAR\NOTAS.xls", 0, True, 4

Return
  
  
  
SALE:
Screen.MousePointer = 0
MsgBox Err.Description, 48, Pub_Titulo
Exit Sub

ABORTA_PROCESO:
CN.Execute "Rollback Transaction", rdExecDirect
MsgBox "Se produjo un error ...Reintente ..."
End

fin:
Unload frmTImport


Exit Sub

GENERA_CLI:

PUB_CODCLIE = GENERA_CODI
cli_llave.AddNew
On Error GoTo GEN_CLI
    cli_llave!CLI_CP = "C"
    cli_llave!cli_codclie = PUB_CODCLIE
    cli_llave!cli_SALDO = 0
    cli_llave!CLI_DET_TOT = "D"
    cli_llave!CLI_MONEDA = " "
    cli_llave!CLI_CODCIA = LK_CODCIA
    cli_llave!CLI_NOMBRE_ESPOSO = Trim(VENTAS.Recordset!FAR_NOMCLI)
    cli_llave!CLI_NOMBRE_ESPOSA = " "
    cli_llave!CLI_NOMBRE_EMPRESA = " "
    cli_llave!CLI_NOMBRE = Trim(VENTAS.Recordset!FAR_NOMCLI)
    cli_llave!CLI_CASA_DIREC = Trim(VENTAS.Recordset!FAR_DIREC)
    cli_llave!CLI_CASA_NUM = 0
    cli_llave!CLI_CASA_ZONA = 48
    cli_llave!CLI_LUGAR_CASA = 1
    cli_llave!CLI_LUGAR_TRAB = 1
    cli_llave!CLI_CASA_SUBZONA = 23
    cli_llave!CLI_ZONA_NEW = 169
    cli_llave!CLI_TRAB_DIREC = Trim(VENTAS.Recordset!FAR_DIREC)
    cli_llave!CLI_TRAB_NUM = 0
    cli_llave!cli_TRAB_ZONA = 48
    cli_llave!cli_TRAB_SUBZONA = 23
    cli_llave!cli_ruc_esposo = Trim(VENTAS.Recordset!FAR_RUC)
    cli_llave!cli_RUC_ESPOSA = Trim(VENTAS.Recordset!FAR_DNI)
    cli_llave!CLI_RUC_EMPRESA = " "
    cli_llave!CLI_CASA1 = " "
    cli_llave!CLI_CASA2 = " "
    cli_llave!CLI_REGPUB1 = " "
    cli_llave!CLI_REGPUB2 = " "
    cli_llave!CLI_AUTOAVALUO = " "
    cli_llave!CLI_AUTO1 = " "
    cli_llave!cli_auto2 = " "
    cli_llave!CLI_PRENDA = " "
    cli_llave!CLI_TELEF1 = " "
    cli_llave!CLI_TELEF2 = " "
    cli_llave!CLI_OTRO_CONTR = 0
    cli_llave!CLI_LETRA = 0
    cli_llave!CLI_GRUPO = 0
    cli_llave!CLI_SUBGRUPO = 0
    cli_llave!CLI_nucleo = 0
    cli_llave!CLI_estado = " "
    cli_llave!CLI_programado = " "
    cli_llave!CLI_CUENTA_CONTAB = "121500"
    cli_llave!CLI_CUENTA_CONTAB2 = " "
    cli_llave!CLI_DET_TOT = " "
    cli_llave!cli_limcre = 0
    cli_llave!cli_DIAS_FAC = 0
    cli_llave!cli_CUENTA_CONTAB22 = " "
    cli_llave!cli_dias_cred = 0
    cli_llave!CLI_estado = "A"
cli_llave.Update
On Error GoTo 0
sigue_otro:
Return

Exit Sub
GEN_CLI:
 MsgBox "Se produjo un conflicto .. Verifique si la lista de clientes se agregaron correctamennte, de lo contrario procesar nuevamente o manualmente  ", vbInformation, Pub_Titulo
 cli_llave.CancelUpdate
 GoTo sigue_otro

End Sub

Private Sub Form_Load()
LKCHEK = False
openvio(0).Value = True

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
PUB_CODCIA = LK_CODCIA
pu_codcia = PUB_CODCIA
LEER_CLI_LLAVE

If cli_mayor.EOF Then
    NUMCAD = "1"
    
Else
    cli_mayor.MoveLast
    NUMCAD = cli_mayor!cli_codclie
End If

VINT2 = 0
NUMCAD = Trim(NUMCAD)
VINT1 = Len(NUMCAD)
If NUMCAD = "1" Or NUMCAD = "2" Or NUMCAD = "0" Then
  VINT2 = 1
  VINT1 = 2
End If
If VINT1 > 1 Then
    VSTR4 = val(Mid(NUMCAD, 1, VINT1 - 2)) + 1
End If
For i = 1 To VINT1 - 2
   VSTR1 = Mid(VSTR4, i, 1)
   VINT2 = VINT2 + val(VSTR1)
Next i
VINT3 = VINT2 * 9

VSTR3 = Right(CStr(VINT3), 2)
If Len(VSTR3) = 1 Then
  VSTR3 = "0" & VSTR3
End If
FIJO = VSTR4
STRpub_cadena = FIJO & VSTR3
INTpub_cadena = val(STRpub_cadena)
GEN:
GENERA_CODI = INTpub_cadena

End Function

Private Sub openvio_Click(Index As Integer)
If Index = 0 Then
  TF_TIPMOV = 26
  lbltransfer.Caption = "Transfer. Ingreso de Mercaderia"
  lblarchi.Caption = Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\STANDAR\TMERCA.MDB"
ElseIf Index = 1 Then
  TF_TIPMOV = 27
  lbltransfer.Caption = "Transfer. Ventas de Mercaderia"
  lblarchi.Caption = Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\STANDAR\TRANSFER.MDB"
ElseIf Index = 2 Then
  TF_TIPMOV = 28
  lbltransfer.Caption = "Transfer. Ajustes de Mercaderia"
  lblarchi.Caption = Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\STANDAR\TAJUSTES.MDB"
End If

End Sub

