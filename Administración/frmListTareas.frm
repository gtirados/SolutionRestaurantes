VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListTareas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de Tareas Pendientes"
   ClientHeight    =   3495
   ClientLeft      =   3630
   ClientTop       =   3015
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2010
      Top             =   2970
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListTareas.frx":0000
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   5640
      TabIndex        =   2
      Top             =   3120
      Width           =   915
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   360
      Left            =   4530
      TabIndex        =   1
      Top             =   3120
      Width           =   990
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1260
      Top             =   2085
   End
   Begin MSComctlLib.ListView lvwTareas 
      Height          =   3060
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   5398
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmListTareas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAceptar_Click()
  If Not lvwTareas.ListItems.Count = 0 Then
   frmTareas.txtCodigo = Trim(Left$(lvwTareas.SelectedItem.Key, _
                         Len(Trim(lvwTareas.SelectedItem.Key)) - 1))
   Unload Me
   If msg = "HELP" Then
    frmTareas.ShowData
   End If
  End If
End Sub
Private Sub Command2_Click()
 Unload Me
End Sub
Private Sub Form_Load()
Dim sql As String
Dim str As String
If msg = "HELP" Then
    cmdAceptar.Default = True
    sql = "SELECT * FROM TAREAS WHERE USER_TAREA='" & LK_CODUSU & "' ORDER BY FECHA"
    rsT.CursorLocation = adUseClient
    On Error GoTo ErrorLoad
    rsT.Open sql, cnT, adOpenForwardOnly, adLockOptimistic, adCmdText
Else
  rsT.CursorLocation = adUseClient
  str = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=BDATOS;Origen de datos=server"
  rsT.Open "Exec sp_task_user '" & LK_CODUSU & "'", str, adOpenForwardOnly, adLockOptimistic, adCmdText
  DoEvents
  cmdAceptar.Default = False
  Timer1.Enabled = True
End If
With lvwTareas
 .ColumnHeaders.Clear
 .ColumnHeaders.Add , "usu", "Usuario", 1400, lvwColumnLeft
 .ColumnHeaders.Add , "desc", "Descripción", 3000, lvwColumnLeft
 .ColumnHeaders.Add , "fecha", "Fecha", 2500, lvwColumnLeft
 .ColumnHeaders.Add , "hora", "Hora", 1200, lvwColumnLeft
 .SmallIcons = ImageList2
 .LabelEdit = lvwManual
 .View = lvwReport
 .FullRowSelect = True
End With
Requery
Exit Sub
ErrorLoad:
 MsgBox Err.Description
End Sub
Sub Requery()
Dim Clave As String
On Error Resume Next
rsT.Requery
With lvwTareas
 .ListItems.Clear
 rsT.MoveFirst
 While Not rsT.EOF
  Clave = rsT!IdTarea & "A"
  Set mItem = lvwTareas.ListItems.Add(Key:=Clave, text:=rsT!user_tarea, SmallIcon:="user")
  mItem.ListSubItems.Add = "" & rsT(2)
  mItem.ListSubItems.Add = "" & Format(rsT(3), "dddddd")
  mItem.ListSubItems.Add = "" & Right(CDate(rsT(4)), 12)
  rsT.MoveNext
 Wend
End With
End Sub
Private Sub Form_Resize()
If Not WindowState = vbMinimized Then
 lvwTareas.Height = Height - cmdAceptar.Height - 450
 lvwTareas.Width = Width - 250
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set rsT = Nothing
End Sub
Private Sub lvwTareas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
 rsT.MoveFirst
 rsT.Find "IdTarea='" & Trim(Left$(lvwTareas.SelectedItem.Key, Len(Trim(lvwTareas.SelectedItem.Key)) - 1)) & "'"
 lvwTareas.ToolTipText = rsT!obs
End If
End Sub
Private Sub Timer1_Timer()
Caption = IIf(Caption = "", "Listado de Tareas Pendientes", "")
DoEvents
End Sub
