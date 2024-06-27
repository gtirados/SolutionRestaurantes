VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProgramacionRequerida 
   Caption         =   "Programación Requerida"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13500
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   13500
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer crvVisor 
      Height          =   4695
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   12855
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   600
      Left            =   10560
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   600
      Left            =   9240
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtMenus 
      Height          =   285
      Left            =   6720
      TabIndex        =   5
      Top             =   278
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   270
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   156434433
      CurrentDate     =   41498
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   300
      Left            =   3840
      TabIndex        =   4
      Top             =   270
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   156434433
      CurrentDate     =   41498
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENUS:"
      Height          =   195
      Left            =   5760
      TabIndex        =   2
      Top             =   323
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HASTA:"
      Height          =   195
      Left            =   2880
      TabIndex        =   1
      Top             =   323
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESDE:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   323
      Width           =   675
   End
End
Attribute VB_Name = "frmProgramacionRequerida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBuscar_Click()

    If Len(Trim(Me.txtMenus.Text)) = 0 Then
        MsgBox "Debe ingresar la Cantidad de menus.", vbCritical, Pub_Titulo
        Me.txtMenus.SetFocus
        Exit Sub
    End If
    
    If val(Me.txtMenus.Text) <= 0 Then
        MsgBox "La Cantidad de Menus es incorrecta.", vbCritical, Pub_Titulo
      Me.txtMenus.SetFocus
        Me.txtMenus.SelStart = 0
        Me.txtMenus.SelLength = Len(Me.txtMenus.Text)
        Exit Sub
    End If
    
    If Not IsNumeric(Me.txtMenus.Text) Then
        MsgBox "La Cantidad es incorrecta.", vbCritical, Pub_Titulo
        Me.txtMenus.SetFocus
        Me.txtMenus.SelStart = 0
        Me.txtMenus.SelLength = Len(Me.txtMenus.Text)
        Exit Sub
    End If
    

    Dim objCrystal  As New CRAXDRT.Application

    Dim RutaReporte As String

    RutaReporte = "c:\Admin\Nordi\PROREQ.rpt"

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
            
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.CommandText = "SP_PROGRAMACION_REQUERIDA"
    'oCmdEjec.CommandText = "SpPrintComanda"

    Dim rsd As ADODB.Recordset
            
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@F1", adDBTimeStamp, adParamInput, , Me.dtpDesde.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@F2", adDBTimeStamp, adParamInput, , Me.dtpHasta.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MENUS", adBigInt, adParamInput, , Me.txtMenus.Text)

    Set rsd = oCmdEjec.Execute
            
    VReporte.Database.SetDataSource rsd, 3, 1
    'frmprint.CRViewer1.ReportSource = VReporte
    'frmprint.CRViewer1.ViewReport
    Me.crvVisor.ReportSource = VReporte
    Me.crvVisor.ViewReport
    
    'VReporte.PrintOut False, 1, , 1, 1
    Set objCrystal = Nothing
    Set VReporte = Nothing

End Sub



Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub dtpDesde_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.dtpHasta.SetFocus
End Sub

Private Sub dtpHasta_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        Me.txtMenus.SetFocus
        Me.txtMenus.SelStart = 0
        Me.txtMenus.SelLength = Len(Me.txtMenus.Text)
    End If

End Sub

Private Sub Form_Resize()
If (Me.ScaleWidth - 6300) <= 0 Then Exit Sub
If (Me.ScaleHeight - 1800) <= 0 Then Exit Sub
Me.crvVisor.Width = Me.ScaleWidth
Me.crvVisor.Height = Me.ScaleHeight - 900

End Sub

Private Sub txtMenus_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdBuscar_Click
End Sub
