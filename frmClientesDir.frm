VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmClientesDir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Clientes"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12435
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   12435
   Begin TabDlg.SSTab SSTCliente 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Cliente"
      TabPicture(0)   =   "frmClientesDir.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(1)=   "txtUrb"
      Tab(0).Control(2)=   "DatZonaR"
      Tab(0).Control(3)=   "txtDir"
      Tab(0).Control(4)=   "txtFono"
      Tab(0).Control(5)=   "txtObs"
      Tab(0).Control(6)=   "txtEmail"
      Tab(0).Control(7)=   "dtpFN"
      Tab(0).Control(8)=   "txtDni"
      Tab(0).Control(9)=   "txtRuc"
      Tab(0).Control(10)=   "txtRS"
      Tab(0).Control(11)=   "lblurb"
      Tab(0).Control(12)=   "Label15"
      Tab(0).Control(13)=   "Label12"
      Tab(0).Control(14)=   "Label14"
      Tab(0).Control(15)=   "Label13"
      Tab(0).Control(16)=   "lblLC"
      Tab(0).Control(17)=   "Label11"
      Tab(0).Control(18)=   "Label10"
      Tab(0).Control(19)=   "Label9"
      Tab(0).Control(20)=   "lblCodigo"
      Tab(0).Control(21)=   "Label7"
      Tab(0).Control(22)=   "Label4"
      Tab(0).Control(23)=   "Label3"
      Tab(0).Control(24)=   "Label2"
      Tab(0).Control(25)=   "Label1"
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Información de Despacho"
      TabPicture(1)   =   "frmClientesDir.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   -71640
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3255
         Visible         =   0   'False
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtUrb 
         Height          =   285
         Left            =   -71640
         TabIndex        =   6
         Tag             =   "X"
         Top             =   2940
         Width           =   6375
      End
      Begin MSDataListLib.DataCombo DatZonaR 
         Height          =   315
         Left            =   -71640
         TabIndex        =   7
         Top             =   3300
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtDir 
         Height          =   285
         Left            =   -71640
         MaxLength       =   200
         TabIndex        =   5
         Tag             =   "X"
         Top             =   2580
         Width           =   6375
      End
      Begin VB.TextBox txtFono 
         Height          =   285
         Left            =   -71640
         TabIndex        =   4
         Tag             =   "X"
         Top             =   2220
         Width           =   2535
      End
      Begin VB.TextBox txtObs 
         Height          =   1125
         Left            =   -71640
         MultiLine       =   -1  'True
         TabIndex        =   10
         Tag             =   "X"
         Top             =   4380
         Width           =   6375
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   -71640
         TabIndex        =   9
         Tag             =   "X"
         Top             =   4020
         Width           =   6375
      End
      Begin MSComCtl2.DTPicker dtpFN 
         Height          =   285
         Left            =   -71640
         TabIndex        =   8
         Tag             =   "X"
         Top             =   3660
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   162463745
         CurrentDate     =   41537
      End
      Begin VB.TextBox txtDni 
         Height          =   285
         Left            =   -66960
         MaxLength       =   8
         TabIndex        =   3
         Tag             =   "X"
         Top             =   1860
         Width           =   1695
      End
      Begin VB.TextBox txtRuc 
         Height          =   285
         Left            =   -71640
         MaxLength       =   11
         TabIndex        =   2
         Tag             =   "X"
         Top             =   1860
         Width           =   2535
      End
      Begin VB.TextBox txtRS 
         Height          =   285
         Left            =   -71640
         MaxLength       =   80
         TabIndex        =   1
         Tag             =   "X"
         Top             =   1500
         Width           =   6375
      End
      Begin VB.Frame Frame2 
         Height          =   3495
         Left            =   120
         TabIndex        =   33
         Top             =   2820
         Width           =   11895
         Begin VB.TextBox txtRefAdd 
            Height          =   285
            Left            =   1440
            MaxLength       =   150
            TabIndex        =   13
            Tag             =   "X"
            Top             =   600
            Width           =   8295
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1935
            Left            =   1440
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1250
            Visible         =   0   'False
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.TextBox txtUrbanizacion 
            Height          =   285
            Left            =   1440
            TabIndex        =   14
            Top             =   960
            Width           =   4335
         End
         Begin VB.CommandButton cmdDirEdit 
            Height          =   345
            Left            =   11280
            Picture         =   "frmClientesDir.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdDirDel 
            Height          =   345
            Left            =   10800
            Picture         =   "frmClientesDir.frx":03C2
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdDirAdd 
            Height          =   345
            Left            =   10320
            Picture         =   "frmClientesDir.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   240
            Width           =   495
         End
         Begin MSDataListLib.DataCombo DatZona 
            Height          =   315
            Left            =   6840
            TabIndex        =   15
            Tag             =   "X"
            Top             =   945
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.TextBox txtDirAdd 
            Height          =   285
            Left            =   1440
            MaxLength       =   150
            TabIndex        =   12
            Tag             =   "X"
            Top             =   255
            Width           =   8295
         End
         Begin MSComctlLib.ListView lvDirecciones 
            Height          =   2055
            Left            =   120
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1320
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   3625
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia:"
            Height          =   195
            Left            =   360
            TabIndex        =   49
            Top             =   645
            Width           =   990
         End
         Begin VB.Label lblUrb2 
            Caption         =   "-1"
            Height          =   255
            Left            =   9960
            TabIndex        =   48
            Top             =   975
            Width           =   615
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Urbanización:"
            Height          =   195
            Left            =   180
            TabIndex        =   47
            Top             =   1005
            Width           =   1170
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Zona:"
            Height          =   195
            Left            =   6240
            TabIndex        =   36
            Top             =   1005
            Width           =   510
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   480
            TabIndex        =   35
            Top             =   300
            Width           =   870
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2415
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   11895
         Begin VB.CommandButton cmdFonoEdit 
            Height          =   345
            Left            =   11280
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmdFonoDel 
            Height          =   345
            Left            =   10800
            Picture         =   "frmClientesDir.frx":0AD6
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdFonoAdd 
            Height          =   345
            Left            =   10320
            Picture         =   "frmClientesDir.frx":0E60
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   240
            Width           =   495
         End
         Begin MSComctlLib.ListView lvTelefonos 
            Height          =   1695
            Left            =   120
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   600
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   2990
            View            =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.TextBox txtFonoAdd 
            Height          =   285
            Left            =   1440
            TabIndex        =   11
            Tag             =   "X"
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   360
            TabIndex        =   34
            Top             =   240
            Width           =   810
         End
      End
      Begin VB.Label lblurb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label16"
         Height          =   195
         Left            =   -65040
         TabIndex        =   46
         Top             =   2940
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Urbanización:"
         Height          =   195
         Left            =   -72960
         TabIndex        =   45
         Top             =   2940
         Width           =   1170
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zona:"
         Height          =   195
         Left            =   -72240
         TabIndex        =   44
         Top             =   3300
         Width           =   510
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   -72600
         TabIndex        =   43
         Top             =   2625
         Width           =   870
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
         Height          =   195
         Left            =   -72540
         TabIndex        =   42
         Top             =   2265
         Width           =   810
      End
      Begin VB.Label lblLC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -67800
         TabIndex        =   41
         Tag             =   "X"
         Top             =   3660
         Width           =   1740
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limite Credito:"
         Height          =   195
         Left            =   -69195
         TabIndex        =   40
         Top             =   3660
         Width           =   1275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia:"
         Height          =   195
         Left            =   -72735
         TabIndex        =   39
         Top             =   4380
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         Height          =   195
         Left            =   -72285
         TabIndex        =   38
         Top             =   4065
         Width           =   540
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -71640
         TabIndex        =   37
         Tag             =   "X"
         Top             =   1140
         Width           =   1740
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Nac.:"
         Height          =   195
         Left            =   -73020
         TabIndex        =   31
         Top             =   3705
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dni:"
         Height          =   195
         Left            =   -67440
         TabIndex        =   30
         Top             =   1905
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruc:"
         Height          =   195
         Left            =   -72135
         TabIndex        =   29
         Top             =   1905
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre o Razón Social:"
         Height          =   195
         Left            =   -73815
         TabIndex        =   28
         Top             =   1545
         Width           =   2070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Left            =   -72420
         TabIndex        =   27
         Top             =   1185
         Width           =   675
      End
   End
   Begin MSComctlLib.ImageList iCliente 
      Left            =   10560
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientesDir.frx":11EA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientesDir.frx":1784
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientesDir.frx":1D1E
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientesDir.frx":22B8
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientesDir.frx":2852
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbCliente 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   635
      ButtonWidth     =   2011
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "iCliente"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Guardar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmClientesDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean
Private vIfiltro As Integer
Public vTIPO As String
Dim loc_key  As Integer
Private vBuscar As Boolean 'variable para la busqueda de clientes
Dim loc_key2  As Integer
Private vBuscar2 As Boolean 'variable para la busqueda de clientes

Public gIDcliente As Double
Public gCliente As String
Public gGraba As Boolean

Sub Mandar_Datos()
CargarZonas
Me.lvDirecciones.ListItems.Clear
Me.lvTelefonos.ListItems.Clear



    
        Me.lblCodigo.Caption = gIDcliente
        Me.txtRS.Text = gCliente
  
        'Me.lblCodigo.Caption = .SelectedItem.Tag
        'Me.txtRS.Text = .SelectedItem.Text
        
   

    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_CLIENTE_FILL"
    
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , Me.lblCodigo.Caption)
    
    Dim ORSdc  As ADODB.Recordset

    Dim oRStmp As ADODB.Recordset
    
    Set ORSdc = oCmdEjec.Execute
        
    If Not ORSdc.EOF Then
        Me.txtFono.Text = ORSdc!FONO
        Me.txtDir.Text = ORSdc!dir
        Me.dtpFN.Value = ORSdc!FN
        Me.txtEmail.Text = ORSdc!MAIL
        Me.txtObs.Text = ORSdc!OBS
        Me.DatZonaR.BoundText = ORSdc!ZON
        Me.txtUrb.Text = ORSdc!urb
        Me.lblUrb.Caption = ORSdc!idurb
        Me.txtRuc.Text = ORSdc!RUC
    Me.txtDni.Text = ORSdc!DNI
    End If
    
    Set oRStmp = ORSdc.NextRecordset
    
    Dim oITEM As Object

    Do While Not oRStmp.EOF
        Set oITEM = Me.lvTelefonos.ListItems.Add(, , oRStmp!FONO)
        oRStmp.MoveNext
    Loop
    
    Set oRStmp = ORSdc.NextRecordset
      
    Do While Not oRStmp.EOF
        Set oITEM = Me.lvDirecciones.ListItems.Add(, , oRStmp!dir)
        oITEM.Tag = oRStmp!IDZ
        oITEM.SubItems(1) = oRStmp!ref
        oITEM.SubItems(2) = oRStmp!ZON
        oITEM.SubItems(3) = oRStmp!urb
        oITEM.SubItems(4) = oRStmp!ideu
        oITEM.SubItems(5) = oRStmp!IDEDIR
        oRStmp.MoveNext
    Loop
    
    Estado_Botones AntesDeActualizar


End Sub

Private Sub CargarZonas()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_CLIENTE_ZONAREPARTO_LIST"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim ORSz As ADODB.Recordset

    Set ORSz = oCmdEjec.Execute(, LK_CODCIA)
    Set Me.DatZona.RowSource = ORSz
    Me.DatZona.ListField = ORSz.Fields(1).Name
    Me.DatZona.BoundColumn = ORSz.Fields(0).Name
    Me.DatZona.BoundText = -1
    
    Set Me.DatZonaR.RowSource = ORSz
    Me.DatZonaR.ListField = ORSz.Fields(1).Name
    Me.DatZonaR.BoundColumn = ORSz.Fields(0).Name
    Me.DatZonaR.BoundText = -1
End Sub

Private Sub cmdDirAdd_Click()

If Len(Trim(Me.txtDirAdd.Text)) = 0 Then
    MsgBox "Debe ingresar la Dirección.", vbCritical, Pub_Titulo
    Me.txtDirAdd.SetFocus
ElseIf Me.DatZona.BoundText = -1 Then
    MsgBox "Debe elegir la Zona para la Dirección.", vbInformation, Pub_Titulo
    Me.DatZona.SetFocus
Else

    Dim iOBJ As Object

    If Me.lvDirecciones.ListItems.count = 0 Then
        Set iOBJ = Me.lvDirecciones.ListItems.Add(, , Me.txtDirAdd.Text)
        iOBJ.Tag = Me.DatZona.BoundText
        iOBJ.SubItems(1) = Me.txtRefAdd.Text
        iOBJ.SubItems(2) = Me.DatZona.Text
        iOBJ.SubItems(3) = Me.txtUrbanizacion.Text
        iOBJ.SubItems(4) = Me.lblUrb2.Caption
        Me.txtDirAdd.Text = ""
        Me.txtUrbanizacion.Text = ""
        Me.txtRefAdd.Text = ""
        Me.lblUrb2.Caption = "-1"
        Me.txtDirAdd.SetFocus
        Me.DatZona.BoundText = -1
            
    Else

        Dim vEd As Boolean

        vEd = False

        For Each iOBJ In Me.lvDirecciones.ListItems

            If iOBJ.Tag = Me.DatZona.BoundText And iOBJ.Text = Me.txtDirAdd.Text Then
                vEd = True

                Exit For

            End If

        Next

        If vEd Then
            MsgBox "Dirección y Zona repetida.", vbInformation, Pub_Titulo
            Me.txtDirAdd.SetFocus
            Me.txtDirAdd.SelStart = 0
            Me.txtDirAdd.SelLength = Len(Me.txtDirAdd.Text)
        Else
            Set iOBJ = Me.lvDirecciones.ListItems.Add(, , Me.txtDirAdd.Text)
            iOBJ.Tag = Me.DatZona.BoundText
            iOBJ.SubItems(1) = Me.txtRefAdd.Text
            iOBJ.SubItems(2) = Me.DatZona.Text
            iOBJ.SubItems(3) = Me.txtUrbanizacion.Text
            iOBJ.SubItems(4) = Me.lblUrb2.Caption
            Me.txtDirAdd.Text = ""
            Me.txtDirAdd.SetFocus
            Me.txtUrbanizacion.Text = ""
            Me.txtRefAdd.Text = ""
            Me.lblUrb2.Caption = "-1"
            Me.DatZona.BoundText = -1
        End If
    End If
End If

End Sub

Private Sub cmdDirDel_Click()
If Me.lvDirecciones.ListItems.count = 0 Then Exit Sub
Me.lvDirecciones.ListItems.Remove Me.lvDirecciones.SelectedItem.index
End Sub

Private Sub cmdDirEdit_Click()

frmClientesDirEdit.txtDirAdd.Text = Me.lvDirecciones.SelectedItem.Text
frmClientesDirEdit.DatZona.BoundText = Me.lvDirecciones.SelectedItem.Tag
frmClientesDirEdit.txtRefAdd.Text = Me.lvDirecciones.SelectedItem.SubItems(1)
frmClientesDirEdit.txtUrbanizacion.Text = Me.lvDirecciones.SelectedItem.SubItems(3)
frmClientesDirEdit.lblUrb.Caption = Me.lvDirecciones.SelectedItem.SubItems(4)
frmClientesDirEdit.lblIdDir.Caption = Me.lvDirecciones.SelectedItem.SubItems(5)
frmClientesDirEdit.Show vbModal

If frmClientesDirEdit.gAcepta Then
    Me.lvDirecciones.SelectedItem.Tag = frmClientesDirEdit.gIDZONA
    Me.lvDirecciones.SelectedItem.Text = frmClientesDirEdit.gDIRECCION
    Me.lvDirecciones.SelectedItem.SubItems(1) = frmClientesDirEdit.gREFERENCIA
    Me.lvDirecciones.SelectedItem.SubItems(2) = frmClientesDirEdit.gZONA
    Me.lvDirecciones.SelectedItem.SubItems(3) = frmClientesDirEdit.gURB
    Me.lvDirecciones.SelectedItem.SubItems(4) = frmClientesDirEdit.gIDURB
    Me.lvDirecciones.SelectedItem.SubItems(5) = frmClientesDirEdit.gIDDIR
End If

End Sub

Private Sub cmdFonoAdd_Click()

    If Len(Trim(Me.txtFonoAdd.Text)) = 0 Then Exit Sub

    Dim iObject As Object

    If Me.lvTelefonos.ListItems.count = 0 Then
        Set iObject = Me.lvTelefonos.ListItems.Add(, , Me.txtFonoAdd.Text)
        Me.txtFonoAdd.Text = ""
        Me.txtFonoAdd.SetFocus
    Else

        Dim vFonoE As Boolean

        vFonoE = False

        For Each iObject In Me.lvTelefonos.ListItems

            If Me.txtFonoAdd.Text = iObject Then
                vFonoE = True

                Exit For

            End If

        Next

        If Not vFonoE Then
            Set iObject = Me.lvTelefonos.ListItems.Add(, , Me.txtFonoAdd.Text)
            Me.txtFonoAdd.Text = ""
            Me.txtFonoAdd.SetFocus
        Else
            MsgBox "Teléfono repetido", vbCritical, Pub_Titulo
            Me.txtFonoAdd.SetFocus
            Me.txtFonoAdd.SelStart = 0
            Me.txtFonoAdd.SelLength = Len(Me.txtFonoAdd.Text)
        End If
    End If

End Sub

Private Sub cmdFonoDel_Click()
If Me.lvTelefonos.ListItems.count = 0 Then Exit Sub
Me.lvTelefonos.ListItems.Remove Me.lvTelefonos.SelectedItem.index
End Sub


Private Sub DatZonaR_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 Me.dtpFN.SetFocus
 End If
End Sub

Private Sub dtpFN_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.txtEmail.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
    ConfigurarLV
    
    Mandar_Datos

    
    'Estado_Botones InicializarFormulario
    ActivarControles Me


        Me.Caption = "Maestro de Clientes"

gGraba = False
    If vTIPO = "P" Then Me.SSTCliente.TabCaption(0) = "Proveedor"
    vBuscar = True
    vIfiltro = 0


End Sub

Private Sub ConfigurarLV()

    With Me.lvTelefonos
        .Gridlines = True
        .LabelEdit = lvwManual
        '.View = lvwReport
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Número", 3500
    End With

    With Me.lvDirecciones
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Dirección", 5500
        .ColumnHeaders.Add , , "Referencia", 3000
        .ColumnHeaders.Add , , "Zona", 2500
        .ColumnHeaders.Add , , "Urb", 1500
        .ColumnHeaders.Add , , "idUrb", 500
        .ColumnHeaders.Add , , "IDDIR", 500
    End With

   
    
    With Me.ListView1
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
  
        .ColumnHeaders.Add , , "Cliente", 5000
  
        .MultiSelect = False
    End With
    
 With Me.ListView2
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
  
        .ColumnHeaders.Add , , "Cliente", 5000
  
        .MultiSelect = False
    End With
End Sub
Private Sub Estado_Botones(val As Valores)
Select Case val
    

    Case AntesDeActualizar
        Me.tbCliente.Buttons(1).Enabled = True
        Me.tbCliente.Buttons(2).Enabled = True
        
        Me.SSTCliente.tab = 0
End Select
End Sub

Private Sub lvListado_DblClick()
Mandar_Datos
End Sub



Private Sub lvDirecciones_DblClick()
cmdDirEdit_Click
End Sub

Private Sub SSTCliente_Click(PreviousTab As Integer)

    If PreviousTab = 0 Then
        If Me.txtFonoAdd.Enabled Then
            Me.txtFonoAdd.SetFocus
            Me.txtFonoAdd.SelStart = 0
            Me.txtFonoAdd.SelLength = Len(Me.txtFonoAdd.Text)
        End If

    ElseIf PreviousTab = 1 Then

        If Me.txtRS.Enabled Then
           ' Me.txtRS.SetFocus
            Me.txtRS.SelStart = 0
            Me.txtRS.SelLength = Len(Me.txtRS.Text)
        End If
    End If

End Sub

Private Sub tbCliente_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.index

    Case 1 'Guardar
        LimpiaParametros oCmdEjec

        Dim vCONTINUA As Boolean

        Dim sMSN      As String
                
        If Len(Trim(Me.txtRS.Text)) = 0 Then
            MsgBox "Debe ingresar el Código", vbCritical, Pub_Titulo
            Me.txtRS.SetFocus
        ElseIf Me.DatZona.BoundText = "" Then
            MsgBox "Debe elegir la zona", vbCritical, Pub_Titulo
            Me.DatZona.SetFocus
        Else

            sMSN = ""
            vCONTINUA = False
            LimpiaParametros oCmdEjec

            Dim orsM As ADODB.Recordset
         
            oCmdEjec.CommandText = "SP_CLIENTE_VALIDA_M"
                
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RUC", adVarChar, adParamInput, 11, Trim(Me.txtRuc.Text))
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DNI", adChar, adParamInput, 8, Trim(Me.txtDni.Text))
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FONO", adVarChar, adParamInput, 20, Me.txtFono.Text)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RS", adVarChar, adParamInput, 80, Me.txtRS.Text)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adChar, adParamInput, 1, vTIPO)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , Me.lblCodigo.Caption)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CONTINUA", adBoolean, adParamOutput, 20, vCONTINUA)

            Set orsM = oCmdEjec.Execute
                
            vCONTINUA = oCmdEjec.Parameters("@CONTINUA").Value
                    
            sMSN = orsM!Mensaje
                    
            If Len(Trim(sMSN)) <> 0 Then
                If Not vCONTINUA Then
                    MsgBox sMSN, vbCritical, Pub_Titulo

                    Exit Sub

                Else

                    If MsgBox(sMSN + vbCrLf + "¿Desea continuar con la operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
                End If
                    
            End If
                    
            oCmdEjec.CommandText = "SP_CLIENTE_MODIFICAR"

            LimpiaParametros oCmdEjec

            On Error GoTo grabar

            Pub_ConnAdo.BeginTrans

            Dim vIDz As Double
               
            vIDz = 0

            oCmdEjec.Prepared = True
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CLI_NOMBRE", adVarChar, adParamInput, 80, Trim(Me.txtRS.Text))
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adChar, adParamInput, 1, "C")
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CLI_TELEF1", adVarChar, adParamInput, 20, Me.txtFono.Text)
                
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CLI_CASA_DIREC", adVarChar, adParamInput, 150, Me.txtDir.Text)
                
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RUC", adChar, adParamInput, 11, Me.txtRuc.Text)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DNI", adChar, adParamInput, 8, Me.txtDni.Text)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHANAC", adDBTimeStamp, adParamInput, , Me.dtpFN.Value)
                
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@EMAIL", adVarChar, adParamInput, 60, Me.txtEmail.Text)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@OBS", adVarChar, adParamInput, 1000, Me.txtObs.Text)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adInteger, adParamInput, , Me.DatZonaR.BoundText)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDURB", adBigInt, adParamInput, , Me.lblUrb.Caption)

            If VNuevo Then
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamOutput, , vIDz)
            Else
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , Me.lblCodigo.Caption)
            End If

            oCmdEjec.Execute
'oCmdEjec.Execute , Array(CodCia, Trim(Me.txtCodigo.Text), Trim(Me.txtDenominacion.Text), Trim(Me.txtZona.Text))

            vIDz = oCmdEjec.Parameters("@IDCLIENTE").Value
                
            If Not VNuevo Then
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SP_CLIENTE_TELEFONOS_DEL"
                   
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , Me.lblCodigo.Caption)
                oCmdEjec.Execute
                    
                oCmdEjec.CommandText = "SP_CLIENTE_DIRECCIONES_DEL"
                   
                oCmdEjec.Execute
            End If

'TELEFONOS
            Dim iOBJ As Object

            For Each iOBJ In Me.lvTelefonos.ListItems

                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SP_CLIENTE_TELEFONOS_REGISTRAR"
                oCmdEjec.CommandType = adCmdStoredProc
                    
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , vIDz)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FONO", adVarChar, adParamInput, 20, iOBJ.Text)
                oCmdEjec.Execute
            Next

'DIRECCIONES
            For Each iOBJ In Me.lvDirecciones.ListItems

                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SP_CLIENTE_DIRECCIONES_REGISTRAR"
                oCmdEjec.CommandType = adCmdStoredProc
                    
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , vIDz)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adInteger, adParamInput, , iOBJ.Tag)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIRECCION", adVarChar, adParamInput, 150, iOBJ.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@REFERENCIA", adVarChar, adParamInput, 100, iOBJ.SubItems(1))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDURB", adInteger, adParamInput, , iOBJ.SubItems(4))
                oCmdEjec.Execute
            Next
               
            DesactivarControles Me

            Estado_Botones grabar
            
            

            

'set itemg=me.lvMesas.ListItems.Add(,,
            Pub_ConnAdo.CommitTrans
            gGraba = True
            Unload Me

            Exit Sub

grabar:
            Pub_ConnAdo.RollbackTrans
            MsgBox Err.Description, vbInformation, NombreProyecto

        End If
      
    Case 3 'Cancelar
        gGraba = False
        Unload Me
      
End Select

End Sub

Private Sub txtDir_KeyPress(KeyAscii As Integer)
 KeyAscii = Mayusculas(KeyAscii)
 If KeyAscii = vbKeyReturn Then
   Me.txtUrb.SetFocus
 End If
End Sub

Private Sub txtDni_KeyPress(KeyAscii As Integer)
 KeyAscii = Mayusculas(KeyAscii)
 If KeyAscii = vbKeyReturn Then
    Me.txtFono.SetFocus
    Me.txtFono.SelStart = 0
    Me.txtFono.SelLength = Len(Me.txtFono.Text)
 End If
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then Me.txtObs.SetFocus
End Sub

Private Sub txtFono_KeyPress(KeyAscii As Integer)
 KeyAscii = Mayusculas(KeyAscii)
 If KeyAscii = vbKeyReturn Then
    Me.txtDir.SetFocus
    Me.txtDir.SelStart = 0
    Me.txtDir.SelLength = Len(Me.txtDir.Text)
 End If
End Sub

Private Sub txtRS_KeyPress(KeyAscii As Integer)
 KeyAscii = Mayusculas(KeyAscii)
 If KeyAscii = vbKeyReturn Then
    Me.txtRuc.SetFocus
    Me.txtRuc.SelStart = 0
    Me.txtRuc.SelLength = Len(Me.txtRuc.Text)
 End If
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
 KeyAscii = Mayusculas(KeyAscii)
 If KeyAscii = vbKeyReturn Then
    Me.txtDni.SetFocus
    Me.txtDni.SelStart = 0
    Me.txtDni.SelLength = Len(Me.txtDni.Text)
 End If
End Sub


Private Sub txtUrb_Change()
vBuscar = True
Me.lblUrb.Caption = -1
End Sub

Private Sub txtUrb_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key = loc_key + 1

        If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 38 Then
        loc_key = loc_key - 1

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 34 Then
        loc_key = loc_key + 17

        If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 33 Then
        loc_key = loc_key - 17

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 27 Then
        Me.ListView1.Visible = False
        Me.txtUrb.Text = ""
Me.lblUrb.Caption = "-1"
    End If

    GoTo fin
posicion:
    ListView1.ListItems.Item(loc_key).Selected = True
    ListView1.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    txtUrb.SelStart = Len(Me.txtUrb.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtUrb_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscar Then
            Me.ListView1.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpListarCliUrbanizaciones"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 80, Me.txtUrb.Text)
                
            Dim ORSurb As ADODB.Recordset
                
            Set ORSurb = oCmdEjec.Execute

            Dim Item As Object
        
            If Not ORSurb.EOF Then

                Do While Not ORSurb.EOF
                    Set Item = Me.ListView1.ListItems.Add(, , ORSurb!nom)
                    Item.Tag = ORSurb!IDE
                    ORSurb.MoveNext
                Loop

                Me.ListView1.Visible = True
                Me.ListView1.ListItems(1).Selected = True
                loc_key = 1
                Me.ListView1.ListItems(1).EnsureVisible
                vBuscar = False
            Else

                If MsgBox("Urbanización no encontrada." & vbCrLf & "¿Desea agregarla?", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
                frmClientesUrbAdd.txtUrb.Text = Me.txtUrb.Text
                frmClientesUrbAdd.Show vbModal

                If frmClientesUrbAdd.gAcepta Then
                    Me.txtUrb.Text = frmClientesUrbAdd.gNombre
                    Me.lblUrb.Caption = frmClientesUrbAdd.gIde
                    Me.ListView1.Visible = False
                    vBuscar = True
                    Me.DatZona.SetFocus
                End If
            End If
        
        Else
            Me.txtUrb.Text = Me.ListView1.ListItems(loc_key).Text
            Me.lblUrb.Caption = Me.ListView1.ListItems(loc_key).Tag
            
            Me.ListView1.Visible = False
            vBuscar = True
            Me.DatZona.SetFocus
            'Me.lvDetalle.SetFocus
        End If
    End If

End Sub

Private Sub txtUrbanizacion_Change()
vBuscar2 = True
Me.lblUrb2.Caption = -1
End Sub

Private Sub txtUrbanizacion_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key2 = loc_key2 + 1

        If loc_key2 > ListView2.ListItems.count Then loc_key2 = ListView2.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 38 Then
        loc_key2 = loc_key2 - 1

        If loc_key2 < 1 Then loc_key2 = 1
        GoTo posicion
    End If

    If KeyCode = 34 Then
        loc_key2 = loc_key2 + 17

        If loc_key2 > ListView2.ListItems.count Then loc_key2 = ListView2.ListItems.count
        GoTo posicion
    End If

    If KeyCode2 = 33 Then
        loc_key2 = loc_key2 - 17

        If loc_key2 < 1 Then loc_key2 = 1
        GoTo posicion
    End If

    If KeyCode = 27 Then
        Me.ListView2.Visible = False
        Me.txtUrbanizacion.Text = ""
Me.lblUrb2.Caption = "-1"
    End If

    GoTo fin
posicion:
    ListView2.ListItems.Item(loc_key).Selected = True
    ListView2.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    txtUrbanizacion.SelStart = Len(Me.txtUrbanizacion.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtUrbanizacion_KeyPress(KeyAscii As Integer)
  KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscar2 Then
            Me.ListView2.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpListarCliUrbanizaciones"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 80, Me.txtUrbanizacion.Text)
                
            Dim ORSurb As ADODB.Recordset
                
            Set ORSurb = oCmdEjec.Execute

            Dim Item As Object
        
            If Not ORSurb.EOF Then

                Do While Not ORSurb.EOF
                    Set Item = Me.ListView2.ListItems.Add(, , ORSurb!nom)
                    Item.Tag = ORSurb!IDE
                    ORSurb.MoveNext
                Loop

                Me.ListView2.Visible = True
                Me.ListView2.ListItems(1).Selected = True
                loc_key2 = 1
                Me.ListView2.ListItems(1).EnsureVisible
                vBuscar2 = False
            Else

                If MsgBox("Urbanización no encontrada." & vbCrLf & "¿Desea agregarla?", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
                frmClientesUrbAdd.txtUrb.Text = Me.txtUrbanizacion.Text
                frmClientesUrbAdd.Show vbModal

                If frmClientesUrbAdd.gAcepta Then
                    Me.txtUrbanizacion.Text = frmClientesUrbAdd.gNombre
                    Me.lblUrb2.Caption = frmClientesUrbAdd.gIde
                    Me.ListView2.Visible = False
                    vBuscar = True
                    Me.DatZona.SetFocus
                End If
            End If
        
        Else
            Me.txtUrbanizacion.Text = Me.ListView2.ListItems(loc_key2).Text
            Me.lblUrb2.Caption = Me.ListView2.ListItems(loc_key2).Tag
            
            Me.ListView2.Visible = False
            vBuscar2 = True
            Me.DatZona.SetFocus
            'Me.lvDetalle.SetFocus
        End If
    End If
End Sub
