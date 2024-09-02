VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDeliveryEnviar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Generar Comprobante"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   195
   ClientWidth     =   6840
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcopias 
      Height          =   285
      Left            =   3000
      TabIndex        =   39
      Text            =   "1"
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkprom 
      Caption         =   "Imprime Promocion"
      Height          =   255
      Left            =   5400
      TabIndex        =   37
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1215
      Left            =   1440
      TabIndex        =   29
      Top             =   2280
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2143
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
   Begin VB.TextBox txtDni 
      Height          =   315
      Left            =   1440
      TabIndex        =   34
      Top             =   2640
      Width           =   4935
   End
   Begin VB.CheckBox chkEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5640
      TabIndex        =   31
      Top             =   720
      Width           =   975
   End
   Begin MSDataListLib.DataCombo DatTiposDoctos 
      Height          =   315
      Left            =   1680
      TabIndex        =   30
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DatDireccion 
      Height          =   315
      Left            =   1680
      TabIndex        =   15
      Top             =   1200
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   600
      Left            =   3720
      TabIndex        =   13
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   600
      Left            =   1440
      TabIndex        =   12
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CheckBox chkConsumo 
      Caption         =   "Por Consumo"
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtAbono 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   2520
      TabIndex        =   9
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtNumero 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   4320
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo DatRepartidor 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Frame Frame1 
      Caption         =   "FACTURAR A"
      Height          =   1815
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   6615
      Begin VB.CommandButton cmdSunat 
         Height          =   315
         Left            =   6000
         TabIndex        =   32
         Top             =   240
         Width           =   510
      End
      Begin VB.TextBox txtDireccion 
         Height          =   315
         Left            =   1320
         TabIndex        =   21
         Top             =   1320
         Width           =   4935
      End
      Begin VB.TextBox txtRuc 
         Height          =   315
         Left            =   1320
         TabIndex        =   20
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   1320
         TabIndex        =   18
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label lblDNI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI:"
         Height          =   195
         Left            =   840
         TabIndex        =   35
         Top             =   960
         Width           =   405
      End
      Begin VB.Label lblcodclie 
         Caption         =   "Label13"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCIÓN:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   1110
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RUC:"
         Height          =   195
         Left            =   780
         TabIndex        =   19
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE:"
         Height          =   195
         Left            =   420
         TabIndex        =   17
         Top             =   240
         Width           =   810
      End
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ICBPER"
      Height          =   195
      Left            =   5280
      TabIndex        =   38
      Top             =   3720
      Width           =   660
   End
   Begin VB.Label lblicbper 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5280
      TabIndex        =   36
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCUENTO"
      Height          =   195
      Left            =   3840
      TabIndex        =   28
      Top             =   3720
      Width           =   1080
   End
   Begin VB.Label lblDscto 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3720
      TabIndex        =   27
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL PEDIDO:"
      Height          =   195
      Left            =   240
      TabIndex        =   26
      Top             =   3720
      Width           =   1380
   End
   Begin VB.Label lblTotal2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      TabIndex        =   25
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOVILIDAD:"
      Height          =   195
      Left            =   2160
      TabIndex        =   24
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblMovilidad 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2040
      TabIndex        =   23
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION:"
      Height          =   195
      Left            =   450
      TabIndex        =   14
      Top             =   1260
      Width           =   1110
   End
   Begin VB.Label lblVuelto 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4800
      TabIndex        =   10
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VUELTO:"
      Height          =   195
      Left            =   5145
      TabIndex        =   7
      Top             =   4320
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAGA CON:"
      Height          =   195
      Left            =   2745
      TabIndex        =   6
      Top             =   4320
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL:"
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   4320
      Width           =   630
   End
   Begin VB.Label lblserie 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DOCUMENTO:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   780
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REPARTIDOR:"
      Height          =   195
      Left            =   345
      TabIndex        =   0
      Top             =   300
      Width           =   1215
   End
End
Attribute VB_Name = "frmDeliveryEnviar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gTOTAL As Double
Public gMOVILIDAD As Double
Public gDESCUENTO As Double
Public gIDDIR As Integer
Public gICBPER As Double
Private vPUNTO As Boolean
Public vGraba As Boolean
Dim loc_key  As Integer
Private vBuscar As Boolean 'variable para la busqueda de clientes
Private ORStd As ADODB.Recordset 'VARIABLE PARA SABER SI EL TIPO DE DOCUMENTO ES EDITABLE
Private Sub CrearArchivoPlano(cTipoDocto As String, cSerie As String, cNumero As Double)

    Dim oRS As ADODB.Recordset

    LimpiaParametros oCmdEjec

    If cTipoDocto = "F" Then
           oCmdEjec.CommandText = "SP_VENTA_FACTURA_SFS"
    ElseIf cTipoDocto = "B" Then
           oCmdEjec.CommandText = "SP_VENTA_BOLETA_SFS"
    ElseIf LK_CODTRA = 1111 Then
    
    End If
    
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@serie", adVarChar, adParamInput, 3, IIf(LK_CODTRA = 1111, PUB_NUMSER_C, cSerie))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numero", adDouble, adParamInput, , IIf(LK_CODTRA = 1111, PUB_NUMFAC_C, cNumero))
    If LK_CODTRA = 1111 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adChar, adParamInput, 1, PUB_FBG)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TRANSACCION", adBigInt, adParamInput, , LK_CODTRA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    End If
    
    Set oRS = oCmdEjec.Execute
    
    Dim sCadena As String

    sCadena = ""
    
    Dim obj_FSO     As Object

    Dim ArchivoCab  As Object
    Dim ArchivoTri As Object
    Dim ArchivoDet  As Object
    Dim ArchivoLey As Object
    Dim ArchivoAca As Object
    
    Dim sARCHIVOcab As String
    Dim sARCHIVOdet As String
    Dim sARCHIVOtri As String
    Dim sARCHIVOley As String
    Dim sARCHIVOaca As String
    
    Dim sRUC        As String
    
    If LK_CODCIA = "01" Then
    sRUC = Leer_Ini(App.Path & "\config.ini", "RUC", "C:\")
    ElseIf LK_CODCIA = "02" Then
    sRUC = Leer_Ini(App.Path & "\config2.ini", "RUC", "C:\")
    ElseIf LK_CODCIA = "03" Then
    sRUC = Leer_Ini(App.Path & "\config3.ini", "RUC", "C:\")
    Else
    sRUC = Leer_Ini(App.Path & "\config4.ini", "RUC", "C:\")
    End If
     
    sARCHIVOcab = sRUC & "-" & oRS!Nombre + IIf(LK_CODTRA = 2412, ".not", IIf(LK_CODTRA = 1111, ".cba", ".cab"))
        
    If LK_CODTRA <> 1111 Then
        sARCHIVOdet = sRUC & "-" & oRS!Nombre + ".det"
        sARCHIVOtri = sRUC & "-" & oRS!Nombre + ".tri" 'IIf(LK_CODTRA = 2412, ".not", IIf(LK_CODTRA = 1111, ".tri", ".tri"))
        sARCHIVOley = sRUC & "-" & oRS!Nombre + ".ley" 'IIf(LK_CODTRA = 2412, ".not", IIf(LK_CODTRA = 1111, ".ley", ".ley"))
        sARCHIVOaca = sRUC & "-" & oRS!Nombre + ".aca"
     If cTipoDocto = "F" Then 'es factura
        sARCHIVOpag = sRUC & "-" & oRS!Nombre + ".pag"
        sARCHIVOdpa = sRUC & "-" & oRS!Nombre + ".dpa"
        sARCHIVOrtn = sRUC & "-" & oRS!Nombre + ".rtn"
        End If
    End If
    
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")

    'Creamos un archivo con el método CreateTextFile
    If LK_CODCIA = "01" Then
    Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
    Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
    Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOley, True)
    Set ArchivoAca = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOaca, True)
    If cTipoDocto = "F" Then
    Set ArchivoPAG = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOpag, True)
    End If
    ElseIf LK_CODCIA = "02" Then
    Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
    Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
    Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOley, True)
    Set ArchivoAca = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOaca, True)
    If cTipoDocto = "F" Then
    Set ArchivoPAG = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOpag, True)
    End If
    ElseIf LK_CODCIA = "03" Then
    Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
    Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
    Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOley, True)
    Set ArchivoAca = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOaca, True)
    If cTipoDocto = "F" Then
    Set ArchivoPAG = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOpag, True)
    End If
    ElseIf LK_CODCIA = "04" Then
    Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
    Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
    Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOley, True)
    Set ArchivoAca = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOaca, True)
    If cTipoDocto = "F" Then
    Set ArchivoPAG = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOpag, True)
    End If
    End If
    If LK_CODCIA = "01" Then
    Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
    ElseIf LK_CODCIA = "02" Then
    Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
    ElseIf LK_CODCIA = "03" Then
    Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
    Else
    Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
    End If
    
    'Set Archivo = obj_FSO.CreateTextFile("C:\" + sARCHIVO, True)
    
    If LK_CODTRA = 2412 Then

        Do While Not oRS.EOF
            sCadena = sCadena & oRS!fecemision & "|" & oRS!CODMOTIVO & "|" & oRS!DESCMOTIVO & "|" & oRS!TIPODOCAFECTADO & "|" & oRS!NUMDOCAFECTADO & "|" & oRS!TIPDOCUSUARIO & "|" & oRS!NUMDOCUSUARIO & "|" & oRS!CLI1 & "|" & oRS!TIPMONEDA & "|" & oRS!SUMOTROSCARGOS & "|" & oRS!MTOOPERGRAVADAS & "|" & oRS!MTOOPERINAFECTAS & "|" & oRS!MTOOPEREXONERADAS & "|" & oRS!MTOIGV & "|" & oRS!MTOISC & "|" & oRS!MTOOTROSTRIBUTOS & "|" & oRS!MTOIMPVENTA & "|"
            oRS.MoveNext
        Loop
    
    ElseIf LK_CODTRA = 1111 Then
         Do While Not oRS.EOF
            sCadena = sCadena & oRS!FEC_GENERACcION & "|" & oRS!FEC_COMUNICACION & "|" & oRS!TIPDOCBAJA & "|" & oRS!NUMDOCBAJA & "|" & oRS!DESMOTIVOBAJA & "|"
            oRS.MoveNext
        Loop
    Else

        Do While Not oRS.EOF
            sCadena = sCadena & oRS!TIPOPERACION & "|" & oRS!fecemision & "|" & oRS!hORA & "|" & oRS!FECHAVENC & "|" & oRS!codlocalemisor & "|" & oRS!TIPDOCUSUARIO & "|" & oRS!NUMDOCUSUARIO & "|" & oRS!rznsocialusuario & "|" & oRS!TIPMONEDA & "|" & oRS!MTOIGV & "|" & oRS!MTOOPERGRAVADAS & "|" & oRS!MTOIMPVENTA & "|" & oRS!SUMDSCTOGLOBAL & "|" & oRS!SUMOTROSCARGOS & "|" & oRS!TOTANTICIPOS & "|" & oRS!IMPTOTALVENTA & "|" & oRS!UBL & "|" & oRS!CUSTOMDOC & "|"
            oRS.MoveNext
        Loop

    End If
   
    'Escribimos lineas
    ArchivoCab.WriteLine sCadena
    
    'Cerramos el fichero
    ArchivoCab.Close
    Set ArchivoCab = Nothing
    
    If LK_CODTRA <> 1111 Then
         'DIRECCION
    oRS.MoveFirst
    sCadena = ""
    Do While Not oRS.EOF
        sCadena = sCadena & oRS!ACA1 & "|" & oRS!ACA2 & "|" & oRS!ACA3 & "|" & oRS!ACA4 & "|" & oRS!ACA5 & "|" & oRS!PAIS & "|" & oRS!UBIGEO & "|" & oRS!dir & "|" & oRS!PAIS1 & "|" & oRS!UBIGEO1 & "|" & oRS!dir1 & "|"
        oRS.MoveNext
    Loop
    
    'Escribimos LINEAS
    ArchivoAca.WriteLine sCadena
    
    'Cerramos el fichero
    ArchivoAca.Close
    Set ArchivoAca = Nothing
    Else
    End If
    
   
    Dim oRSdet As ADODB.Recordset

    Set oRSdet = oRS.NextRecordset
   
    sCadena = ""
    Dim c As Integer
    c = 1

    If LK_CODTRA = 2412 Then

        Do While Not oRSdet.EOF
         
            sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & oRSdet!DESITEM & "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!CODTIPTRIBUTOIGV & "|" & oRSdet!MTOIGVITEM1 & "|" & oRSdet!NOMTRIBITEM & "|" & oRSdet!CODTIPTRIBUTOITEM & "|" & oRSdet!TIPAFEIGV & "|" & FormatNumber(oRSdet!PORCIGV, 2) & "|" & oRSdet!CODISC & "|" & oRSdet!CODOTROITEM & "|" & oRSdet!GRATUITO & "|"
            If c < oRSdet.RecordCount Then
                sCadena = sCadena + vbCrLf
            End If
             c = c + 1
            oRSdet.MoveNext
            
        Loop

    ElseIf LK_CODTRA <> 1111 Then
    

        Do While Not oRSdet.EOF
       
           ' sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & oRSdet!DESITEM & "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTODSCTOITEM & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!TIPAFEIGV & "|" & oRSdet!MTOISCITEM & "|" & oRSdet!TIPSISISC & "|" & oRSdet!MTOPRECIOVENTAITEM & "|" & oRSdet!MTOVALORVENTAITEM & "|"
           sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & Trim(oRSdet!DESITEM) & _
           "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!CODTIPTRIBUTOIGV & "|" & oRSdet!MTOIGVITEM1 & "|" & oRSdet!BASEIMPIGV & "|" & _
           oRSdet!NOMTRIBITEM & "|" & oRSdet!CODTIPTRIBUTOITEM & "|" & oRSdet!TIPAFEIGV & "|" & FormatNumber(oRSdet!PORCIGV, 2) & "|" & oRSdet!CODISC & "|" & oRSdet!MONTOISC & _
           "|" & oRSdet!BASEIMPONIBLEISC & "|" & oRSdet!NOMBRETRIBITEM & "|" & oRSdet!CODTRIBITEM & "|" & oRSdet!CODSISISC & "|" & oRSdet!PORCISC & "|" & oRSdet!CODTRIBOTO & _
           "|" & oRSdet!MONTOTRIBOTO & "|" & oRSdet!BASEIMPONIBLEOTO & "|" & oRSdet!NOMBRETRIBOTO & "|" & oRSdet!TIPSISISC & "|" & oRSdet!PORCOTO & "|" & oRSdet!CODIGOICBPER & _
           "|" & oRSdet!IMPORTEICBPER & "|" & oRSdet!CANTIDADICBPER & "|" & oRSdet!TITULOICBPER & "|" & oRSdet!IDEICBPER & "|" & oRSdet!MONTOICBPER & "|" & _
           oRSdet!PRECIOVTAUNITARIO & "|" & oRSdet!VALORVTAXITEM & "|" & oRSdet!GRATUITO & "|"
            If c < oRSdet.RecordCount Then
                sCadena = sCadena + vbCrLf
            End If
             c = c + 1
            oRSdet.MoveNext
             
        Loop

    End If

    'Escribimos lineas
    If LK_CODTRA <> 1111 Then
    ArchivoDet.WriteLine sCadena
    
     'Cerramos el fichero
    ArchivoDet.Close
    Set ArchivoDet = Nothing
    
    Dim orsTri As ADODB.Recordset
    Set orsTri = oRS.NextRecordset
    
    sCadena = ""
    c = 1
    'ARCIVO .TRI
    Do While Not orsTri.EOF
    sCadena = sCadena & orsTri!Codigo & "|" & orsTri!Nombre & "|" & orsTri!cod & "|" & orsTri!BASEIMPONIBLE & "|" & orsTri!TRIBUTO & "|"
    If c < orsTri.RecordCount Then
        sCadena = sCadena & vbCrLf
    End If
    c = c + 1
        orsTri.MoveNext
    Loop
    
    
     ArchivoTri.WriteLine sCadena
    
     'Cerramos el fichero
    ArchivoTri.Close
    Set ArchivoTri = Nothing
    
    Dim orsLey As ADODB.Recordset
    Set orsLey = oRS.NextRecordset
    
    c = 1
    sCadena = ""
    Do While Not orsLey.EOF
        sCadena = sCadena & orsLey!cod & "|" & Trim(CONVER_LETRAS(Me.lbltotal2.Caption, "S")) & "|"
        If c < orsLey.RecordCount Then
            sCadena = sCadena & vbCrLf
        End If
        c = c + 1
        orsLey.MoveNext
    Loop
    
    ArchivoLey.WriteLine sCadena
    ArchivoLey.Close
    Set ArchivoLey = Nothing
    
    Dim xFormaPago As String
    If cTipoDocto = "F" Then
            'PAG
            Dim orsPAG As ADODB.Recordset
            Set orsPAG = oRS.NextRecordset
            
            c = 1
            sCadena = ""
            Do While Not orsPAG.EOF
                xFormaPago = orsPAG!formapago
                sCadena = sCadena & orsPAG!formapago & "|" & orsPAG!pendientepago & "|" & orsPAG!TIPMONEDA & "|"
                If c < orsPAG.RecordCount Then
                    sCadena = sCadena & vbCrLf
                End If
                c = c + 1
                orsPAG.MoveNext
            Loop
            
            ArchivoPAG.WriteLine sCadena
            ArchivoPAG.Close
            Set ArchivoPAG = Nothing
            
            'DPA
            Dim orsDPA As ADODB.Recordset
            Set orsDPA = oRS.NextRecordset
            If UCase(xFormaPago) = "CREDITO" Or UCase(xFormaPago) = "CRÉDITO" Then
                Set ArchivoDPA = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOdpa, True)
               
                
                c = 1
                sCadena = ""
                Do While Not orsDPA.EOF
                    sCadena = sCadena & orsDPA!cuotapago & "|" & orsDPA!fechavcto & "|" & orsDPA!TIPMONEDA & "|"
                    If c < orsDPA.RecordCount Then
                        sCadena = sCadena & vbCrLf
                    End If
                    c = c + 1
                    orsDPA.MoveNext
                Loop
                
                ArchivoDPA.WriteLine sCadena
                ArchivoDPA.Close
                Set ArchivoDPA = Nothing
            End If
             'RTN
'            Dim orsRTN As ADODB.Recordset
'            Set orsRTN = oRS.NextRecordset
'
'            c = 1
'            sCadena = ""
'            Do While Not orsRTN.EOF
'                sCadena = sCadena & orsRTN!impoperacion & "|" & orsRTN!porretencion & "|" & orsRTN!impretencion & "|"
'                If c < orsRTN.RecordCount Then
'                    sCadena = sCadena & vbCrLf
'                End If
'                c = c + 1
'                orsRTN.MoveNext
'            Loop
'
'            ArchivoRTN.WriteLine sCadena
'            ArchivoRTN.Close
'            Set ArchivoRTN = Nothing
        End If
    
    End If
    
   
    
    Set obj_FSO = Nothing
End Sub

Private Sub Imprimir(TipoDoc As String, Esconsumo As Boolean)
    'OBTENIENDO DATOS DEL CLIENTE
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_DELIVERY_DOCTOCLIENTE"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim ORSd As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , frmDeliveryApp.lblCliente.Caption)
    
    Set ORSd = oCmdEjec.Execute

    Dim Vdocto As String
    
    If Not ORSd.EOF Then
        Vdocto = Trim(ORSd!DOCTO)
    End If

    On Error GoTo printe

    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal  As New CRAXDRT.APPLICATION

    Dim vIgv        As Currency

    Dim vSubTotal   As Currency

    Dim RutaReporte As String

    If Esconsumo Then
        If TipoDoc = "BOLETA" Then
            RutaReporte = "C:\Admin\Nordi\BolCon.rpt"
        ElseIf TipoDoc = "FACTURA" Then
            RutaReporte = "C:\Admin\Nordi\FacCon.rpt"
            vSubTotal = Round((Me.lbltotal2.Caption / ((100 + LK_IGV) / 100)), 2)
            'vSubTotal = Round((Me.lblImporte.Caption / ((100 + LK_IGV + 5) / 100)), 2)
            'vrec = Round(vSubTotal * 0.05, 2)
            vIgv = Me.lbltotal2.Caption - vSubTotal
            'vIgv = Me.lblImporte.Caption - vSubTotal - vrec
        ElseIf TipoDoc = "NOTA DE PEDIDO" Then
            RutaReporte = "C:\Admin\Nordi\FacDet.rpt"
            vSubTotal = Round((Me.lbltotal2.Caption / ((100 + LK_IGV) / 100)), 2)
            'vSubTotal = Round((Me.lblImporte.Caption / ((100 + LK_IGV + 5) / 100)), 2)
            'vrec = Round(vSubTotal * 0.05, 2)
            vIgv = Me.lbltotal2.Caption - vSubTotal
            'vIgv = Me.lblImporte.Caption - vSubTotal - vrec
        End If

    Else

        If TipoDoc = "BOLETA" Then
            RutaReporte = "C:\Admin\Nordi\BolDet.rpt"
        ElseIf TipoDoc = "FACTURA" Then
            RutaReporte = "C:\Admin\Nordi\FacDet.rpt"
            vSubTotal = Round((Me.lbltotal2.Caption / ((100 + LK_IGV) / 100)), 2)
            'vSubTotal = Round((Me.lblImporte.Caption / ((100 + LK_IGV + 5) / 100)), 2)
            'vrec = Round(vSubTotal * 0.05, 2)
            vIgv = Me.lbltotal2.Caption - vSubTotal
            'vIgv = Me.lblImporte.Caption - vSubTotal - vrec
        ElseIf TipoDoc = "NOTA DE PEDIDO" Then
            RutaReporte = "C:\Admin\Nordi\FacDet.rpt"
            vSubTotal = Round((Me.lbltotal2.Caption / ((100 + LK_IGV) / 100)), 2)
            'vSubTotal = Round((Me.lblImporte.Caption / ((100 + LK_IGV + 5) / 100)), 2)
            'vrec = Round(vSubTotal * 0.05, 2)
            vIgv = Me.lbltotal2.Caption - vSubTotal
            'vIgv = Me.lblImporte.Caption - vSubTotal - vrec
        End If
    End If

    'If TipoDoc = "B" Then
  
    'Else
    '    oCmdEjec.CommandText = "SpPrintFacDet"
    'End If

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    Set crParamDefs = VReporte.ParameterFields

    For Each crParamDef In crParamDefs

        Select Case crParamDef.ParameterFieldName

            Case "cliente"

                'crParamDef.AddCurrentValue IIf(Len(Trim(Me.txtRS.Text)) = 0, "CLIENTES VARIOS", Trim(Me.txtRS.Text))
                'crParamDef.AddCurrentValue IIf(Len(Trim(Me.txtRS.Text)) = 0, Trim(Me.txtcli.Text), Trim(Me.txtRS.Text))
                'crParamDef.AddCurrentValue Trim(frmDeliveryApp.txtCliente.Text)
                crParamDef.AddCurrentValue Trim(txtCliente.Text)

            Case "FechaEmi"
                crParamDef.AddCurrentValue LK_FECHA_DIA

            Case "Son"
                crParamDef.AddCurrentValue CONVER_LETRAS(Me.lbltotal2.Caption, "S")

            Case "total"
                crParamDef.AddCurrentValue Me.lbltotal2.Caption

            Case "subtotal"
                crParamDef.AddCurrentValue CStr(vSubTotal)

            Case "igv"
                crParamDef.AddCurrentValue CStr(vIgv)

            Case "SerFac"
                crParamDef.AddCurrentValue Me.lblSerie.Caption

            Case "NumFac"
                crParamDef.AddCurrentValue Me.txtNumero.Text

            Case "DirClie"

                '                crParamDef.AddCurrentValue frmDeliveryApp.DatDireccion.Text
                crParamDef.AddCurrentValue Me.txtDireccion.Text

            Case "RucClie"

                'crParamDef.AddCurrentValue Vdocto
                crParamDef.AddCurrentValue Me.txtRuc.Text

            Case "Importe" 'linea nueva
                crParamDef.AddCurrentValue frmDeliveryApp.lblTot.Caption 'linea nueva
                ' Case "rec"          'SR BEFE
                '     crParamDef.AddCurrentValue CStr(vrec)

        End Select

    Next

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc

    'oCmdEjec.CommandText = "SpPrintComanda"

    Dim rsd As ADODB.Recordset

    oCmdEjec.CommandText = "SpPrintFacturacion"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Serie", adChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@nro", adInteger, adParamInput, , Me.txtNumero.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fbg", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1)) ' IIf(Me.ComDocto.ListIndex = 0, "F", IIf(Me.ComDocto.ListIndex = 1, "B", "")))
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NroCom", adInteger, adParamInput, , vNroCom)

    Set rsd = oCmdEjec.Execute

    'COCINA
    'rsd.Filter = "PED_FAMILIA=2"
    Dim DD As ADODB.Recordset

    ' For i = 0 To Printers.count - 1
    '        MsgBox Printers(i).DeviceName
    '    Next
    If Not rsd.EOF Then

        VReporte.DataBase.SetDataSource rsd, 3, 1 'lleno el objeto reporte
        'VReporte.SelectPrinter Printer.DriverName, "\\laptop\doPDF v6", Printer.Port
        '
        VReporte.PrintOut False, 1, , 1, 1
        frmVisor.cr.ReportSource = VReporte
        'frmVisor.cr.ViewReport
        'frmVisor.Show vbModal
    
    End If

    'Set VReporte = Nothing
    'Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    'cr.DataSource = VReporte
    'cr.Destination = crptToWindow
    '
    'rsd.Filter = "PED_FAMILIA=3"
    'If Not rsd.EOF Then
    '    VReporte.Database.SetDataSource rsd, 3, 1 'lleno el objeto reporte
    '    VReporte.SelectPrinter Printer.DriverName, "\\SERVIDOR\Canon MP140 series Printer", Printer.Port 'doPDF v6
    '    VReporte.PrintOut ' , 1, , 1, 1
    'End If

    Set objCrystal = Nothing
    Set VReporte = Nothing

    Exit Sub

printe:
    MostrarErrores Err

End Sub

Private Sub cmdAceptar_Click()

    If Me.DatRepartidor.BoundText = -1 Then
        MsgBox "Debe elegir el repartidor.", vbCritical, Pub_Titulo
        Me.DatRepartidor.SetFocus

        Exit Sub

    End If

    If Me.DatTiposDoctos.BoundText = "" Then
        MsgBox "Debe elegir el Documento.", vbInformation, Pub_Titulo
        Me.DatTiposDoctos.SetFocus

        Exit Sub

    End If

    If Len(Trim(Me.txtNumero.Text)) = 0 Then
        MsgBox "Debe ingresar el Número del documento.", vbCritical, Pub_Titulo
        Me.txtNumero.SetFocus

        Exit Sub

    End If

    If Not IsNumeric(Me.txtNumero.Text) Then
        MsgBox "El número del documento es incorrecto.", vbInformation, Pub_Titulo
        Me.txtNumero.SetFocus
        Me.txtNumero.SelStart = 0
        Me.txtNumero.SelLength = Len(Me.txtNumero.Text)

        Exit Sub

    End If

    If Not IsNumeric(Me.txtAbono.Text) Then
        MsgBox "El Abono no es correcto.", vbCritical, Pub_Titulo

        Exit Sub

    End If
    
    If val(Me.lblVuelto.Caption) < 0 Then
        MsgBox "El Abono es incorrecto.", vbCritical, Pub_Titulo
        Me.txtAbono.SetFocus
        Me.txtAbono.SelStart = 0
        Me.txtAbono.SelLength = Len(Me.txtAbono.Text)

        Exit Sub

    End If
    
    If Me.DatTiposDoctos.BoundText = "" And Len(Trim(Me.txtRuc.Text)) = 0 Then
        MsgBox "Debe ingresar el Nro de RUC.", vbInformation, Pub_Titulo

        Exit Sub

    End If
    
    If Me.DatTiposDoctos.BoundText = "01" Then
        If Len(Trim(Me.txtRuc.Text)) = 0 Then
            MsgBox "Debe ingresar el Ruc para poder generar la Factura", vbInformation, "Error"

            Exit Sub

        End If
    End If
    
    If Me.DatTiposDoctos.BoundText = "" And Len(Trim(Me.txtRuc.Text)) < 11 Then
        MsgBox "Ruc incorrecto", vbInformation, Pub_Titulo

        Exit Sub

    End If
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_VALIDA_UIT"

    Dim ORSuit As ADODB.Recordset

    Set ORSuit = oCmdEjec.Execute(, Me.lbltotal2.Caption)

    If Not ORSuit.EOF Then
        If ORSuit!Dato = 1 Then
          If Len(Trim(Me.txtCliente.Text)) = 0 Or (Len(Trim(Me.txtRuc.Text)) = 0 And Len(Trim(Me.txtDni.Text)) = 0) Then
            MsgBox "El Importe sobrepasa media UIT, debe ingresar el cliente.", vbCritical, Pub_Titulo

            Exit Sub
          End If
        End If
    End If
    

    Dim xEXITO         As String

    Dim xCONTINUA      As Boolean

    Dim xARCENCONTRADO As Boolean

    xCONTINUA = False
    xARCENCONTRADO = False

    xEXITO = ""
   
    'VALIDANDO SI EL ARCHIVO DEL REPORTE EXISTE
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_ARCHIVO_PRINT"
    oCmdEjec.CommandType = adCmdStoredProc
    
    Dim ORSd        As ADODB.Recordset

    Dim RutaReporte As String

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGO", adChar, adParamInput, 2, Me.DatTiposDoctos.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@COMSUMO", adBoolean, adParamInput, , Me.chkConsumo.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    
    Set ORSd = oCmdEjec.Execute
    RutaReporte = PUB_RUTA_REPORTE & ORSd!ReportE
    
    If ORSd!ReportE = "" Then
        If MsgBox("El Archivo no existe." & vbCrLf & "¿Desea continuar sin imprimir?.", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
            xCONTINUA = True
        End If

    Else
        FileName = dir(RutaReporte)

        If FileName = "" Then
            If MsgBox("El Archivo no existe." & vbCrLf & "¿Desea continuar sin imprimir?.", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
                xCONTINUA = True
            Else

                Exit Sub

            End If

        Else
            xCONTINUA = True
            xARCENCONTRADO = True
            'ImprimirDocumentoVenta Me.DatTiposDoctos.BoundText, Me.DatTiposDoctos.Text, Me.chkConsumo.Value, "009", 2, 100, "33", "43", "343"
        End If
    End If
    
    If xCONTINUA Then
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_DELIVERY_ENVIAR"
        oCmdEjec.CommandType = adCmdStoredProc
    
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adVarChar, adParamInput, 3, frmDeliveryApp.lblSerie.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adBigInt, adParamInput, , frmDeliveryApp.lblNumero.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PAGO", adDouble, adParamInput, , Me.txtAbono.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIRECCION", adVarChar, adParamInput, 150, Me.DatDireccion.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 20, LK_CODUSU)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERDOC", adVarChar, adParamInput, 3, Trim(Me.lblSerie.Caption))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NRODOC", adBigInt, adParamInput, , Me.txtNumero.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1)) ' IIf(Me.ComDocto.ListIndex = 0, "F", IIf(Me.ComDocto.ListIndex = 1, "B", "P")))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCLI", adInteger, adParamInput, , Trim(Me.lblcodclie.Caption))
        'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCLI", adInteger, adParamInput, , Trim(Me.txtRuc.Tag))
        'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCLI", adInteger, adParamInput, , frmDeliveryApp.lblCliente.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@totalfac", adDouble, adParamInput, , Me.lbltotal2.Tag)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@farjabas", adTinyInt, adParamInput, , IIf(Me.chkConsumo.Value = 1, 1, 0))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDREPARTIDOR", adBigInt, adParamInput, , CDbl(Me.DatRepartidor.BoundText))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGODOCTO", adChar, adParamInput, 2, Me.DatTiposDoctos.BoundText)
        
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DSCTO", adDouble, adParamInput, , lblDscto.Caption)
         oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NOMBRE_CLI", adVarChar, adParamInput, 100, Me.txtCliente.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RUC_CLI", adVarChar, adParamInput, 11, Me.txtRuc.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TARIFA", adDouble, adParamInput, , Me.lblMovilidad.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ICBPER", adDouble, adParamInput, , Me.lblicbper.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@EXITO", adVarChar, adParamOutput, 300, xEXITO)
    
        oCmdEjec.Execute
    
        xEXITO = oCmdEjec.Parameters("@EXITO").Value
        
        Dim xxIgv As String
        xxIgv = val(Me.lbltotal2.Caption) - Round((val(Me.lbltotal2.Caption) / val((LK_IGV / 100) + 1)), 2)

        If Len(Trim(xEXITO)) = 0 Then
            vGraba = True
            'MsgBox "Datos almacenados Correctamente.", vbInformation, Pub_Titulo
            CreaCodigoQR "6", Me.DatTiposDoctos.BoundText, Me.lblSerie.Caption, Me.txtNumero.Text, LK_FECHA_DIA, xxIgv, Me.lbltotal2.Caption, Me.txtRuc.Text, Me.txtDni.Text
            'Imprimir Left(Me.DatTiposDoctos.Text, 1), Me.chkConsumo.Value
            If xARCENCONTRADO Then
                ImprimirDocumentoVenta Me.DatTiposDoctos.BoundText, Me.DatTiposDoctos.Text, Me.chkConsumo.Value, Me.lblSerie.Caption, Me.txtNumero.Text, Me.lbltotal2.Caption, 0, 0, Me.txtDireccion.Text, Me.txtRuc.Text, Me.txtCliente.Text, Me.txtDni.Text, LK_CODCIA, Me.lblicbper.Caption, Me.chkprom.Value, txtcopias.Text
            End If
             
           ' If Me.DatTiposDoctos.BoundText = "01" Then
            If LK_PASA_BOLETAS = "A" And (Me.DatTiposDoctos.BoundText = "01" Or Me.DatTiposDoctos.BoundText = "03") Then
                 CrearArchivoPlano Left(Me.DatTiposDoctos.Text, 1), Me.lblSerie.Caption, Me.txtNumero.Text
            ElseIf Me.DatTiposDoctos.BoundText = "01" Then
            CrearArchivoPlano Left(Me.DatTiposDoctos.Text, 1), Me.lblSerie.Caption, Me.txtNumero.Text
            End If
            Unload Me
        Else
            MsgBox xEXITO, vbCritical, Pub_Titulo
        End If
    End If

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdSunat_Click()

     On Error GoTo cCruc

Dim p          As Object

Dim Texto      As String, xTOk As String

Dim cadena     As String, xvRUC As String

Dim sInputJson As String, xEsRuc As Boolean

xEsRuc = True

MousePointer = vbHourglass
Set httpURL = New WinHttp.WinHttpRequest
    
If IsNumeric(Me.txtDni.Text) Then
        If Len(Trim(Me.txtDni.Text)) = 8 Then
            xEsRuc = False
        End If

        xvRUC = Me.txtDni.Text
    Else

        If Len(Trim(Me.txtRuc.Text)) = 11 Then
            xEsRuc = True
        End If

        xvRUC = Me.txtRuc.Text
    End If
       
  '  xvRUC = Me.txtCliente.Text

xTOk = Leer_Ini(App.Path & "\config.ini", "TOKEN", "")
    
If xEsRuc Then
        cadena = "http://dniruc.apisperu.com/api/v1/ruc/" & xvRUC & "?token=" & xTOk
    Else
        cadena = "http://dniruc.apisperu.com/api/v1/dni/" & xvRUC & "?token=" & xTOk
    End If
   ' cadena = "http://dniruc.apisperu.com/api/v1/ruc/" & xvRUC & "?token=" & xTOk

    
httpURL.Open "GET", cadena
httpURL.Send
    
Texto = httpURL.ResponseText

'sInputJson = "{items:" & Texto & "}"

Set p = JSON.parse(Texto)
    
If Texto = "[]" Then
    MousePointer = vbDefault
    MsgBox ("No se obtuvo resultados")
    Me.txtRuc.Text = ""
    Me.txtCliente.Text = ""
    

    Exit Sub

End If

If Len(Trim(Texto)) = 0 Then
    MousePointer = vbDefault
    MsgBox ("No se obtuvo resultados")
    Me.txtRuc.Text = ""
    Me.txtCliente.Text = ""
    

    Exit Sub

End If



        Me.txtDireccion.Text = p.Item("direccion")
        Me.txtCliente.Text = p.Item("razonSocial")
        Me.txtRuc.Text = p.Item("ruc")
        Me.txtDni.Text = p.Item("dni")
        Me.lblcodclie.Caption = "0"
       
MousePointer = vbDefault

Exit Sub

cCruc:
MousePointer = vbDefault
MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub DatDireccion_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Me.txtAbono.SetFocus
        Me.txtAbono.SelStart = 0
        Me.txtAbono.SelLength = Len(Me.txtAbono.Text)
    End If
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub DatRepartidor_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.DatTiposDoctos.SetFocus
End Sub

Private Sub DatTiposDoctos_Click(Area As Integer)
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_SERIES_CARGAR"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim xSerie As String

    Dim xNro   As Double

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1)) ' Me.cboTipoDocto.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamOutput, 3, 1)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MAXIMO", adBigInt, adParamOutput, , 1)
    oCmdEjec.Execute

    xSerie = oCmdEjec.Parameters("@SERIE").Value
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)

    Me.lblSerie.Caption = xSerie
    Me.txtNumero.Text = oCmdEjec.Parameters("@MAXIMO").Value

    ORStd.Filter = "CODIGO='" & Me.DatTiposDoctos.BoundText & "'"

    If ORStd.RecordCount <> 0 Then
        Me.chkEdit.Enabled = ORStd!Editable
    End If

    ORStd.Filter = ""
    Me.txtNumero.Enabled = False
    Me.chkEdit.Value = False
End Sub

Private Sub DatTiposDoctos_KeyDown(KeyCode As Integer, Shift As Integer)
LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_SERIES_CARGAR"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim xSerie As String

    Dim xNro   As Double

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1)) ' Me.cboTipoDocto.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamOutput, 3, 1)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MAXIMO", adBigInt, adParamOutput, , 1)
    oCmdEjec.Execute

    xSerie = oCmdEjec.Parameters("@SERIE").Value
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)

    Me.lblSerie.Caption = xSerie
    Me.txtNumero.Text = oCmdEjec.Parameters("@MAXIMO").Value

    ORStd.Filter = "CODIGO='" & Me.DatTiposDoctos.BoundText & "'"

    If ORStd.RecordCount <> 0 Then
        Me.chkEdit.Enabled = ORStd!Editable
    End If

    ORStd.Filter = ""
    Me.txtNumero.Enabled = False
    Me.chkEdit.Value = False
End Sub

Private Sub DatTiposDoctos_KeyUp(KeyCode As Integer, Shift As Integer)
LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_SERIES_CARGAR"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim xSerie As String

    Dim xNro   As Double

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1)) ' Me.cboTipoDocto.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamOutput, 3, 1)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MAXIMO", adBigInt, adParamOutput, , 1)
    oCmdEjec.Execute

    xSerie = oCmdEjec.Parameters("@SERIE").Value
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)

    Me.lblSerie.Caption = xSerie
    Me.txtNumero.Text = oCmdEjec.Parameters("@MAXIMO").Value

    ORStd.Filter = "CODIGO='" & Me.DatTiposDoctos.BoundText & "'"

    If ORStd.RecordCount <> 0 Then
        Me.chkEdit.Enabled = ORStd!Editable
    End If

    ORStd.Filter = ""
    Me.txtNumero.Enabled = False
    Me.chkEdit.Value = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub ConfiguraLV()
With Me.ListView1
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Codigo", 1000
    .ColumnHeaders.Add , , "Cliente", 2500
    .ColumnHeaders.Add , , "Ruc", 1000
    .ColumnHeaders.Add , , "Direccion", 0
    .ColumnHeaders.Add , , "DNI", 1000
    .MultiSelect = False
End With
End Sub

Private Sub Form_Load()
    ConfiguraLV
    Me.txtCliente.Text = Trim(frmDeliveryApp.txtCliente.Text)
    Me.txtDireccion.Text = Trim(frmDeliveryApp.DatDireccion.Text)
    Me.txtRuc.Text = Trim(frmDeliveryApp.lblRUC.Caption)
    Me.txtDni.Text = Trim(frmDeliveryApp.lblDNI.Caption)
    Me.lblcodclie.Caption = Trim(frmDeliveryApp.lblCliente.Caption)
    Me.lblicbper.Caption = Format(frmDeliveryApp.lblicbper.Caption, "#####0.#0")
    vGraba = False

    '    With Me.ComDocto
    '        .AddItem "FACTURA"
    '        .AddItem "BOLETA"
    '        .AddItem "NOTA DE PEDIDO"
    '    End With
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_TIPOS_DOCTOS_LIST"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    Set ORStd = oCmdEjec.Execute
    Set Me.DatTiposDoctos.RowSource = ORStd
    Me.DatTiposDoctos.ListField = ORStd.Fields(1).Name
    Me.DatTiposDoctos.BoundColumn = ORStd.Fields(0).Name

    If ORStd.RecordCount <> 0 Then
        Me.DatTiposDoctos.BoundText = ORStd!Codigo
    
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_SERIES_CARGAR"
        oCmdEjec.CommandType = adCmdStoredProc

        Dim xSerie As String

        Dim xNro   As Double

        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1)) ' Me.cboTipoDocto.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamOutput, 3, 1)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MAXIMO", adBigInt, adParamOutput, , 1)
        oCmdEjec.Execute

        xSerie = oCmdEjec.Parameters("@SERIE").Value
        'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)

        Me.lblSerie.Caption = xSerie
        Me.txtNumero.Text = oCmdEjec.Parameters("@MAXIMO").Value
        
        ORStd.Filter = "CODIGO='" & Me.DatTiposDoctos.BoundText & "'"

        If ORStd.RecordCount <> 0 Then
            Me.chkEdit.Enabled = ORStd!Editable
        End If

        ORStd.Filter = ""
    End If

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_DELIVERY_CARGAR_REPARTIDORES"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim oRSr As ADODB.Recordset

    Set oRSr = oCmdEjec.Execute(, LK_CODCIA)

    Set Me.DatRepartidor.RowSource = oRSr
    Me.DatRepartidor.ListField = oRSr.Fields(1).Name
    Me.DatRepartidor.BoundColumn = oRSr.Fields(0).Name
    Me.DatRepartidor.BoundText = -1

    Me.lblTotal.Tag = gTOTAL
    'Me.lblTotal.Caption = "S/. " + Format(gTOTAL, "#####0.00")
    Me.lblTotal.Caption = Format(gTOTAL, "#####0.00")
    
    Me.lblMovilidad.Tag = gMOVILIDAD
    'Me.lblMovilidad.Caption = "S/. " + Format(gMOVILIDAD, "#####0.00")
    Me.lblMovilidad.Caption = Format(gMOVILIDAD, "#####0.00")
    
    If pINCMOV Then
        Me.lbltotal2.Tag = gTOTAL + gMOVILIDAD - gDESCUENTO ' + Me.lblicbper.Caption
        'Me.lblTotal2.Caption = "S/. " + Format(Me.lblTotal2.Tag, "#####0.00")
        Me.lbltotal2.Caption = Format(Me.lbltotal2.Tag, "#####0.#0")
    Else
        Me.lbltotal2.Tag = gTOTAL - gDESCUENTO ' + Me.lblicbper.Caption
        'Me.lblTotal2.Caption = "S/. " + Format(Me.lblTotal2.Tag, "#####0.00")
        Me.lbltotal2.Caption = Format(Me.lbltotal2.Tag, "#####0.#0")
   
    End If
   
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_DELIVERY_DATOS_CLIENTE"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim orsP As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adVarChar, adParamInput, 3, frmDeliveryApp.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adBigInt, adParamInput, , frmDeliveryApp.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    Set orsP = oCmdEjec.Execute

    If Not orsP.EOF Then
        Me.txtAbono.Text = orsP!PAGO
        Me.lblVuelto.Caption = orsP!VUELTO
        Me.lblDscto.Caption = orsP!descuento
    End If
    
    'Me.txtAbono.Text = "0.00"
    Me.txtAbono.SelStart = 0
    Me.txtAbono.SelLength = Len(Me.txtAbono.Text)

    If val(Me.txtAbono.Text) = 0 Then Me.lblVuelto.Caption = "0.00"
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_DELIVERY_CLIENTE_DIRECCIONES"
            
    Dim oRSdir As ADODB.Recordset

    Set oRSdir = oCmdEjec.Execute(, Array(LK_CODCIA, frmDeliveryApp.lblCliente.Caption))
    Set Me.DatDireccion.RowSource = oRSdir
    Me.DatDireccion.BoundColumn = oRSdir.Fields(4).Name
    Me.DatDireccion.ListField = oRSdir.Fields(0).Name
            
    'Me.DatDireccion.BoundText = orsP.Fields(4).Value
    Me.DatDireccion.BoundText = gIDDIR
    vBuscar = True
End Sub

Private Sub txtAbono_Change()

'    If InStr(Me.txtAbono.Text, ".") Then
'        vPUNTO = True
'    Else
'        vPUNTO = False
'    End If
'
'    If IsNumeric(Me.txtAbono.Text) Then
'        Me.lblVuelto.Caption = val(Me.txtAbono.Text) - val(Me.lbltotal2.Tag)
'    Else
'        Me.lblVuelto.Caption = "0.00"
'    End If

End Sub

Private Sub txtAbono_KeyPress(KeyAscii As Integer)

    If NumerosyPunto(KeyAscii) Then KeyAscii = 0
    If KeyAscii = 46 Then
        If vPUNTO Or Len(Trim(Me.txtAbono.Text)) = 0 Then
            KeyAscii = 0
        End If
    End If

    If KeyAscii = vbKeyReturn Then cmdAceptar_Click
End Sub

Private Sub txtCliente_Change()
vBuscar = True
'Me.txtRuc.Text = ""
'Me.txtDireccion.Text = ""
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)

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
        Me.txtCliente.Text = ""
        Me.txtRuc.Text = ""
        Me.txtDireccion.Text = ""
    End If

    GoTo fin
posicion:
    ListView1.ListItems.Item(loc_key).Selected = True
    ListView1.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    txtCliente.SelStart = Len(Me.txtCliente.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
 KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        '    If loc_key > 0 Then
        '        Me.txtRuc.Tag = Me.ListView1.ListItems(loc_key)
        '        Me.txtRuc.Text = Me.ListView1.ListItems(loc_key).SubItems(2)
        '        Me.txtDireccion.Text = Me.ListView1.ListItems(loc_key).SubItems(3)
        '        Me.txtRS.Text = Me.ListView1.ListItems(loc_key).SubItems(1)
        '        Me.ListView1.Visible = False
        '        Me.lvDetalle.SetFocus
        '    End If

        If vBuscar Then
            Me.ListView1.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpListarCliProv"
            Set oRsPago = oCmdEjec.Execute(, Array(LK_CODCIA, "C", Me.txtCliente.Text))

            Dim Item As Object
        
            If Not oRsPago.EOF Then

                Do While Not oRsPago.EOF
                    Set Item = Me.ListView1.ListItems.Add(, , oRsPago!CodClie)
                    Item.SubItems(1) = Trim(oRsPago!Nombre)
                    Item.SubItems(2) = IIf(IsNull(oRsPago!RUC), "", oRsPago!RUC)
                    Item.SubItems(3) = Trim(oRsPago!dir)
                    Item.SubItems(4) = IIf(IsNull(oRsPago!DNI), "", oRsPago!DNI)
                    oRsPago.MoveNext
                Loop

                Me.ListView1.Visible = True
                Me.ListView1.ListItems(1).Selected = True
                loc_key = 1
                Me.ListView1.ListItems(1).EnsureVisible
                vBuscar = False
'            Else
'
'                If MsgBox("Cliente no existe." + vbCrLf + "¿Desea Crearlo.?", vbQuestion + vbYesNo, "Restaurantes") = vbYes Then
'                    frmCLI.Show vbModal
'                End If
            End If
        
        Else
           
            Me.txtCliente.Text = Me.ListView1.ListItems(loc_key).SubItems(1)
             Me.txtCliente.Tag = Me.ListView1.ListItems(loc_key)
            Me.txtRuc.Text = Me.ListView1.ListItems(loc_key).SubItems(2)
            Me.txtDireccion.Text = Me.ListView1.ListItems(loc_key).SubItems(3)
            Me.lblcodclie.Caption = Me.ListView1.ListItems(loc_key)
            Me.txtDni.Text = Me.ListView1.ListItems(loc_key).SubItems(4)
            Me.ListView1.Visible = False
            'Me.lvDetalle.SetFocus
        End If
    End If
End Sub

