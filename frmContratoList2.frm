VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#12.1#0"; "CODEJO~1.OCX"
Begin VB.Form frmContratoList2 
   Caption         =   "Form1"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   12435
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6330
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin XtremeCalendarControl.DatePicker dpMeses 
      Height          =   5775
      Left            =   9120
      TabIndex        =   1
      Top             =   0
      Width           =   2415
      _Version        =   786433
      _ExtentX        =   4260
      _ExtentY        =   10186
      _StockProps     =   64
      RowCount        =   2
   End
   Begin XtremeCalendarControl.CalendarControl ccDatos 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _Version        =   786433
      _ExtentX        =   13150
      _ExtentY        =   8705
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmContratoList2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Me.dpMeses.AttachToCalendar Me.ccDatos

    'Me.ccDatos.Options.TooltipAddNewText = "Agregar Contrato"
    Me.ccDatos.Options.EnableAddNewTooltip = False
    Me.ccDatos.Options.EnableInPlaceEditEventSubject_AfterEventResize = False
    Me.ccDatos.Options.EnableInPlaceEditEventSubject_ByMouseClick = False
    Me.ccDatos.Options.EnableInPlaceEditEventSubject_ByTab = False
    Me.ccDatos.Options.EnablePrevNextEventButtons = False
    Me.ccDatos.DayView.TimeScale = 60 'cada 20 minutos
    

   

Me.ccDatos.ViewType = xtpCalendarFullWeekView

End Sub
