VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Object = "{BAC8A387-2CA7-4372-ADF1-C1A1CC9A08D0}#1.0#0"; "prjXTab.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{749E5815-D8E7-4012-BF8D-C219AD48734F}#1.2#0"; "InnovaDemo1.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   -75
   ClientTop       =   225
   ClientWidth     =   14535
   ForeColor       =   &H00008000&
   Icon            =   "frmNDMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   14535
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "Define Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   12840
      Picture         =   "frmNDMain.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1080
      Width           =   1545
   End
   Begin VB.CommandButton abortBtn 
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6570
      Picture         =   "frmNDMain.frx":0DDD
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1980
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   14130
      TabIndex        =   23
      Top             =   5130
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton btnManualProcess 
      Caption         =   "Start capture"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14175
      Picture         =   "frmNDMain.frx":0F6A
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3600
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton searchBtn 
      BackColor       =   &H80000009&
      Caption         =   "continuar"
      Height          =   630
      Left            =   6570
      Picture         =   "frmNDMain.frx":1562
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1215
      Width           =   1440
   End
   Begin VB.CommandButton GoBtn 
      Caption         =   "GO!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   13995
      TabIndex        =   20
      Top             =   4320
      Visible         =   0   'False
      Width           =   1545
   End
   Begin FramePlusCtl.FramePlus FramePlus2 
      Height          =   1605
      Left            =   135
      TabIndex        =   13
      Top             =   1080
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   2831
      Style           =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Begin FramePlusCtl.FramePlus FramePlus1 
         Height          =   1305
         Left            =   135
         TabIndex        =   14
         Top             =   180
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   2302
         BorderStyle     =   4
         BackColor       =   14737632
         HighlightDkColor=   12632256
         Style           =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Search"
         Begin VB.Label lblMensajeSup 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   180
            TabIndex        =   17
            Top             =   870
            Width           =   5535
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   60
            Width           =   3465
         End
         Begin VB.Label lblMensaje 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Not Ready"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   270
            TabIndex        =   15
            Top             =   405
            Width           =   5235
         End
      End
   End
   Begin InnovaDemoCtls.dmoHTMLLabel htmlInfo 
      Height          =   1695
      Left            =   8235
      TabIndex        =   12
      Top             =   1035
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   2990
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   12
      BackColor       =   -2147483624
      Text            =   $"frmNDMain.frx":18DD
   End
   Begin VB.Timer LoadTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   30
   End
   Begin prjXTab.XTab XTabModo 
      Height          =   1410
      Left            =   840
      TabIndex        =   7
      Top             =   9990
      Visible         =   0   'False
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   2487
      TabCount        =   2
      TabCaption(0)   =   "Normal mode"
      TabContCtrlCnt(0)=   2
      Tab(0)ContCtrlCap(1)=   "searchtxt"
      Tab(0)ContCtrlCap(2)=   "Label3"
      TabCaption(1)   =   "Interactive mode"
      TabContCtrlCnt(1)=   2
      Tab(1)ContCtrlCap(1)=   "btnManualProcess2"
      Tab(1)ContCtrlCap(2)=   "lblMensaje20"
      ActiveTabHeight =   22
      InActiveTabHeight=   15
      TabTheme        =   3
      ActiveTabBackStartColor=   16316664
      InActiveTabBackStartColor=   15066597
      ActiveTabForeColor=   16711680
      InActiveTabForeColor=   0
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   9474192
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   8421504
      XRadius         =   15
      YRadius         =   15
      Begin VB.CommandButton btnManualProcess2 
         BackColor       =   &H80000009&
         Caption         =   "Start capture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -69690
         Picture         =   "frmNDMain.frx":1959
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   540
         Width           =   2025
      End
      Begin VB.TextBox searchtxt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   225
         TabIndex        =   8
         Top             =   855
         Width           =   5565
      End
      Begin VB.Label lblMensaje2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Detail page detected!!!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Index           =   0
         Left            =   -74595
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   4740
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Keywords to search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   225
         TabIndex        =   9
         Top             =   540
         Width           =   1995
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   150
      Top             =   0
   End
   Begin VB.CommandButton save2XLS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "Save to CSV"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12840
      Picture         =   "frmNDMain.frx":1CD4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1545
   End
   Begin VB.CommandButton eraseBtn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "CLEAR RESULTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   12840
      Picture         =   "frmNDMain.frx":1DF7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   1545
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   7605
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Visible         =   0   'False
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "7/18/2009"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "1:43 PM"
         EndProperty
      EndProperty
   End
   Begin prjXTab.XTab XTab1 
      Height          =   4305
      Left            =   90
      TabIndex        =   3
      Top             =   3060
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   7594
      TabCaption(0)   =   "   Event log   "
      TabContCtrlCnt(0)=   1
      Tab(0)ContCtrlCap(1)=   "List1"
      TabCaption(1)   =   "   Captured results    "
      TabContCtrlCnt(1)=   1
      Tab(1)ContCtrlCap(1)=   "DataGrid1"
      TabCaption(2)   =   "   Debugging   "
      TabContCtrlCnt(2)=   2
      Tab(2)ContCtrlCap(1)=   "WebBrowserView"
      Tab(2)ContCtrlCap(2)=   "ImgBlock"
      ActiveTab       =   1
      ActiveTabHeight =   25
      TabStyle        =   1
      TabTheme        =   2
      InActiveTabBackStartColor=   -2147483626
      InActiveTabBackEndColor=   -2147483626
      ActiveTabForeColor=   16711680
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   -2147483628
      TabStripBackColor=   16777215
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483632
      Begin SHDocVwCtl.WebBrowser WebBrowserView 
         Height          =   3525
         Left            =   -74850
         TabIndex        =   6
         Top             =   630
         Width           =   6105
         ExtentX         =   10769
         ExtentY         =   6218
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3525
         Left            =   150
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   630
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   6218
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Enabled         =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   2
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Information captured"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.ListBox List1 
         Height          =   3375
         Left            =   -74850
         TabIndex        =   4
         Top             =   690
         Width           =   5865
      End
      Begin VB.Image ImgBlock 
         Enabled         =   0   'False
         Height          =   795
         Left            =   -74850
         Top             =   600
         Visible         =   0   'False
         Width           =   6375
      End
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar 
      Height          =   315
      Left            =   8460
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "5"
      Top             =   495
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   556
      Picture         =   "frmNDMain.frx":1F0A
      BackColor       =   14737632
      ForeColor       =   0
      Appearance      =   2
      BorderStyle     =   2
      BarColor        =   16711680
      BarPicture      =   "frmNDMain.frx":1F26
      Max             =   66
      ShowText        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operation in progress"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   270
      Left            =   9540
      TabIndex        =   18
      Top             =   180
      Width           =   2055
   End
   Begin VB.Image Image3 
      Height          =   1995
      Left            =   -90
      Picture         =   "frmNDMain.frx":1F42
      Stretch         =   -1  'True
      Top             =   945
      Width           =   14760
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   540
      Picture         =   "frmNDMain.frx":30B1
      Top             =   420
      Width           =   3585
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   4890
      Picture         =   "frmNDMain.frx":39BA
      Stretch         =   -1  'True
      Top             =   420
      Width           =   360
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1365
      Left            =   0
      Top             =   0
      Width           =   15405
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   6990
      Left            =   -60
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   16440
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const navNoHistory = &H2
Const navNoReadFromCache = &H4
Const navNoWriteToCache = &H8

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private WithEvents WebBrowser1 As SHDocVw.InternetExplorer
Attribute WebBrowser1.VB_VarHelpID = -1
Private WithEvents WebBrowserDet As SHDocVw.InternetExplorer
Attribute WebBrowserDet.VB_VarHelpID = -1
Private WithEvents WebBrowserMain As SHDocVw.InternetExplorer
Attribute WebBrowserMain.VB_VarHelpID = -1

Dim LinksParsed As LinkClassCollection

'Dim rsRes As clsResultSource
Dim rst As ADODB.Recordset
Dim rsState As ADODB.Recordset

Dim urlActualID As String
Dim ultimoUrl As String

Const dominio = "www.yellowpages.com"
Const mPagina = "http://www.yellowpages.com/"
Const mPostPage = "http://www.yellowpages.com.au/"

Dim actionPage As String
Dim postString As String
Dim paginaPOST As String

Dim navTimeStamp As String
Dim puedeContinuar As Boolean
Dim loadTimerSeg As Long

Dim inicioProceso As String

Dim resultsTxt As String

Dim CookiesCls As Collection
  

Sub actualizarRuntimeInfo()

'Dim anio As String

'If cmbYears = "" Then
    'anio = Year(Now)
'Else
    'anio = Config("lastSelectedYear")
'End If

'processedCat = LinksParsed.Count - LinksParsed.unprocesedItems

htmlInfo.Text = "Detected results:<br><b>" & resultsTxt & "</b><br>" & _
                "Captures records:<br><b>" & rst.RecordCount & "</b><br>" & _
                "Process started on: <br><b>" & inicioProceso & "</b>"
End Sub



Private Sub abortBtn_Click()

'urlActualID = "Stoped"

        With lblMensaje
            .BackColor = &H0&
            .caption = "Ready"
        End With

setReasumirBusqueda True

btnManualProcess.enabled = True
abortBtn.enabled = False
searchBtn.enabled = True
LinksParsed.enabled = False
Timer1.enabled = False

WebBrowser1.stop

logger ""
logger "Searching process interrupted by the user...."
logger ""


'cargarPagina ultimoUrl, "Manual" 'urlActualID 'analisisResult WebBrowser1.document


StatusBar1.Panels(1).Text = "Searching process interrupted by the user...."
End Sub

Private Sub btnManualProcess_Click()
urlActualID = "Busqueda"
    
  inicioProceso = Now

  abortBtn.enabled = True
  btnManualProcess.enabled = False
  searchBtn.enabled = False
  
'        lblMensaje.Caption = "Capturing records (" & rsState.Fields("state").Value & "...."
'        lblMensaje.ForeColor = &H8000&    'RGB(0, 255, 0)
'        lblMensajeSup.Caption = ""
'
  
  'logger "Starting manual processing..."
  
  StatusBar1.Panels(1).Text = "analyzing process....waiting........"
'  analisisPrimerResultado WebBrowser1.document
  
 analizarPagina WebBrowser1.document
  
  actualizarRuntimeInfo
  
End Sub



Private Sub Command2_Click()
    frmSV.oCon = Config("connectionstring")
    frmSV.Show 1
End Sub

'=============================================================
'                    L o a d   F o r m
'=============================================================

Private Sub Form_Load()
setReasumirBusqueda False


'Label1.BackColor = RGB(253, 247, 82)
urlActualID = ""
searchBtn.enabled = False

Set LinksParsed = New LinkClassCollection
'Set rsRes = New clsResultSource
Set CookiesCls = New Collection


XTab1.ActiveTab = 0
XTabModo.ActiveTab = 0
redrawGrid

XTabModo.TabEnabled(1) = False
XTab1.TabEnabled(2) = False
    
'------------ configuracion-------------

configInicial

' windows position
With Me

    .Left = Config("startx")
    .Top = Config("starty")
    .Width = Config("width")
    .Height = Config("height")
    .WindowState = Config("state")
    
End With

'------------ configuracion-------------

'----------------------------------------
' objetos de acceso a internet
'----------------------------------------
createIExplorerObject

Set rst = New ADODB.Recordset

With rst
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .open "select * from data", Config("connectionstring")
End With


Set rsState = New ADODB.Recordset
With rsState
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .open "select * from statesUSA where parsed=true", Config("connectionstring")
    
    If Not .EOF And Not .BOF Then .MoveFirst
End With

Set DataGrid1.DataSource = rst 'rsRes.RS
redrawGrid


' session persistente activada??
If Config("persistentSession") = 1 Then
    Select Case MsgBox("Do you want to try to restore a previous stored session (if it exists)?" & vbNewLine & vbNewLine _
     & "A previous session keeps the information about the selected states and the last url/ type in analysis.", _
    vbYesNo Or vbQuestion Or vbSystemModal Or vbDefaultButton1, App.Title)
    Case vbYes
            LinksParsed.setPersistentStorage App.Path & persistentFile
    End Select
End If

''-----------------------------------------------------------------
'' Si esta en modo debugging (3), le asocio el WebbrowserControl
''-----------------------------------------------------------------
'If Config("debugLevel") > 2 Then
'    Set WebBrowser1 = WebBrowserView
'Else
'    Set WebBrowser1 = New SHDocVw.InternetExplorer
'End If
'
'Do Until WebBrowser1.Busy = False
'        DoEvents
'Loop
'
'WebBrowser1.Silent = True

createIExplorerObject

caption = appTitle
'Skinner1.BoxesDefaultTitle = appTitle

cargarPagina mPagina, "pagInicial"

logger "Application starting....."
StatusBar1.Panels(1).Text = "Application is starting, please wait......."

End Sub

'---------------------------------------------------------------------------------------
' Procedure : createIExplorerObject
' DateTime  : 05/02/2007 20:37
' Author    : Administrador
' Purpose   : Objetos varios como instancias del Webbrowser y/o SHDocVw.InternetExplorer
'---------------------------------------------------------------------------------------
'
Sub createIExplorerObject()

'-----------------------------------------------------------------
' Si esta en modo debugging (3), le asocio el WebbrowserControl
'-----------------------------------------------------------------


   On Error GoTo createIExplorerObject_Error
   '---------------------------------------------------------------------------------------------------------------

If Config("debugLevel") > 1 Then
    Set WebBrowserMain = WebBrowserView
    XTab1.TabEnabled(2) = True
Else
'    Set WebBrowserMain = New SHDocVw.InternetExplorer
    Set WebBrowserMain = WebBrowserView
    XTab1.TabEnabled(2) = False
    'WebBrowserView.Visible = False
    WebBrowserView.Visible = True
End If

Set WebBrowser1 = WebBrowserMain
Do Until WebBrowser1.Busy = False
    DoEvents
Loop

'XTab1.TabEnabled(2) = False
'WebBrowserView.Visible = True

'Set WebBrowserDet = New SHDocVw.InternetExplorer
'Do Until WebBrowserDet.Busy = False
'    DoEvents
'Loop

WebBrowserView.Silent = True


  '---------------------------------------------------------------------------------------------------------------
   On Error GoTo 0
   Exit Sub

createIExplorerObject_Error:

   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createIExplorerObject of Formulario frmMain"
'    ErrorHandler Err, "createIExplorerObject", "Formulario", "frmMain"


End Sub


Sub redrawGrid()
Dim col As Column

For Each col In DataGrid1.Columns
    col.Width = DataGrid1.Width / DataGrid1.Columns.Count
Next

End Sub

Private Sub setReasumirBusqueda(estado As Boolean)

    'Exit Sub
    'urlActualID = "pagInicial"
    
    puedeContinuar = estado
    If estado = True Then
        searchBtn.caption = "search/resume"
    Else
        searchBtn.caption = "Search"
    End If
End Sub

Private Sub Form_Resize()
    
   On Error GoTo Form_Resize_Error

    If Me.WindowState = 1 Then Exit Sub
    
    XTab1.Width = Me.Width - (XTab1.Left * 3)
    XTab1.Height = Me.Height - XTab1.Top - 800
    
    Image3.Width = Me.Width
    Shape2.Width = Me.Width

    Select Case XTab1.ActiveTab
        
        Case 0
            
            List1.Width = XTab1.Width - (List1.Left * 2)
            List1.Height = XTab1.Height - List1.Top - List1.Left
        
        Case 1
            
            DataGrid1.Width = XTab1.Width - (DataGrid1.Left * 2)
            DataGrid1.Height = XTab1.Height - DataGrid1.Top - DataGrid1.Left
            redrawGrid
            
        Case 2
        
        
            WebBrowserView.Width = XTab1.Width - (WebBrowserView.Left * 2)
            WebBrowserView.Height = XTab1.Height - WebBrowserView.Top - WebBrowserView.Left
            
'            ImgBlock.Top = WebBrowserView.Top
'            ImgBlock.Left = WebBrowserView.Left
'            ImgBlock.Width = WebBrowserView.Width
'            ImgBlock.Height = WebBrowserView.Height
    
    End Select
    
    

   On Error GoTo 0
   Exit Sub

Form_Resize_Error:

 '   MsgBox "Error " & Err.Number & " (" & Err.Description & ") en el procedimiento Form_Resize del Formulario frmMain"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.enabled = False
    WebBrowser1.stop
    Set WebBrowser1 = Nothing
    
    logger "Application ended."
    
    With Me
         Config("startx") = .Left
         Config("starty") = .Top
        Config("width") = .Width
         Config("height") = .Height
        Config("state") = .WindowState
    End With
    
    Config.saveToArchive
End Sub


Private Sub Image4_Click()
frmSplash.Show vbModal
'frmAbout.Show vbModal
End Sub


Private Sub GoBtn_Click()
LinksParsed.enabled = True
armarTypeLinkList
End Sub

Private Sub save2XLS_Click()
Dim result As Boolean
Dim res As Integer

    result = exportToCSV(rst, True)

    
    If result Then
        
'        res = MsgBox("Do you want to discard the present result so that a new search starts with a empty grid??", vbInformation + vbYesNo, appTitle)
'        If res = vbYes Then eraseBtn_Click
        MsgBox "Exporting tasks complete.", vbInformation, appTitle
    Else
        
        MsgBox "The operation was canceled", vbInformation, appTitle
    
    End If
End Sub

Private Sub searchBtn_Click()
On Error GoTo mError
Dim res As Integer

abortBtn.enabled = True
searchBtn.enabled = False
LinksParsed.enabled = True
   
     'cargarPagina ultimoUrl, urlActualID
        
        
    lblMensaje.caption = "Capturing records...."
    lblMensaje.ForeColor = &H8000&    'RGB(0, 255, 0)
    lblMensajeSup.caption = "" 'LinksParsed.categoriaActual
    
        
        'Exit Sub
    '------------------------------------------------
    ' Soporta resume una busqueda anterior????
    If puedeContinuar Then
    
        res = MsgBox("Do you want to resume the previous search?" & vbNewLine & vbNewLine _
                & "If you accept, the search will follow in the point that was canceled previously." & vbNewLine _
                & "If you choose 'No', for a new search.", _
                vbInformation + vbYesNoCancel, appTitle)
        
        If res = vbYes Then
            
            lblMensajeSup.caption = "[" & rsState.Fields("state").Value & "] -" & LinksParsed.categoriaActual

            ' continuo con la ultima pagina.........
            cargarPagina ultimoUrl, urlActualID
            Exit Sub
        End If
        
        If res = vbCancel Then GoTo operacionCancelada
    End If
    '------------------------------------------------
    
   setReasumirBusqueda False
    
    If rst.RecordCount > 0 Then
'        res = MsgBox(vbNewLine & "Do you want to discard the present results so that a new search starts with a empty grid??" & vbNewLine, vbQuestion + vbYesNoCancel, appTitle)
 '
  '      If res = vbYes Then eraseBtn_Click
        If eraseResults = False Then GoTo operacionCancelada
   End If


'vacio la coleccion de links
While LinksParsed.Count > 0
    LinksParsed.clear
Wend

selectStates

armarTypeLinkList

Exit Sub

'----------------------------------------------------------------------


'armo el string de parametros de busqueda
Dim strPost As String

Dim miURL  As String
'miURL = "http://www.yellowpages.com.au/search/postSearchEntry.do?clueType=0&clue=" & searchtxt.Text & "&locationClue="
'miURL = URLEncode(miURL)


miURL = "http://www.yellowpages.com.au/search/postSearchEntry.do?clueType=0&clue=" & URLEncode(searchtxt.Text) & "&locationClue="

cargarPagina miURL, "Busqueda"

'ogger "Opening page '" & paginaPOST & " " & strPost & "'."

logger "Opening page '" & miURL & "'."
StatusBar1.Panels(1).Text = "Loading " & miURL & "......."

'logger "Unable to get necessary information from the page. Please, review if the original site is donw."

Exit Sub

'....................................

operacionCancelada:

abortBtn.enabled = False
searchBtn.enabled = True
LinksParsed.enabled = False

Exit Sub

'....................................
mError:
    
    logger "Error " & Err.Number & " - " & Err.Description
    cargarPagina mPagina, "pagInicial"

End Sub

Private Sub eraseBtn_Click()
eraseResults

rst.Requery
End Sub


'---------------------------------------------------------------------------------------
' Procedure : eraseResults
' DateTime  : 31/03/2007 17:48
' Author    : Administrador
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function eraseResults() As Boolean
    
   On Error GoTo eraseResults_Error
   '---------------------------------------------------------------------------------------------------------------

    eraseResults = True
    
    logger ""
    logger "Erasing the result grid content...."
    StatusBar1.Panels(2).Text = ""
    
    Dim res As Integer
    res = MsgBox("Do you want to discard the present results and search starts with a empty grid??" & vbNewLine _
                & "ALL THE PREVIOUS CAPTURED ENTRIES WILL BE LOST!!!", vbInformation + vbYesNo, appTitle)
    
    If res = vbNo Then
        MsgBox "The operation was cancelled", vbInformation + vbOKOnly
        eraseResults = False
        Exit Function
    End If
   
    setReasumirBusqueda False
        
    Dim rstDel As ADODB.Recordset
    Set rstDel = New ADODB.Recordset
    
    rstDel.open "delete * from data", Config("connectionstring")
    Set rstDel = Nothing
    
    LinksParsed.clear
    
    rst.Requery
    
    'actualizo la grilla de resultado
    'DataGrid1.ReBind
    redrawGrid
    
    logger ""
    logger " Grid erased."
    logger ""
    
    actualizarRuntimeInfo

  '---------------------------------------------------------------------------------------------------------------
   On Error GoTo 0
   Exit Function

eraseResults_Error:
    eraseResults = False
    
'   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure eraseResults of Formulario frmMain"
    ErrorHandler Err, "eraseResults", "Formulario", "frmMain"

End Function

'Private Sub save2XLS_Click()
'Dim result As Boolean
'Dim res As Integer
'
'    result = rst.save2XLS_HTML()
'
'    If result Then
'
'        res = MsgBox("Do you want to discard the present result so that a new search starts with a empty grid??", vbInformation + vbYesNo, appTitle)
'        If res = vbYes Then eraseBtn_Click
'
'    Else
'
'        MsgBox "The operation was canceled", vbInformation, appTitle
'
'    End If
'End Sub
'
'Private Sub save2XLS_2_Click()
'Dim result As Boolean
'Dim res As Integer
'
'    result = rst.save2XLSAutomation()
'
'    If result Then
'
'        res = MsgBox("Do you want to discard the present result so that a new search starts with a empty grid??", vbInformation + vbYesNo, appTitle)
'        If res = vbYes Then eraseBtn_Click
'
'    Else
'
'        MsgBox "The operation was canceled", vbInformation, appTitle
'
'    End If
'
'End Sub

Private Sub searchtxt_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
        searchBtn_Click
    End If
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, url As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, headers As Variant, Cancel As Boolean)
    
    If pDisp Is WebBrowser1 Then
        
        If InStr(1, url, dominio) <> 0 Then
            logger "Navigating to '" & url & "'....."
            ultimoUrl = url
            Exit Sub
        Else
            logger "Blocking url '" & url & "'....."
            Cancel = True
            Exit Sub
        End If
        
    End If
    
   ' logger "Blocking undesired url '" & url & "'....."
    'Cancel = True

    
End Sub

'Private Sub searchtxt_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'   If KeyCode = 13 Then
'        searchBtn_Click
'    End If
'End Sub

'.........................................................................................
'-----------------------------------------------------------------------------------------
'                            W  E  B        B  R  O  W  S  E  R
'-----------------------------------------------------------------------------------------
'.........................................................................................

Private Sub WebBrowser1_NavigateError(ByVal pDisp As Object, url As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)

Cancel = True

Dim msg As String
    
    Select Case StatusCode
    
        Case -2146697207
            msg = "INET_E_AUTHENTICATION_REQUIRED (0x800C0009L) - Authentication is needed to access the object."

        Case -2146697212
            msg = "INET_E_CANNOT_CONNECT (0x800C0004L) - The attempt to connect to the Internet has failed."

         Case -2146697200
            msg = "INET_E_CANNOT_INSTANTIATE_OBJECT (0x800C0010L) - CoCreateInstance failed."

        Case -2146697201
            msg = "INET_E_CANNOT_LOAD_DATA (0x800C000FL) - The object could not be loaded."

        Case -2146697194
            msg = "INET_E_CANNOT_LOCK_REQUEST (0x800C0016L) - The requested resource could not be locked."
    
        Case -2146696448
            msg = "'INET_E_CANNOT_REPLACE_SFP_FILE (0x800C0300L) - Cannot replace a file that is protected by System File Protection (SFP)."

        Case -2146696960
            msg = "INET_E_CODE_DOWNLOAD_DECLINED (0x800C0100L) - The component download was declined by the user."
    
        Case -2146695936
            msg = "INET_E_CODE_INSTALL_BLOCKED_BY_HASH_POLICY (0x800C0500L) - Internet Explorer 6 for Windows XP SP2 and later. Installation of ActiveX control (as identified by cryptographic file hash) has been disallowed by registry key policy."

        Case -2146696192
            msg = "INET_E_CODE_INSTALL_SUPPRESSED (0x800C0400L) - Microsoft Internet Explorer 6 for Microsoft Windows XP Service Pack 2 (SP2) and later. The Microsoft Authenticode prompt for installing a Microsoft ActiveX control was not shown because the page restricts the installation of the ActiveX controls. The usual cause is that the Information Bar is shown instead of the Authenticode prompt."

        Case -2146697205
            msg = "INET_E_CONNECTION_TIMEOUT (0x800C000BL) - The Internet connection has timed out."

        Case -2146697209
            msg = "INET_E_DATA_NOT_AVAILABLE (0x800C0007L) - An Internet connection was established, but the data cannot be retrieved."
    
        Case -2146697208
            msg = "INET_E_DOWNLOAD_FAILURE (0x800C0008L) - The download has failed (the connection was interrupted)."

'"INET_E_DEFAULT_ACTION (0x800C0011L) - Use the default security manager for this action. A custom security manager should only process input that is both valid and specific to itself and return INET_E_DEFAULT_ACTION for all other methods or URL actions."
'"INET_E_QUERYOPTION_UNKNOWN (0x800C0013L) - The requested option is unknown. (See IInternetProtocolInfo::QueryInfo.)"

'
        Case -2146697191
            msg = "INET_E_DOWNLOAD_FAILURE (0x800C0008L) - The download has failed (the connection was interrupted)."

        Case -2146697204
            msg = "INET_E_INVALID_REQUEST (0x800C000CL) - The request was invalid."

        Case -2146697214
            msg = "INET_E_INVALID_URL (0x800C0002L) - The URL could not be parsed."

        Case -2146697213
            msg = "INET_E_NO_SESSION (0x800C0003L) - No Internet session was established."

        Case -2146697206
            msg = "INET_E_NO_VALID_MEDIA (0x800C000AL) - The object is not in one of the acceptable MIME types."

        Case -2146697210
            msg = "INET_E_OBJECT_NOT_FOUND (0x800C0006L) - The object was not found."
    
        Case -2146697196
            msg = "INET_E_REDIRECT_FAILED (0x800C0014L) - Microsoft Win32 Internet (WinInet) cannot redirect. This error code might also be returned by a custom protocol handler."

        Case -2146697195
            msg = "INET_E_REDIRECT_TO_DIR (0x800C0015L) - The request is being redirected to a directory."

        Case -2146697196
            msg = "INET_E_REDIRECTING (0x800C0014L) - The request is being redirected. (Pass this value to IInternetProtocolSink::ReportResult.)"

        Case -2146697211
            msg = "INET_E_RESOURCE_NOT_FOUND (0x800C0005L) - The server or proxy was not found."

        Case -2146696704
            msg = "INET_E_RESULT_DISPATCHED (0x800C0200L) - The binding has already been completed and the result has been dispatched, so your abort call has been canceled."

        Case -2146696704
            msg = "INET_E_SECURITY_PROBLEM (0x800C000EL) - A security problem was encountered"

        Case -2146697192
            msg = "INET_E_TERMINATED_BIND (0x800C0018L) - Binding was terminated. (See IBinding::GetBindResult.)"

        Case -2146697203
            msg = "INET_E_UNKNOWN_PROTOCOL (0x800C000DL) - The protocol is not known and no pluggable protocols have been entered that match."

        Case -2146697193
            msg = "INET_E_USE_EXTEND_BINDING (0x800C0017L) - (Microsoft internal.) Reissue request with extended binding."
 
        Case Else
            msg = "Unespecified error message."

'INET_E_USE_DEFAULT_PROTOCOLHANDLER (0x800C0011L)
'    Use the default protocol handler. (See IInternetProtocolRoot::Start.)
'
'INET_E_USE_DEFAULT_SETTING (0x800C0012L)
'    Use the default settings. (See IInternetBindInfo::GetBindString.)
'
   
    End Select

    logger "========================================================================================="
    logger "Error (" & StatusCode & ") - " & url
    logger msg
    logger "========================================================================================="
    logger "Retrying navegate page '" & url & "'"
    
    Timer1.enabled = False
    WebBrowser1.stop
    
    ' delay en segundos
    pausarNavegacion Config("retryAfterError"), 2
    
    Dim miPag As String
    miPag = url
    cargarPagina miPag, urlActualID
    
End Sub

Private Sub pausarNavegacion(pausa As Integer, mStep As Integer)
Dim cont As Integer
    
    cont = 0
    Do
        DoEvents
        Sleep 1000 * mStep
        cont = cont + mStep
    Loop While cont <= pausa

End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
   On Error GoTo WebBrowser1_ProgressChange_Error

On Error Resume Next
    
    If Progress = -1 Then vbalProgressBar.Value = 100  'the name of the progress bar is "ProgressBar1".
        vbalProgressBar.Text = "Done"
    
    If Progress > 0 And ProgressMax > 0 Then
        'ProgressBar1.Visible = True
        'Image1.Visible = False
        If ((Progress / 100) * 99 / (ProgressMax / 100)) > 100 Then
            vbalProgressBar.Value = 100
        Else
            vbalProgressBar.Value = Progress * 100 / ProgressMax
        End If
        
        vbalProgressBar.Text = "Loading " & vbalProgressBar.Value & "%..."
    End If
    Exit Sub

   On Error GoTo 0
   Exit Sub

WebBrowser1_ProgressChange_Error:

    logger "Error " & Err.Number & " (" & Err.Description & ") en el procedimiento WebBrowser1_ProgressChange del Formulario frmMain"
End Sub

Private Sub aplicacionLista()
    
    searchBtn.enabled = True
    abortBtn.enabled = False
            
    XTabModo.TabEnabled(1) = True
 '   XTab1.TabEnabled(2) = False
 
 If LinksParsed.lastAnalizedLink <> "" And LinksParsed.persistent = True Then
        
        setReasumirBusqueda True

        'FramePlusMenuBk.enabled = True
        btnManualProcess.enabled = True
        abortBtn.enabled = False
        searchBtn.enabled = True
        LinksParsed.enabled = False
        Timer1.enabled = False
    
        StatusBar1.Panels(2).Text = "Records [" & Val(rst.RecordCount) & "]"
        
        'cargarPagina LinksParsed.lastAnalizedLink, "Analisis", Config("loadTimer")
        urlActualID = "Analisis"
        ultimoUrl = LinksParsed.lastAnalizedLink
        
        logger ""
        logger " Last url restored " & ultimoUrl
        logger " Last category restored " & LinksParsed.categoriaActual
        logger ""
        
        logger "---------------------------------------------------------------------------------"
        logger ""
        logger " All information about previous session was restored and it is ready to continue."
        logger ""
        logger "---------------------------------------------------------------------------------"
        logger ""
        
    End If
    
    logger ""
    logger ""
    logger " Aplication ready."
    logger ""
        
End Sub



Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, url As Variant)

If (pDisp Is WebBrowser1.Application) Then
    
    'pagina completa, deshabilito el timer de control
    Timer1.enabled = False
    
    logger "Load process completed for '" & url & "' - [" & urlActualID & "]"
    
    'estaba analizando y mande abortar toda operacion?
    If (abortBtn.enabled = False And urlActualID <> "pagInicial") Then
            
            logger "Pending abort request detected."
            Exit Sub
    End If
  '  Or (abortBtn.enabled = False And urlActualID <> "Categoria") Then Exit Sub
    
    Select Case urlActualID
        
        Case "pagInicial"
                
               
                ' Este evento generalmente se usa para obtener los cookies de session que seran utilizados
                ' por la aplicacion durante la ejecucion. Otras tareas posibles son establecer algunos controles
                ' y dejarlos listos para empezar a ejecutar.
                    
                StatusBar1.Panels(1).Text = "Application ready."
'                aplicacionLista
                
                analisisPaginaInicial WebBrowser1.document
        
        Case "Busqueda"
  
                logger "Starting Analisis for '" & searchtxt & "'...."
                StatusBar1.Panels(1).Text = "Searching process....waiting........"
    
                analisisPrimerResultado WebBrowser1.document
         
        Case "Analisis"
                
                logger "Analizing result for '" & url & "'..."
               
                analizarPagina WebBrowser1.document
    
        Case "Business"
                
                logger "Analizing busuness result for '" & url & "'..."
                StatusBar1.Panels(1).Text = "Analizing....."
                
         '       analisisBusiness WebBrowser1.document
                        
        Case "Manual"
                
                logger "Analizing result page  '" & url & "'..."
                StatusBar1.Panels(1).Text = "Analizing....."
                
                analisisResult WebBrowser1.document
                        
        Case "Categoria"
                
                logger "Analizing result page  '" & url & "'..."
                StatusBar1.Panels(1).Text = "Analizing....."
                
                analisisResult2 WebBrowser1.document
        
        Case Else
        
                StatusBar1.Panels(1).Text = "Done."
    
        End Select
    Else
        logger "(FRAME)------ Completed loading process for '" & url & "'", 3
        
        Select Case urlActualID
                        
            Case "Manual"
                
                logger "Analizing result page  '" & url & "'..."
                StatusBar1.Panels(1).Text = "Analizing....."
                
                analisisResult WebBrowser1.document
                        
            Case Else
    
        End Select
    End If

End Sub

Private Sub Timer1_Timer()

'READYSTATE_UNINITIALIZED = 0
'' Default initialization state.
'READYSTATE_LOADING = 1
'' Object is currently loading its properties.
'READYSTATE_LOADED = 2
'' Object has been initialized.
'READYSTATE_INTERACTIVE = 3
'' Object is interactive, but not all of its data is available.
'READYSTATE_COMPLETE = 4
'' Object has received all of its data.
Dim estado As String

With WebBrowser1
    
    If Config("debugLevel") > 1 Then
        
        Select Case .readyState
            Case 0
                estado = "READYSTATE_UNINITIALIZED"
            Case 1
                estado = "READYSTATE_LOADING"
            Case 2
                estado = "READYSTATE_LOADED"
            Case 3
                estado = "READYSTATE_INTERACTIVE"
            Case 4
                estado = "READYSTATE_COMPLETE"
        End Select
        
        StatusBar1.Panels(3).Visible = True
        StatusBar1.Panels(3).Text = "State: " & estado

    End If
    
    If .readyState <> READYSTATE_COMPLETE Then

        If DateDiff("s", navTimeStamp, Now) > CInt(Config("connectionTimeout")) Then

            logger "========================================================================================="
            logger "--- TIMEOUT WAITING RESPONSE ----- Retrying navegate page '" & .LocationURL & "'"
            logger "========================================================================================="

            .stop

            cargarPagina .LocationURL, urlActualID

        End If

    End If

End With

End Sub

'---------------------------------------------------------------------------------------
' Procedure : navegarPagina
' DateTime  : 08/11/2006 12:19
' Author    : Administrador
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub navegarPagina(pag As String)
Dim flag As Long
Dim sHeaders

'flag = 0 'navNoReadFromCache  ' vbEmpty
'sHeaders = "Cookie: " & "CP=null*; " & langCookie & vbCrLf  'culturenameCookie=es-CO ' Add extra headers as needed
'WebBrowser1.Navigate2 pag, , , , sHeaders


    'WebBrowser1.Navigate2 pag
   On Error GoTo navegarPagina_Error

    WebBrowser1.navigate pag, navNoHistory & navNoReadFromCache & navNoWriteToCache
    Timer1.enabled = True
    
    logger "Opening '" & pag & "'.............."

'Do Until Not WebBrowser1.Busy
'    DoEvents
'Loop
'
'Do While WebBrowser1.document.body Is Nothing
'    DoEvents
'Loop

   On Error GoTo 0
   Exit Sub

navegarPagina_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en el procedimiento navegarPagina del Formulario frmMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cargarPagina
' DateTime  : 26/02/2007 23:49
' Author    : Administrador
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub cargarPagina(url As String, urlID As String, Optional loadDelay As Long = 0)

    urlActualID = urlID
    ultimoUrl = url
    
    If loadDelay = 0 Then
        cargarPaginaPostTimer url, urlID
    Else
        loadTimerSeg = loadDelay
        logger "Pausing the load process for " & loadDelay & " secs.....", 2
        LoadTimer.Tag = loadDelay
        LoadTimer.enabled = True
    End If

'    ' almaceno el timestamp de request
'    navTimeStamp = Now
'
'    urlActualID = urlId
'    ultimoUrl = url
'
'    navegarPagina url
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cargarPaginaPostTimer
' DateTime  : 26/02/2007 23:49
' Author    : Administrador
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub cargarPaginaPostTimer(url As String, urlID As String)

    ' almaceno el timestamp de request
    navTimeStamp = Now

'    urlActualID = urlId
'    ultimoUrl = url
    
    navegarPagina url
End Sub

'---------------------------------------------------------------------------------------
' Procedure : loadTimer_Timer
' DateTime  : 27/02/2007 00:01
' Author    : Administrador
' Purpose   : Timer dedicado a poner un demora entre cada carga de pagina cuando es necesario
'---------------------------------------------------------------------------------------
'
Private Sub loadTimer_Timer()
Dim actual As Long

   On Error GoTo loadTimer_Timer_Error
   '---------------------------------------------------------------------------------------------------------------

With LoadTimer
    'actual = (.Tag + 1) * (.Interval / 1000)
    'If actual >= loadTimerSeg Then
     
    .Tag = .Tag - 1
    
     If .Tag <= 0 Then
        Debug.Print "Timeout. Firing process..."
        .enabled = False 'detengo el timer, es de un solo disparo
        
        logger "Load process restarted. ", 2
        
        'cargo la pagina despues de agotado el timer de espera
        cargarPaginaPostTimer ultimoUrl, urlActualID
        
        .Tag = 0
        Exit Sub
    End If
    
    '.Tag = actual

End With

  '---------------------------------------------------------------------------------------------------------------
   On Error GoTo 0
   Exit Sub

loadTimer_Timer_Error:

   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadTimer_Timer of Formulario frmMain"
'    ErrorHandler Err, "loadTimer_Timer", "Formulario", "frmMain"

    
End Sub


Sub buscar(sUrl As String, sPost As String)
Dim bPostData() As Byte
Dim sHeaders As String
Dim res As Integer

    urlActualID = "Busqueda"  ' set del id de resultado
    
    'sPost = "foo=1&bar=test&quux=42"  ' Post data
    sHeaders = "Content-Type: application/x-www-form-urlencoded" & vbCrLf '' Add extra headers as needed
    ReDim bPostData(Len(sPost))
    bPostData = StrConv(sPost, vbFromUnicode)
   
    WebBrowser1.navigate sUrl, 0, vbEmpty, bPostData, sHeaders

    Do Until Not WebBrowser1.Busy
    DoEvents
    Loop

    Do While WebBrowser1.document.body Is Nothing
    DoEvents
    Loop

End Sub

'====================================================================================================
' Arma los parametros del POST de busqueda con formateos de urlencode, urlsafe y similares
'----------------------------------------------------------------------------------------------------
'
' POST /search.ds newSearch=true&locale=nl_NL&what=Reclame&where=&x=17&y=22
'
'====================================================================================================
Function armarPostDeBusqueda(searchtxtVal As String) As String
Dim postDetail As String
Dim str1 As String

str1 = URLEncode(searchtxtVal)

postDetail = "newSearch=true&locale=nl_NL&what=" & str1 & "&where=&x=17&y=22"

logger searchtxt & " -> " & "POST: " & postDetail

armarPostDeBusqueda = postDetail

'    '-------------Write to Text-----------------
'    Open App.Path & "\testDetail.txt" For Output As #2     '// open the text file
'    Write #2, postDetail
'    Close 2 '// close the text file
'    '-------------Write to Text-----------------
End Function

'.........................................................................................
'-----------------------------------------------------------------------------------------
'                            W  E  B        B  R  O  W  S  E  R
'-----------------------------------------------------------------------------------------
'.........................................................................................








'.........................................................................................
'-----------------------------------------------------------------------------------------
'       P R O C E S O    D E    P A G I N A S    D E   R E S U L T A D O S
'-----------------------------------------------------------------------------------------
'.........................................................................................

Sub analisisPaginaInicial(mDoc As MSHTML.HTMLDocument)
Dim elem As IHTMLElement
            
   On Error GoTo analisisPaginaInicial_Error
   '---------------------------------------------------------------------------------------------------------------
    
    If Config("DebugLevel") > 1 Then
                    
        Open App.Path & "\initPage.html" For Output As #3
        Write #3, "<HTML>" & mDoc.body.outerHTML & "</HTML>"
        Close 3
                
    End If
        
If Not mDoc Is Nothing Then
       
    ' parser de la pagina para identificar los componentes buscados
    'actionPage = mDoc.Forms(1).Action
    'mDoc.Forms(1).All("_dyncharset").Value
    Debug.Print mDoc.cookie

     '============= MAIN PAGE DETECTION ROUTINE ==============================
            
    Dim mainObj As Object
    Set mainObj = mDoc.getElementById("search")
    
    'If mDoc.Forms.Item("ypform") Is Nothing Then
     If Not checkObjeto(mainObj, "HTMLDivElement", "homePageObj") Then
        
            logger "Main page missing or invalid."
        
            logger ""
            logger "#####################################################################################################"
            logger "The loaded home page page is invalid."
            logger "Please verify if the site is not down at this moment."
            logger "#####################################################################################################"
            logger ""
            
            cargarPagina mPagina, "pagInicial", 30
            Exit Sub
        
    End If
        
    
  '  Dim mainObj As Object
  '  Set mainObj = mDoc.Forms.ypform
    
  '  If Not checkObjeto(mainObj, "HTMLFormElement", "Main page FormObject missing") Then
  '      Exit Sub
   ' Else
        logger "Main page detected."
        
        With lblMensaje
            .BackColor = &H0&
            .caption = "Ready"
        End With
        
        lblMensajeSup.caption = "Application ready to start the capture"
        
    'End If
    
    
    '============= MAIN PAGE DETECTION ROUTINE ==============================
    
    aplicacionLista
    
    While CookiesCls.Count > 0
        CookiesCls.Remove 0
    Wend
    
    getParams mDoc.cookie, CookiesCls, ";"

    'armarTypeLinkList

'
'    Dim recordGroupObj  As Object
'    Set recordGroupObj = mDoc.getElementById("ypform")
'    If Not checkObjeto(recordGroupObj, "HTMLFormElement", "listings_RecordGroupObj") Then
'
'            lblMensaje.Visible = False
'            btnManualProcess.enabled = False
'            Exit Sub
'    End If
'
'    lblMensaje.Visible = True
'    btnManualProcess.enabled = True


Else
    logger "(initial page) Error processing '" & mDoc.url & "'...."
End If

   On Error GoTo 0
   Exit Sub

   '---------------------------------------------------------------------------------------------------------------
   
analisisPaginaInicial_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") en el procedimiento analisisPaginaInicial del Formulario frmMain"
    ErrorHandler Err, "analisisPaginaInicial", "Formulario", "frmMain"

End Sub

Private Sub selectStates()
Dim a As frmSelect
Dim rs2 As ADODB.Recordset
  
  Set rs2 = New ADODB.Recordset
  With rs2
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .open "select *, texto + ' (' + state +')' as msg from statesUSA", Config("connectionstring")


        Set a = New frmSelect
        a.setData rs2, "msg", "parsed"
        a.Show vbModal
  End With
  
  rsState.Requery
End Sub



'---------------------------------------------------------------------------------------
' Procedure : armarTypeLinkList
' DateTime  : 29/03/2007 15:18
' Author    : Administrador
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub armarTypeLinkList()
Dim cookie As String

    '---get cookies----
    
   ' cookie = Trim(CookiesCls("JSESSIONID"))
    
    Dim typeRst As ADODB.Recordset
    Set typeRst = New ADODB.Recordset
    
    Dim estado, mLink As String
    
    With rsState
        .Requery
    
        If .EOF Then
'            logger ""
'            logger " No hay mas estados sin procesar."
'            logger ""
            ' fuezo esto porque alli tengo codigo para manejar el fin de la busqueda
            siguienteLinkEnEspera
            
            Exit Sub
        End If
        
        estado = rsState.Fields("state").Value
        logger "  >======================================================="
        logger "  Processing state [" & estado & "]"
        logger "  >-------------------------------------------------------"
        logger "  "
    
    End With
    
    LinksParsed.clear
    
    With typeRst
        
        .open "select * from types", Config("connectionString")
        .MoveFirst
    
        While Not .EOF
            
            'mLink = dominio & "/sp/yellowpages/ypresults.jsp;jsessionid=" & cookie & "?t=0&v=3&s=2&q=" & EncodeUrl(.Fields("type").Value) & "&st=" & estado
            'mLink = dominio & "/" & estado & "/" & YellowEncodeUrl(.Fields("type").Value) '"sp/yellowpages/ypresults.jsp;jsessionid=" & cookie & "?t=0&v=3&s=2&q=" & EncodeUrl(.Fields("type").Value) & "&st=" & estado
            mLink = dominio & "/search?search_terms=" & Replace(YellowEncodeUrl(.Fields("type").Value), " ", "+") & "&geo_location_terms=" & estado
            'mLink = "http://www.yellowpages.com/search?search_terms=limousine+service&geo_location_terms=WY"
            LinksParsed.Add mLink, .Fields("type").Value, -1
            .MoveNext
        Wend

    End With
    
    Dim mlnk As New LinkClass
    If Not LinksParsed.getFirstLink(mlnk) Then
        logger "Error recovering first category....."
        Exit Sub
    End If
    
    lblMensajeSup.caption = "[" & rsState.Fields("state").Value & "] -" & LinksParsed.categoriaActual

    'empezar el proceso
    cargarPagina mlnk.link, "Categoria"
    
End Sub










'========================================================================
' Operaciones sobre la primer pagina recuperada. Gte se usa para obtener
' valor que se utilizaran para analizar las pag asociadas a esta.
'========================================================================

Sub analisisPrimerResultado(mDoc As MSHTML.HTMLDocument)
Dim elem As IHTMLElement
            
   On Error GoTo analisisPrimerResultado_Error
   '---------------------------------------------------------------------------------------------------------------
    
If Config("DebugLevel") > 1 Then
            
    '-------------Write to Text-----------------
    Open App.Path & "\paginaResultadoInicial.html" For Output As #3 'Append As #3     '// open the text file
    Write #3, "<HTML>" & mDoc.body.outerHTML & "</HTML>"
    Close 3 '// close the text file
    '-------------Write to Text-----------------
End If

If Not mDoc Is Nothing Then
       
    ' parser de la pagina para identificar los componentes buscados
    analizarPagina mDoc

Else
    logger "Error processing '" & mDoc.url & "'...."
End If

   On Error GoTo 0
   Exit Sub

   '---------------------------------------------------------------------------------------------------------------
   
analisisPrimerResultado_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") en el procedimiento analisisPrimerResultado del Formulario frmMain"
    ErrorHandler Err, "analisisPrimerResultado", "Formulario", "frmMain"

End Sub

' groninger
Sub siguienteLinkEnEspera()
   ' tengo link por procesar aun???
   ' Dim lnk As LinkClass
   On Error GoTo siguienteLinkEnEspera_Error
   '---------------------------------------------------------------------------------------------------------------
    
    Dim lnk As LinkClass
    Set lnk = New LinkClass
        
    LinksParsed.checkItem LinksParsed.categoriaActual
        
    If LinksParsed.getFirstLink(lnk) Then
            
            If Not lnk Is Nothing Then
'                cargarPagina mPagina & lnk.link, "Analisis"

                lblMensajeSup.caption = "[" & rsState.Fields("state").Value & "] -" & LinksParsed.categoriaActual

                logger ""
                logger "Searching for '" & lnk.texto & "'...."
                cargarPagina lnk.link, "Analisis"
            Else
                 logger "Searching process completed."
                    abortBtn.enabled = False
                    searchBtn.enabled = True
                    btnManualProcess.enabled = True
                    
                    setReasumirBusqueda False 'termiando normalmente no soporta resume
            End If
    Else
        
        
        '000000000000000000000000000000000000000000000000000000000000000000000000000000000000
        '000000000000000000000000000000000000000000000000000000000000000000000000000000000000
        'capture todos los link de este estado, me muevo al siguiente
        With rsState
        
            If .EOF Then
                
                logger ""
                logger "No more state to be processed."
                logger ""
            
                '----------------------------------------------------------------------------------
                StatusBar1.Panels(1).Text = "Ready." 'Records [" & Val(rsRes.RS.RecordCount) & "]"
        
                ' LinksParsed.dump
                abortBtn.enabled = False
                searchBtn.enabled = True
                btnManualProcess.enabled = True
                setReasumirBusqueda False 'termiando normalmente no soporta resume
        
        
                lblMensaje.caption = "Capture complete"
                lblMensaje.ForeColor = RGB(0, 0, 255)
                lblMensajeSup.caption = ""
                '----------------------------------------------------------------------------------
            
                Exit Sub
            End If
            
            .Fields("parsed").Value = False 'True
            .Update
            
            logger ""
            logger "Process completed for the present state [" & rsState("state").Value & "]."
            logger ""
        
            .MoveNext
            
            armarTypeLinkList
            
            Exit Sub
        
        End With
        '000000000000000000000000000000000000000000000000000000000000000000000000000000000000
        '000000000000000000000000000000000000000000000000000000000000000000000000000000000000
    
     
        'urlActualID = "Manual"
      '  analisisResult WebBrowser1.document
    End If

   On Error GoTo 0
   Exit Sub

   '---------------------------------------------------------------------------------------------------------------
   
siguienteLinkEnEspera_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") en el procedimiento siguienteLinkEnEspera del Formulario frmMain"
    ErrorHandler Err, "siguienteLinkEnEspera", "Formulario", "frmMain"

End Sub


Sub analisisResult(ByRef mDoc As MSHTML.HTMLDocument)

   On Error GoTo analisisResult_Error
   '---------------------------------------------------------------------------------------------------------------
    
If Config("DebugLevel") > 1 Then
    Open App.Path & "\pageInAnalysisManual.html" For Output As #4
    Write #4, "<HTML>" & mDoc.body.outerHTML & "</HTML>"
    Close 4 '// close the text file
End If

    Dim mObj As Object

    If Not mDoc Is Nothing Then

    Set mObj = mDoc.All.Item("PrintSearch", 0)

    If checkObjeto(mObj, "HTMLSpanElement", "detail_detectionObjs", False) Then
        
        btnManualProcess.enabled = True
                  
'        'lblMensaje.Visible = True
'        lblMensaje.caption = "Detail page detected!"
'        lblMensaje.ForeColor = RGB(255, 0, 0)
'        lblMensajeSup.caption = "Page detected and ready for analysis"
'
         
         logger ""
         logger "Detail page detected and ready for analysis."
         logger ""
         
         
        Exit Sub
    Else
        
'        'lblMensaje.Visible = True
'        lblMensaje.caption = "Free browsing..."
'        lblMensaje.ForeColor = RGB(0, 0, 255)
'        lblMensajeSup.caption = ""
'
        
        'lblMensaje.Visible = False
        btnManualProcess.enabled = False
    
    End If

Else
    logger "Error processing '" & mDoc.url & "'...."

End If


   On Error GoTo 0
   Exit Sub

   '---------------------------------------------------------------------------------------------------------------
   
analisisResult_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") en el procedimiento analisisResult del Formulario frmMain"
    ErrorHandler Err, "analisisResult", "Formulario", "frmMain"

End Sub

Sub analisisResult2(ByRef mDoc As MSHTML.HTMLDocument)

   On Error GoTo analisisResult_Error
   '---------------------------------------------------------------------------------------------------------------
    
If Config("DebugLevel") > 1 Then
    Open App.Path & "\pageInAnalysisManual.html" For Output As #4
    Write #4, "<HTML>" & mDoc.body.outerHTML & "</HTML>"
    Close 4 '// close the text file
End If

    Dim mObj As Object

    If Not mDoc Is Nothing Then

    Set mObj = mDoc.All.Item("PrintSearch", 0)

    If checkObjeto(mObj, "HTMLSpanElement", "detail_detectionObjs", False) Then
        
        btnManualProcess.enabled = True
                  
'        'lblMensaje.Visible = True
'        lblMensaje.caption = "Detail page detected!"
'        lblMensaje.ForeColor = RGB(255, 0, 0)
'        lblMensajeSup.caption = "Page detected and ready for analysis"
'
'
'         logger ""
'         logger "Detail page detected and ready for analysis."
'         logger ""
'
         btnManualProcess_Click
        
        Exit Sub
    Else
        
'        'lblMensaje.Visible = True
'        lblMensaje.caption = "Free browsing..."
'        lblMensaje.ForeColor = RGB(0, 0, 255)
'        lblMensajeSup.caption = ""
'
        
        'lblMensaje.Visible = False
        btnManualProcess.enabled = False
    
        btnManualProcess_Click
        
    End If


Else
    logger "Error processing '" & mDoc.url & "'...."

End If


   On Error GoTo 0
   Exit Sub

   '---------------------------------------------------------------------------------------------------------------
   
analisisResult_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") en el procedimiento analisisResult del Formulario frmMain"
    ErrorHandler Err, "analisisResult", "Formulario", "frmMain"

End Sub



'----------------------------------------------------------------------------------------
' Analisis general del contenido de una pagina
' Aqui generalmente se separan y muestran encabezados, total de pagina
' y se envia el grupo de registro a la funcion que los procesara.
' Si esta ultima funcion indica que no hay pagina siguiente, la presente
' funcion llamara al siguiente link en la lista para continuar con el
' 'crawler' de los sig links.
' ( NO SE USO ESTO ULTIMO EN ESTE PROYECTO)
'----------------------------------------------------------------------------------------

Sub analizarPagina(ByRef mDoc As MSHTML.HTMLDocument)

'Dim elems As IHTMLElementCollection
Dim htmlItem As Object
Dim res As Boolean
Dim i As Integer
Dim actual As String

   On Error GoTo analizarPagina_Error
   '---------------------------------------------------------------------------------------------------------------
    
If Config("DebugLevel") >= 2 Then

    '-------------Write to Text-----------------
    Open App.Path & "\pageInAnalysis.html" For Output As #4
    Write #4, "<HTML>" & mDoc.body.outerHTML & "</HTML>"
    Close 4 '// close the text file
    '-------------Write to Text-----------------
End If


' busco un objeto de referencia para saber si la pagina recuperada es una pagina de resultado
' valida
Dim mainHtlmObj As Object

Set mainHtlmObj = getElementByClass(mDoc.All, "listings")

If mainHtlmObj Is Nothing Then
    logger "Error: 'listings' class not found."
    GoTo seguir
End If

'=========== CONTADOR DE PAGINA ====================='

Dim htmlCont As Object
Dim txtObj As Object

Set htmlCont = mDoc.getElementById("toolbar-btm")
If checkObjeto(htmlCont, "HTMLDivElement", "detail_counterDetect") Then

    Set txtObj = htmlCont.getElementsByTagName("P")

    resultsTxt = txtObj.Item(txtObj.length - 1).innerText
    
    logger ""
    logger " ----- [" & resultsTxt & "] ------"
    logger ""

    actualizarRuntimeInfo

End If

'=========== CONTADOR DE PAGINA ====================='


'------- proceso en grupo de registros ----------
Dim mObj As Object

Set mObj = mDoc.getElementById("mid-column")
If Not checkObjeto(mObj, "HTMLDivElement", "detail_MainObj", False) Then GoTo seguir



Dim recordGroupObj As Object
Dim Item As Object
Set recordGroupObj = getElementByClass(mObj.All, "listings")
    
If Not checkObjeto(recordGroupObj, "HTMLUListElement", "detail_detectionParsing", False) Then GoTo seguir
procesarRegistros recordGroupObj
      
    
'------- proceso en grupo de registros ----------

seguir:


Config("lastanalizedlink") = mDoc.url

'actualizo la grilla de resultado despues de procesar el lote de registros
DataGrid1.ReBind

'StatusBar1.Panels(2).Text = "Records [" & Val(rsRes.RS.RecordCount) & "]"
 StatusBar1.Panels(2).Text = "Records [" & Val(rst.RecordCount) & "]"

'btnManualProcess.enabled = True
'Exit Sub

Dim sigPagina As String

sigPagina = detectarSiguientePagina(mDoc)

If sigPagina = "" Then
    logger "No more pages detected."
    siguienteLinkEnEspera
Else
    logger "New page detected '" & sigPagina & "'."
    cargarPagina sigPagina, "Busqueda"
End If

   On Error GoTo 0
   Exit Sub

   '---------------------------------------------------------------------------------------------------------------
   
analizarPagina_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") en el procedimiento analizarPagina del Formulario frmMain"
    ErrorHandler Err, "analizarPagina", "Formulario", "frmMain"

End Sub



'----------------------------------------------------------------------------------------
' separo cada registro individual y mando su procesamiento
'----------------------------------------------------------------------------------------

Sub procesarRegistros(recGroupObj As Object)
Dim records As Object
Dim rec As Object

   On Error GoTo procesarRegistros_Error
   '---------------------------------------------------------------------------------------------------------------
Dim contador As Integer

'------ Objecto "listing advertiser first"---------
contador = 0
Do
    
    Set records = getElementByClass(recGroupObj.All, "listing advertiser first", contador)
    logger "'listing advertiser first' object(" & contador & ")", 2
    If records Is Nothing Then Exit Do
    
    procesarRegistro records
    
    contador = contador + 1

Loop While contador < 100

'------ Objecto "listing advertiser"---------
contador = 0
Do
    
    Set records = getElementByClass(recGroupObj.All, "listing advertiser", contador)
    logger "'listing advertiser' object(" & contador & ")", 2
    If records Is Nothing Then Exit Do
    
    procesarRegistro records
    
    contador = contador + 1

Loop While contador < 100


'------ Objecto "listing"---------
contador = 0
Do
    
    Set records = getElementByClass(recGroupObj.All, "listing", contador)
    logger "'Listing' object(" & contador & ")", 2
    If records Is Nothing Then Exit Do
    
    procesarRegistro records
    
    contador = contador + 1

Loop While contador < 100
 
Exit Sub

'Set records = recGroupObj.getElementsByTagName("div") ' getelementsbyclass(recGroupObj.All, "listing")
'
'For Each rec In records
'
'    If rec.className = "description" Then procesarRegistro rec
'
'Next

   On Error GoTo 0
   Exit Sub

   '---------------------------------------------------------------------------------------------------------------
   
procesarRegistros_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") en el procedimiento procesarRegistros del Formulario frmMain"
    ErrorHandler Err, "procesarRegistros", "Formulario", "frmMain"

End Sub


'----------------------------------------------------------------------------------------
' proceso un registro individual y obtengo la informacion buscada
'----------------------------------------------------------------------------------------
Sub procesarRegistro(objRec As Object)
Dim params As Collection
Dim business As String
Dim direccion As String
Dim telefono As String
Dim cpostal As String
Dim ciudad As String
Dim email As String
Dim website As String
Dim estado As String

Dim temp As String
Dim rec As Object
Dim tempObj As Object
Dim i As Integer, contador As Integer
Dim strSplit() As String
Dim nodes As Object
Dim Item As Object
        
Dim reg As New RegExp
Dim myMatches As MatchCollection
Dim myMatch As Match
Dim addressTemp As String
reg.IgnoreCase = True
              
        
        
        On Error GoTo ErrorHandler
        
        Dim objDesc  As Object
        Set objDesc = getElementByClass(objRec.All, "description")
        If Not checkObjeto(objDesc, "HTMLDivElement", "'description' object not found") Then GoTo salida
        
        '============-- business --===============
        Set Item = objDesc.getElementsByTagName("H2").Item(0)
        If Not checkObjeto(Item, "HTMLHeaderElement", "'bussiness' object not found") Then GoTo salida
        
        business = Trim(Item.innerText)
        
        ' ----- direccion -----
        
        Set Item = objDesc.getElementsByTagName("P").Item(0)
        direccion = ""
        
        '--------------------------------------------------------------
        
        '<P>3939 Lavista Rd <BR>Tucker, GA 30084 <A onmousedown="omnitu........
        reg.Pattern = "\<P>(.*)<BR>(.*),\s(\w*)\s(\d*)\s\<A"
        Set myMatches = reg.Execute(Item.outerHTML)
        
        ' Encontrado
        If myMatches.Count > 0 Then
                direccion = Trim(myMatches(0).SubMatches(0))
                ciudad = Trim(myMatches(0).SubMatches(1))
                estado = Trim(myMatches(0).SubMatches(2))
                cpostal = Trim(myMatches(0).SubMatches(3))
        End If
        
        '--------------------------------------------------------------
        
        '<P>3939 Lavista Rd <BR>Tucker, GA 30084 </P>
        '<P>4023 4029 Spring Mountain Rd <BR>Henderson, NV 89044 </P>
        reg.Pattern = "\<P>(.*)<BR>(.*),\s(\w*)\s(\d*)\s\</P"
        Set myMatches = reg.Execute(Item.outerHTML)
        
        ' Encontrado
        If myMatches.Count > 0 And direccion = "" Then
                direccion = Trim(myMatches(0).SubMatches(0))
                ciudad = Trim(myMatches(0).SubMatches(1))
                estado = Trim(myMatches(0).SubMatches(2))
                cpostal = Trim(myMatches(0).SubMatches(3))
        End If
        
        '--------------------------------------------------------------
        
        '<P>Henderson, NV 89044 </P>
        reg.Pattern = "\<P>(.*),\s(\w*)\s(\d*)\s\</P"
        Set myMatches = reg.Execute(Item.outerHTML)
        
        ' Encontrado
        If myMatches.Count > 0 And direccion = "" Then
                direccion = " "
                ciudad = Trim(myMatches(0).SubMatches(0))
                estado = Trim(myMatches(0).SubMatches(1))
                cpostal = Trim(myMatches(0).SubMatches(2))
        End If
        
        '--------------------------------------------------------------
        
        '<P>Pinedale, WY 82941 <A onmousedown="om....
        reg.Pattern = "\<P>(.*),\s(\w*)\s(\d*)\s\<A"
        Set myMatches = reg.Execute(Item.outerHTML)
        
        ' Encontrado
        If myMatches.Count > 0 And direccion = "" Then
                direccion = " "
                ciudad = Trim(myMatches(0).SubMatches(0))
                estado = Trim(myMatches(0).SubMatches(1))
                cpostal = Trim(myMatches(0).SubMatches(2))
        End If
        
        
        '<P>Serving the Atlanta Area</P>
        '--------------------------------------------------------------
        reg.Pattern = "\<P>Serving(.*)</P"
        Set myMatches = reg.Execute(Item.outerHTML)
        
        ' Encontrado
        If myMatches.Count > 0 And direccion = "" Then
                direccion = "Serving " & Trim(myMatches(0).SubMatches(0))
                ciudad = ""
                estado = ""
                cpostal = ""
        End If
        
        
        If direccion = "" Then
            logger ""
            logger "----------->>>>>>>>>> 'address' not found. <<<<<<<<<<<<<<------------"
            logger ""
        End If
        
        ' -----  telefono -----
        
        Set tempObj = getElementByClass(objDesc.All, "number")
        
        If tempObj Is Nothing Then GoTo salida
        telefono = Trim(tempObj.innerText)
           
        ' -----  email -----
        
        Set tempObj = getElementByClass(objDesc.All, "email")
        
        If checkObjeto(tempObj, "HTMLAnchorElement", "'email' object not found", False) Then
            email = Trim(Replace(tempObj.href, "mailto:", ""))
        End If
           
        
        'XXXXXXxxxxx=======------- website ------=========xxxxxxXXXXXXXX
        
        Dim objWeb  As Object
        Set objWeb = getElementByClass(objRec.All, "options")
        If Not checkObjeto(objWeb, "HTMLDivElement", "'options' object not found", False) Then GoTo registrar

        Set tempObj = getElementByClass(objWeb.All, "web")
        If checkObjeto(tempObj, "HTMLAnchorElement", "'web' object not found", False) Then
            website = tempObj.href
        End If

registrar:
        
        ciudad = Trim(Replace(ciudad, "<BR>", ""))
        
        logger "  >>" & business & "|" & direccion & "|" & ciudad & "|" & cpostal & "|" & estado & "|" & telefono & "|" & website & "|" & email
        
        'Exit Sub
    
        '---------------------------------------------------------------
        '===============================================================
        '    R E C O R D S E T    D E   R E S U L T A D O
        '===============================================================
        '---------------------------------------------------------------
        On Error GoTo AdoErrorHandler

        With rst 'rsRes.RS
            .AddNew

            .Fields("business") = business
            .Fields("address") = direccion
            .Fields("zip").Value = IIf(cpostal = "", Null, cpostal)
            .Fields("city") = ciudad
            .Fields("tel") = telefono
            .Fields("state") = IIf(estado <> "", estado, rsState.Fields("state").Value)
            .Fields("website") = website
            .Fields("email") = email
            '.Fields("keyWord") = searchtxt.Text
            .Fields("category") = LinksParsed.categoriaActual
            
            .Update

        '===============================================================
        '---------------------------------------------------------------
        End With

salida:
        
        actualizarRuntimeInfo

Exit Sub

AdoErrorHandler:
    
    logger ""
    logger "ADO Error " & Err.Number & " - " & Err.Description
    logger ""
        
        Exit Sub

ErrorHandler:
    
    logger "================================================================"
    logger "Error " & Err.Number & " - " & Err.Description
    logger ""
    logger rec.outerHTML
    logger "================================================================"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : getObjbyAttribute
' DateTime  : 13/03/2007 13:05
' Author    : Administrador
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function getObjbyAttribute(mAttrib As String, doc As Object, Optional tagName As String = "") As Object

Dim tempObj As Object
Dim tempItem As Object
Dim tmpObj1 As Object
Dim st() As String

   On Error GoTo getObjbyAttribute_Error
   '---------------------------------------------------------------------------------------------------------------
    
st = Split(mAttrib, "=")

If UBound(st) = 1 Then
    
    If tagName <> "" Then
        Set tempObj = doc.getElementsByTagName(tagName)
    Else
        Set tempObj = doc
    End If

    For Each tempItem In tempObj
        Set tmpObj1 = tempItem.Attributes
        
        If Not tmpObj1 Is Nothing Then
        
            If Not tmpObj1.getNamedItem((st(0))) Is Nothing Then
                If tmpObj1.getNamedItem((st(0))).Value = st(1) Then
                    Set getObjbyAttribute = tempItem
                    Exit Function
                End If
            End If
        
        End If
    Next

Else

    Set tempObj = doc.getElementById(mAttrib)
    If Not tempObj Is Nothing Then
        Set getObjbyAttribute = tempObj
        Exit Function
    End If
End If


Set getObjbyAttribute = Nothing

   On Error GoTo 0
   Exit Function

   '---------------------------------------------------------------------------------------------------------------
   
getObjbyAttribute_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") en el procedimiento getObjbyAttribute del Formulario frmMain"
    ErrorHandler Err, "getObjbyAttribute", "Formulario", "frmMain"

End Function

Function checkObject(Objeto As Object, tipo As String) As Boolean

        If Objeto Is Nothing Then
            logger "Error: Object '" & tipo & "' not found."
            checkObject = False
            Exit Function
        End If
        
checkObject = True
End Function



'----------------------------------------------------------------------------------------
' Analisis general del contenido de una pagina
' Aqui generalmente se separan y muestran encabezados, total de pagina
' y se envia el grupo de registro a la funcion que los procesara.
' Si esta ultima funcion indica que no hay pagina siguiente, la presente
' funcion llamara al siguiente link en la lista para continuar con el
' 'crawler' de los sig links.
' ( NO SE USO ESTO ULTIMO EN ESTE PROYECTO)
'----------------------------------------------------------------------------------------

Function checkObjeto(miObj As Object, tipo As String, msg As String, Optional loguear As Boolean = True) As Boolean
        
        If miObj Is Nothing Or TypeName(miObj) <> tipo Then
            If loguear Then logger " --Object '" & msg & "' [" & TypeName(miObj) & "] not found."
            
            checkObjeto = False
            Exit Function
        End If
        
        checkObjeto = True

End Function

Function controlHTMLNullItem(ByRef miObj As Object) As String
    
    If miObj Is Nothing Then
        controlHTMLNullItem = ""
    Else
        controlHTMLNullItem = miObj.innerText
    End If
End Function

Sub mostrarCantidadRegs()

    StatusBar1.Panels(2).Text = "Records [" & Val(rst.RecordCount) & "]"
       
End Sub

'----------------------------------------------------------------------------------------
'       Detecta la existencia o no del link de siguiente pagina
'----------------------------------------------------------------------------------------
Function detectarSiguientePagina(ByRef mDoc As Object) As String 'MSHTML.HTMLDocument) As String

   On Error GoTo detectarSiguientePagina_Error
   '---------------------------------------------------------------------------------------------------------------
    
Dim elems As Object
Dim Item As Object
Dim nextLink As Object
Dim posicion As Integer

If mDoc Is Nothing Then
    detectarSiguientePagina = ""
    
    Open App.Path & "\categoriesErrs.txt" For Append As #4
    Print #4, Now & ""
    Print #4, Now & "State: " & rsState.Fields("state").Value & " - Type: " & LinksParsed.categoriaActual
    Print #4, Now & "Link " & WebBrowser1.LocationURL
    Print #4, Now & ""
    Close 4 '// close the text file
    GoTo detectarSiguientePagina_Error

End If

Set elems = mDoc.getElementById("toolbar-btm")
If Not checkObjeto(elems, "HTMLDivElement", "'Toolbar-btn' object not found") Then GoTo detectarSiguientePagina_Error

Set nextLink = Nothing

Set Item = getElementByClass(elems.All, "next")
If Item Is Nothing Then GoTo noMasPaginas

Set nextLink = Item.getElementsByTagName("A").Item(0)
If Not checkObjeto(nextLink, "HTMLAnchorElement", "'Next Lnk' object not found") Then GoTo detectarSiguientePagina_Error

detectarSiguientePagina = nextLink.href
Exit Function

noMasPaginas:

  'logger "No more pages..."
  detectarSiguientePagina = ""

  '---------------------------------------------------------------------------------------------------------------
   On Error GoTo 0
   Exit Function

detectarSiguientePagina_Error:

'   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure detectarSiguientePagina of Formulario frmMain"
    ErrorHandler Err, "detectarSiguientePagina", "Formulario", "frmMain", False
 
End Function

Private Sub XTab1_TabSwitch(ByVal iLastActiveTab As Integer)
  
  Select Case XTab1.ActiveTab
        
        Case 0
            
            List1.Width = XTab1.Width - (List1.Left * 2)
            List1.Height = XTab1.Height - List1.Top - List1.Left
        
        Case 1
            
            DataGrid1.Width = XTab1.Width - (DataGrid1.Left * 2)
            DataGrid1.Height = XTab1.Height - DataGrid1.Top - DataGrid1.Left
            redrawGrid
        Case 2
        
'        If Config("DebugLevel") < 3 Then
'            WebBrowserView.Visible = False
'        Else
            WebBrowserView.Width = XTab1.Width - (WebBrowserView.Left * 2)
            WebBrowserView.Height = XTab1.Height - WebBrowserView.Top - WebBrowserView.Left
        
'        End If
    
    End Select
    
End Sub


Private Sub XTabModo_BeforeTabSwitch(ByVal iNewActiveTab As Integer, bCancel As Boolean)
    
    ' modo normal
    If iNewActiveTab = 0 Then
        XTab1.TabEnabled(2) = False
        XTab1.ActiveTab = 0
    End If

    'modo interactivo
    If iNewActiveTab = 1 Then
        XTab1.TabEnabled(2) = True
        XTab1.ActiveTab = 2
    End If

End Sub
