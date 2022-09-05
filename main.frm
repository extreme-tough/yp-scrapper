VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMainDummy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   11025
   ClientLeft      =   1095
   ClientTop       =   600
   ClientWidth     =   14055
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11025
   ScaleWidth      =   14055
   StartUpPosition =   2  'CenterScreen
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
      Height          =   735
      Left            =   6555
      Picture         =   "main.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   1725
   End
   Begin VB.CommandButton save2XLS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "SAVE TO .XLS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      Picture         =   "main.frx":0555
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   1725
   End
   Begin VB.ComboBox ComboPaises 
      Height          =   315
      ItemData        =   "main.frx":0678
      Left            =   180
      List            =   "main.frx":067A
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1665
      Width           =   3930
   End
   Begin VB.CommandButton searchBtn 
      Caption         =   "Search"
      Enabled         =   0   'False
      Height          =   690
      Left            =   5985
      Picture         =   "main.frx":067C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2070
      Width           =   1230
   End
   Begin VB.OptionButton LanguageOption 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Spanish"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   9045
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   12
      Top             =   1665
      UseMaskColor    =   -1  'True
      Width           =   1320
   End
   Begin VB.OptionButton LanguageOption 
      BackColor       =   &H00FFFFFF&
      Caption         =   "English"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   10575
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   11
      Top             =   1665
      UseMaskColor    =   -1  'True
      Width           =   1320
   End
   Begin VB.CommandButton abortBtn 
      Caption         =   "STOP SEARCH"
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
      Height          =   735
      Left            =   4545
      Picture         =   "main.frx":0932
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   225
      Width           =   1680
   End
   Begin VB.CommandButton save2XLS_2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "SAVE TO .XLS (OPTION2)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10005
      Picture         =   "main.frx":0ABF
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   1680
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   330
      Left            =   8190
      TabIndex        =   3
      Top             =   2520
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   180
      TabIndex        =   1
      Top             =   5520
      Width           =   13740
      _ExtentX        =   24236
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12240
      Top             =   1980
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   13095
      Top             =   1530
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":0BD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":10F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":127D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1400
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1586
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1713
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   10710
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSForms.Label Label2 
      Height          =   195
      Left            =   11700
      TabIndex        =   19
      Top             =   1260
      Visible         =   0   'False
      Width           =   1680
      VariousPropertyBits=   8388627
      Size            =   "2963;344"
      FontHeight      =   165
      FontCharSet     =   177
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "About me..."
      Height          =   240
      Left            =   13095
      TabIndex        =   18
      Top             =   90
      Width           =   870
   End
   Begin MSForms.TextBox searchtxt 
      Height          =   360
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   5475
      VariousPropertyBits=   -1398781925
      Size            =   "9657;635"
      FontName        =   "Lucida Sans Unicode"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.ListBox List1 
      Height          =   1980
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   13815
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "24368;3492"
      SpecialEffect   =   3
      FontName        =   "Lucida Sans Unicode"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Image banderas 
      Height          =   300
      Index           =   1
      Left            =   12375
      Picture         =   "main.frx":189B
      Stretch         =   -1  'True
      Top             =   1575
      Width           =   495
   End
   Begin VB.Image banderas 
      Height          =   300
      Index           =   0
      Left            =   12375
      Picture         =   "main.frx":1A16
      Stretch         =   -1  'True
      Top             =   1575
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Language selection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9315
      TabIndex        =   13
      Top             =   1305
      Width           =   2130
   End
   Begin VB.Image Image4 
      Height          =   1215
      Left            =   3105
      Picture         =   "main.frx":1B8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11100
   End
   Begin VB.Image Image3 
      Height          =   1200
      Left            =   -45
      Picture         =   "main.frx":1D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3180
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Keyword to search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   10
      Top             =   2115
      Width           =   1995
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   9
      Top             =   1395
      Width           =   1950
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      Height          =   285
      Left            =   8190
      TabIndex        =   4
      Top             =   2250
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "Events registered during the process"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   135
      TabIndex        =   2
      Top             =   3195
      Width           =   13785
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   2070
      Left            =   45
      Picture         =   "main.frx":41F2
      Stretch         =   -1  'True
      Top             =   1170
      Width           =   14115
   End
   Begin VB.Image Image1 
      Height          =   14130
      Left            =   -1935
      Picture         =   "main.frx":436A
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   16140
   End
End
Attribute VB_Name = "frmMainDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents RS As ADODB.Recordset
Attribute RS.VB_VarHelpID = -1

Private WithEvents mHttp As RanaInside.HTTPConnect
Attribute mHttp.VB_VarHelpID = -1
Private WithEvents WebBrowser1 As SHDocVw.InternetExplorer
Attribute WebBrowser1.VB_VarHelpID = -1


Public paisActual As Integer
Public langCookie As String

Const post1 = "&Header1%3AbusquedaDrop=&HomeSearchBox1%3ASearchRadiobutton=2&HomeSearchBox1%3AsearchField="
Const post2 = "&HomeSearchBox1%3AcityBox=&HomeSearchBox1%3AstateDrop=0&HomeSearchBox1%3AsearchBt.x=45&HomeSearchBox1%3AsearchBt.y=24"

Const linkEsp = "http://www.paginasamarillas.com/reLocalization.aspx?UICulture=es-CO&ref="
Const linkUS = "http://www.paginasamarillas.com/reLocalization.aspx?UICulture=en-US&ref="

'Const patchCodePageHTML = "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
'<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
'<meta http-equiv="Content-Type" content="text/html; charset=win 1252">

Dim VIEWSTATE  As String
Public postDetail As String
Const mainPagina = "http://www.paginasamarillas.com/pagamanet/web/home.aspx?ipa="

Const workUrl = "http://www.paginasamarillas.com/pagamanet/web/"

Dim urlActualID As String  ' esto reemplaza al urlID que tenia al otro obj http y ahora lo tomo como una variable global


'==========================================================================
'---------------------------- colecciones ---------------------------------
'Public colaUrl As New Collection


'---------------------------- colecciones ---------------------------------
'==========================================================================

'==========================================================================
'---------------------------- colecciones ---------------------------------
' Public LinksParsed As New linksCollection
 Dim LinksParsed As LinkClassCollection

 Dim rsRes As clsResultSource

Private Function mPagina()
    mPagina = mainPagina & paisActual
End Function







Private Sub eraseBtn_Click()
    Set rsRes = Nothing
    
    Set rsRes = New clsResultSource
    Set DataGrid1.DataSource = rsRes.RS

    logger "Erasing the result grid content...."
    
    'actualizo la grilla de resultado
    DataGrid1.ReBind
End Sub

'---------------------------- colecciones ---------------------------------
'==========================================================================



'==========================================================================
Private Sub Form_Load()

urlActualID = ""
searchBtn.enabled = False

Set LinksParsed = New LinkClassCollection
Set rsRes = New clsResultSource
Set DataGrid1.DataSource = rsRes.RS

'----------------------------------------
' objetos de acceso a internet
'----------------------------------------

Set mHttp = New RanaInside.HTTPConnect

Set WebBrowser1 = New SHDocVw.InternetExplorer
Do Until WebBrowser1.Busy = False
     DoEvents
Loop
'----------------------------------------
'----------------------------------------


'------------ configuracion-------------
Config("DebugLevel") = 3
'------------ configuracion-------------

paisActual = -1

frmMain.Caption = appTitle
'Skinner1.BoxesDefaultTitle = appTitle

ComboPaises.AddItem "Todos los Paises"
ComboPaises.ItemData(ComboPaises.NewIndex) = 0

ComboPaises.AddItem "Bolivia"
ComboPaises.ItemData(ComboPaises.NewIndex) = 22

ComboPaises.AddItem "Brasil"
'ComboPaises.ItemData(ComboPaises.NewIndex) = 0

ComboPaises.AddItem "Colombia"
ComboPaises.ItemData(ComboPaises.NewIndex) = 1

ComboPaises.AddItem "Costa Rica"
ComboPaises.ItemData(ComboPaises.NewIndex) = 8

ComboPaises.AddItem "Ecuador"
ComboPaises.ItemData(ComboPaises.NewIndex) = 6

ComboPaises.AddItem "El Salvador"
ComboPaises.ItemData(ComboPaises.NewIndex) = 2

ComboPaises.AddItem "Guatemala"
ComboPaises.ItemData(ComboPaises.NewIndex) = 3

ComboPaises.AddItem "Honduras"
ComboPaises.ItemData(ComboPaises.NewIndex) = 7

ComboPaises.AddItem "Mexico"
ComboPaises.ItemData(ComboPaises.NewIndex) = 9

ComboPaises.AddItem "Nicaragua"
ComboPaises.ItemData(ComboPaises.NewIndex) = 5

ComboPaises.AddItem "Panama"
ComboPaises.ItemData(ComboPaises.NewIndex) = 4

ComboPaises.AddItem "Peru"
ComboPaises.ItemData(ComboPaises.NewIndex) = 20

ComboPaises.AddItem "Puerto Rico"
ComboPaises.ItemData(ComboPaises.NewIndex) = 14

ComboPaises.AddItem "Rep. Dominicana"
ComboPaises.ItemData(ComboPaises.NewIndex) = 23

ComboPaises.AddItem "Venezuela"
ComboPaises.ItemData(ComboPaises.NewIndex) = 15

logger ""
logger "Application starting.", , True
logger ""
StatusBar1.SimpleText = "Application starting..."


LanguageOption(0).Value = True

'cargarPagina mPagina, "pagInicial"
End Sub
'==============================================================================

'================== COLA URL ==============================
'================== COLA URL ==============================

'Se usa para cargar la pagina temporal de disco
Sub navegarPagina(pag As String)
Dim flag As Long
Dim sHeaders

flag = 0 'navNoReadFromCache  ' vbEmpty
sHeaders = "Cookie: " & "CP=null*; " & langCookie & vbCrLf  'culturenameCookie=es-CO ' Add extra headers as needed
'WebBrowser1.navigate pag, 0, vbEmpty, , sHeaders
'WebBrowser1.navigate pag, flag, vbEmpty, , sHeaders


'WebBrowser1.Navigate2 pag
WebBrowser1.Navigate2 pag, , , , sHeaders

Do Until Not WebBrowser1.Busy
DoEvents
Loop

Do While WebBrowser1.document.body Is Nothing
DoEvents
Loop
End Sub


Sub cargarPagina(url As String, urlId As String)
urlActualID = urlId
navegarPagina url
End Sub


'_______________________ DEPRECIADA_______________________________
Sub cargarPaginaDummy(url As String, urlId As String)
Dim mUrl As urlInfo
Timer1.enabled = True
    
    ' Encolo el nuevo pedido
    Set mUrl = New urlInfo
    mUrl.url = url
    mUrl.id = urlId
    
    'colaUrl.Add mUrl

End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label1_Click()
frmSplash.Show vbModal
End Sub

'English
'http://www.paginasamarillas.com/reLocalization.aspx?UICulture=en-US&ref=http://www.paginasamarillas.com/pagamanet/web/home.aspx?ipa=

'Spanish
'http://www.paginasamarillas.com/reLocalization.aspx?UICulture=es-CO&ref=http://www.paginasamarillas.com/pagamanet/web/home.aspx?ipa=

Private Sub LanguageOption_Click(Index As Integer)

searchBtn.enabled = False

Dim mLink As String
    LinksParsed.enabled = True
   
   Select Case Index
        Case 0
            mLink = linkEsp & mPagina
            logger "loading '" & mLink & "'....."
            logger ""
            logger "Wait to complete the language change to Spanish"
            logger ""
            
            banderas(0).Visible = True
            banderas(1).Visible = False
            
           langCookie = "culturenameCookie=en-CO"
            
           cargarPagina mLink, "pagInicial"

        
        Case 1
            mLink = linkUS & mPagina
            logger ""
            logger "Wait to complete the language change to English"
            logger ""
            
            logger "loading '" & mLink & "'....."
            
            langCookie = "culturenameCookie=en-US"
            cargarPagina mLink, "pagInicial"
    
            banderas(1).Visible = True
            banderas(0).Visible = False
    
    End Select

End Sub

Private Sub Picture1_Click()

End Sub


'_______________________ DEPRECIADA_______________________________
Private Sub Timer1_Timer()
Dim mUrl As urlInfo

' se aborto la operacion
If Not LinksParsed.enabled Then Exit Sub

'esta ocupado, espero.....
If mHttp.isBusy Then Exit Sub

'vacio? no hago nada
'If colaUrl.Count <= 0 Then
'    Timer1.enabled = False
    'Exit Sub
'End If

'Set mUrl = colaUrl(1)

cargarPagina2 mUrl.url, mUrl.id

'colaUrl.Remove (1)
End Sub


'_______________________ DEPRECIADA_______________________________
Sub cargarPagina2(pag As String, idPag As String)

mHttp.SetCookie = langCookie

logger "loading '" & pag & "'....."
mHttp.FetchURL pag, , , , , vbLf & "Accept-Charset: ISO-8859-1,utf-8;q=0.7,*;q=0.7" & vbLf & "Accept-Language: es-ar,es;q=0.8,en-us;q=0.5,en;q=0.3", , idPag

End Sub


'================== COLA URL ==============================
'================== COLA URL ==============================


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Button.Value = tbrPressed
End Sub






Sub armarPostDeBusqueda()
Dim str1 As String

    ' Obtengo un string con los caracteres UTF-8 como se ve en un editor hexadecimal para que al convertirlo
    ' despues a url safe represente un post con codepage UTF-8  ' Accept-Charset: ISO-8859-1,utf-8;q=0.7,*;q=0.7
     
    ' ej:
    '   Avisos Electr?nicos  (tiene la 'o' acentuada pero aca no se ve)
    '   Resultado buscado: Avisos+Electr%C3%B3nicos
     
    ' El problema:
    ' La funcion FixForQuery no convierte bien este string y usar la funcion del API URLEncode produce Avisos%20Electr%3Fnicos
    
     
    str1 = Encode_UTF8(searchtxt)
   
   ' Despues de esta funcion 'Avisos ElectrÃ³nicos'
   ' En donde el caracter especial se representa en UTF-8 con dos simbolos
       
   ' Esta funcion es mas reproduce exactamente string que manda un POST usando un codepage UTF-8
   
    str1 = URLEncode(str1)
    ' Despues de esta funcion 'Avisos+Electr%C3%B3nicos'
  

' Original
'postDetail = "__VIEWSTATE=" & mHttp.FixForQuery(VIEWSTATE) & post1 & mHttp.FixForQuery(searchtxt, False) & post2

postDetail = "__VIEWSTATE=" & mHttp.FixForQuery(VIEWSTATE) & post1 & str1 & post2


logger searchtxt & " -> " & "POST: " & mHttp.FixForQuery(searchtxt) & "  " & str1

'postDetail = "__VIEWSTATE=" & VIEWSTATE & post1 & searchTxt & post2

'postDetail = mHttp.FixForQuery(postDetail, False)
'Debug.Print
'Debug.Print postDetail

'    '-------------Write to Text-----------------
'    Open App.Path & "\testDetail.txt" For Output As #2     '// open the text file
'    Write #2, postDetail
'    Close 2 '// close the text file
'    '-------------Write to Text-----------------
End Sub


Sub buscar(sUrl As String, sPost As String)
    Dim bPostData() As Byte
    Dim sHeaders As String

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



Sub buscarDummy(pag As String, params As String)
logger pag
'logger params

mHttp.SetCookie = langCookie
mHttp.FetchURL pag, , , , , , params, "Busqueda"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : myhttpreq_OnError
' Purpose   :
'---------------------------------------------------------------------------------------
'_______________________ DEPRECIADA_______________________________
Private Sub mHttp_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean, uniqueID As String)
    logger "INTERNET Error occured: Number " & Number & " - Description: " & Description
    logger mHttp.errorMessage
End Sub

Private Sub save2XLS_Click()
Dim sFilter As String
Dim NomArchivo As String
Dim res As Integer
    
    sFilter = "Excel files (*.xls)|*.xls|Text files (*.txt)|*.txt|All files (*.*)|*.*"
    NomArchivo = ShowSaveFileDialog(sFilter, "xls", App.Path, OFN_OVERWRITEPROMPT)
   
    If NomArchivo <> "" Then
        logger "Exporting results to target file '" & NomArchivo & "'...."
        WriteXlsHTMLFileADO rsRes.RS, NomArchivo
            
        logger "Exporting tasks complete."
            
        res = MsgBox("Do you want to discard the present result so that a new search starts with a empty grid??", vbInformation + vbYesNo, appTitle)
            
        If res = vbYes Then eraseBtn_Click
    
    Else
            MsgBox "The operation was canceled", vbInformation, appTitle
    End If
End Sub



Private Sub save2XLS_2_Click()

Dim sFilter As String
Dim NomArchivo As String
Dim res As Integer
    
    sFilter = "Excel files (*.xls)|*.xls|Text files (*.txt)|*.txt|All files (*.*)|*.*"
    NomArchivo = ShowSaveFileDialog(sFilter, "xls", App.Path, OFN_OVERWRITEPROMPT)
   
    If NomArchivo <> "" Then
        
        logger "Exporting results to target file '" & NomArchivo & "'...."
        WriteXlsFileADO rsRes.RS, NomArchivo
            
        logger "Exporting tasks complete."
            
        res = MsgBox("Do you want to discard the present result so that a new search starts with a empty grid??", vbInformation + vbYesNo, appTitle)
            
        If res = vbYes Then eraseBtn_Click
    
    Else
            MsgBox "The operation was canceled", vbInformation, appTitle
    End If

End Sub


Private Sub abortBtn_Click()
abortBtn.enabled = False
searchBtn.enabled = True
LinksParsed.enabled = False
'mHttp.Abort

WebBrowser1.Stop

logger ""
logger "Searching process interrupted by the user...."
logger ""

StatusBar1.SimpleText = "Searching process interrupted by the user...."

Timer1.enabled = False

' vacio la cola de url's
'While colaUrl.Count > 0
 '   colaUrl.Remove (1)
'Wend
End Sub


Private Sub searchBtn_Click()
Dim res As Integer

If ComboPaises = "" Then
    MsgBox "You must to select a country before to search something", vbExclamation, appTitle
    Exit Sub
End If

If searchtxt = "" Then
    MsgBox "There is not a search criterion specified.", vbExclamation, appTitle
    Exit Sub
End If

If rsRes.RS.RecordCount > 0 Then
    res = MsgBox(vbNewLine & "Do you want to discard the present result so that a new search starts with a empty grid??" & vbNewLine, vbQuestion + vbYesNo, appTitle)
            
    If res = vbYes Then eraseBtn_Click
End If

abortBtn.enabled = True
searchBtn.enabled = False
LinksParsed.enabled = True

' vacio la cola de url's
'While colaUrl.Count > 0
 '   colaUrl.Remove (1)
'Wend

'vacio la coleccion de links
While LinksParsed.Count > 0
    LinksParsed.Remove (1)
Wend


'armo el string de parametros de busqueda
armarPostDeBusqueda
    
'hace la busqueda de las palabras ingresadas
buscar mPagina, postDetail

logger "Starting search for '" & searchtxt & "'."
'logger postDetail

StatusBar1.SimpleText = "Loading " & mPagina & "......."
End Sub





Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, url As Variant)

If pDisp Is WebBrowser1 Then
    
    'estaba analizando y mande abortar toda operacion?
    If abortBtn.enabled = False And urlActualID <> "pagInicial" Then Exit Sub
    
    Select Case urlActualID
        
        Case "pagInicial"
                
                ' proceso en forma sincronica la pagina para obtener el id de session
                procesarPagInicial WebBrowser1.document
                
                StatusBar1.SimpleText = "Application ready."
                searchBtn.enabled = True
                abortBtn.enabled = False
                
        Case "Busqueda"
  
                logger "Starting Analisis for '" & searchtxt & "'...."
                StatusBar1.SimpleText = "Searching process....waiting........"
    
                analisis WebBrowser1.document
         
        Case "Analisis"
                
                logger "Analizing result for '" & url & "'..."
               
                analisis WebBrowser1.document
    
        Case "Business"
                
                logger "Analizing busuness result for '" & url & "'..."
                StatusBar1.SimpleText = "Analizing....."
                
                analisisBusiness WebBrowser1.document
                        
        Case Else
        
                StatusBar1.SimpleText = "Done."
    
        End Select
    Else
        logger ""
        logger ""
        Set mHttp = Nothing
        Set mHttp = New RanaInside.HTTPConnect
    End If
    
    

End Sub



Sub siguienteLinkEnEspera()
   ' tengo link por procesar aun???
    Dim lnk As LinkClass
    If LinksParsed.unprocesedItems > 0 Then
        Set lnk = LinksParsed.getFirst
            If Not lnk Is Nothing Then
                cargarPagina workUrl & lnk.link, "Analisis"
            Else
                 logger "Searching process completed."
                    abortBtn.enabled = False
                    searchBtn.enabled = True
            End If
    Else
        logger ""
        logger "Process completed."
        logger ""
        
        StatusBar1.SimpleText = "Records [" & Val(rsRes.RS.RecordCount) & "]"
       
       ' LinksParsed.dump
        abortBtn.enabled = False
        searchBtn.enabled = True
    End If

End Sub
'----------------------------------------------------------------
' Obtiene el valor de VIEWSTATE que es usado como id de session
'----------------------------------------------------------------
Sub procesarPagInicial(mDoc As MSHTML.HTMLDocument)
Dim elem As IHTMLElement
Dim elem2
    '-------------Write MSHTML.HTMLDocument object content to Text-----------------
'    Open App.Path & "\paginaInicial.html" For Append As #3     '// open the text file
'    Write #3, "<HTML>" & mDoc.body.outerHTML & "</HTML>"
'    Close 3 '// close the text file
    '-------------Write MSHTML.HTMLDocument object content to Text-----------------
    
If Not mDoc Is Nothing Then

    Set elem2 = mDoc.getElementById("HomeSearchBox1_languageLb")
    If Not elem2 Is Nothing Then Label2 = elem2.innerText
    
    Set elem = mDoc.getElementById("__VIEWSTATE")
    If Not elem Is Nothing Then
        VIEWSTATE = elem.Value
        urlActualID = "ViewStateRecuperado"
    Else
        logger "Error getting VIEWSTATE value. Plese try again in a few minutes."
    End If
    
    'logger VIEWSTATE
Else
    logger "Error processing '" & mDoc.url & "'...."
End If
End Sub

Sub analisis(mDoc As MSHTML.HTMLDocument)

'Dim MSHTMLobj As New MSHTML.HTMLDocument
'Dim mDoc As MSHTML.HTMLDocument
Dim elem As IHTMLElement
            
'Set mDoc = MSTHML_obtenerDocumento(pag, MSHTMLobj)
    
    '-------------Write to Text-----------------
    Open App.Path & "\paginaResultadoInicial.html" For Output As #3 'Append As #3     '// open the text file
    Write #3, "<HTML>" & mDoc.body.outerHTML & "</HTML>"
    Close 3 '// close the text file
    '-------------Write to Text-----------------

If Not mDoc Is Nothing Then
       
    ' parser de la pagina para identificar los componentes buscados
    analizarPagina mDoc

Else
    logger "Error processing '" & mDoc.url & "'...."
End If

End Sub


Sub analisisBusiness(mDoc As MSHTML.HTMLDocument)
'Dim MSHTMLobj As New MSHTML.HTMLDocument
'Dim mDoc As MSHTML.HTMLDocument
Dim elem As IHTMLElement
            
'Set mDoc = MSTHML_obtenerDocumento(pag, MSHTMLobj)

If Not mDoc Is Nothing Then
       
    'procesarBusiness mDoc
    If Not procesarBusiness(mDoc) Then siguienteLinkEnEspera
    
            
Else
    logger "Error processing '" & mDoc.url & "'...."
End If

End Sub



Sub procesarPagInicialDummy(pag As String)
Dim MSHTMLobj As New MSHTML.HTMLDocument
Dim mDoc As MSHTML.HTMLDocument
Dim elem As IHTMLElement
            
Set mDoc = MSTHML_obtenerDocumento(pag, MSHTMLobj)

If Not mDoc Is Nothing Then

    '-------------Write to Text-----------------
'    Open App.Path & "\paginaInicial.html" For Append As #3     '// open the text file
 '   Write #3, "<HTML>" & mDoc.body.outerHTML & "</HTML>"
  '  Close 3 '// close the text file
   ' '-------------Write to Text-----------------
    Set elem = mDoc.getElementById("__VIEWSTATE")
    
    If Not elem Is Nothing Then
        VIEWSTATE = elem.Value
    Else
        logger "Error getting VIEWSTATE value. Plese try again in a few minutes."
    End If
    
    'logger VIEWSTATE
Else
    logger "Error processing '" & pag & "'...."
End If

End Sub


' analizo el primer resultado
Sub analisisInicialDummy(pag As String)

Dim MSHTMLobj As New MSHTML.HTMLDocument
Dim mDoc As MSHTML.HTMLDocument
Dim elem As IHTMLElement
            
Set mDoc = MSTHML_obtenerDocumento(pag, MSHTMLobj)

If Not mDoc Is Nothing Then
       
    ' parser de la pagina para identificar los componentes buscados
    analizarPagina mDoc
            
Else
    logger "Error processing '" & pag & "'...."
End If

End Sub

Sub analisisBusinessDummy(pag As String)
Dim MSHTMLobj As New MSHTML.HTMLDocument
Dim mDoc As MSHTML.HTMLDocument
Dim elem As IHTMLElement
            
Set mDoc = MSTHML_obtenerDocumento(pag, MSHTMLobj)

If Not mDoc Is Nothing Then
       
    'procesarBusiness mDoc
    If Not procesarBusiness(mDoc) Then siguienteLinkEnEspera
    
            
Else
    logger "Error processing '" & pag & "'...."
End If

End Sub







Sub analizarPagina(ByRef mDoc As MSHTML.HTMLDocument)

'Dim mDoc As HTMLDocument
Dim elems As IHTMLElementCollection
Dim htmlItem As IHTMLElement
Dim i As Integer
'Set mDoc = WebBrowser1.document
Dim cont As Integer


'===========================================
' PAGINA DE RESULTADO SOLO DE CATEGORIAS
'===========================================
Set htmlItem = mDoc.getElementById("Category1_HBXLb")
If Not htmlItem Is Nothing Then

    logger "Identified type: categories result page"
    
    ' proceso la pagina
    procesarTipoCategoria mDoc
    
    siguienteLinkEnEspera
  '' Dim objL As IHTMLElement
   ' For Each objL In mDoc.links
    '    Debug.Print objL.className & " - " & objL.id & " - " & objL.innerText & " - "
    'next
    Exit Sub
End If


'===========================================
' PAGINA DE RESULTADO SIN RESULTADOS
'===========================================

Const strEmptyRes = "No fueron encontrados registos relacionados con su solicitud"

Dim TDs As IHTMLElementCollection
Set TDs = mDoc.getElementsByTagName("TD")
If Not TDs Is Nothing Then
    For Each htmlItem In TDs
'        If htmlItem.innerText = "No fueron encontrados registos relacionados con su solicitud" Then
           If InStr(1, htmlItem.innerText, strEmptyRes, vbTextCompare) <> 0 Then
                logger "Identified type: empty result page"
           Exit Sub
        End If
    Next
End If


'==============================================================================================
'                              PAGINA DE RESULTADO : EMPRESAS
'==============================================================================================
' Hay 2 componentes que pueden aparece aqui;
' Pagina de resultados y subcategorias
' Combinaciones de ambos deben ser tratadas para poder capturar todos los resultados posibles.
'==============================================================================================
'Set elems = mDoc.getElementsByName("TableCute")
Dim elemsTableCute As IHTMLElementCollection
Set elemsTableCute = mDoc.getElementsByName("TableCute") 'OJO QUE ESTO SI NO ENCUENTRA NADA DEVUELVE EL OBJECTO PERO DE LARGO=0
    
Set elems = mDoc.links
cont = 0

For Each htmlItem In elems
        
    ' links que estoy  buscando pertenecientes a la clase "link" para subcategorias en la parte superior
    ' de la pagina de resultado
    If htmlItem.className = "link" Then
        cont = cont + 1
        'logger vbTab & htmlItem.innerText & " - " & htmlItem.nameProp
        LinksParsed.Add htmlItem.nameProp, htmlItem.innerText, -1, False
         
    End If
Next

logger "SubCategories detected(" & Str(cont) & ")"
If cont > 0 Then LinksParsed.nextOrder

'==============================================================================================


'If (Not elemsTableCute Is Nothing And cont > 0) Or (elemsTableCute.length > 0 And cont > 0) Then
If elemsTableCute.length > 0 And cont > 0 Then
    logger "Identified type: business+subcategories result page"
    
    If Not procesarBusiness(mDoc) Then siguienteLinkEnEspera
    
    Exit Sub
End If

If Not elemsTableCute Is Nothing And cont = 0 Then

    logger "Identified type: only business result page"
    
    If Not procesarBusiness(mDoc) Then siguienteLinkEnEspera
    
    Exit Sub
End If

'If elemsTableCute Is Nothing And cont > 0 Then
If elemsTableCute.length = 0 And cont > 0 Then
    logger "Identified type: only subcategories result page"
    siguienteLinkEnEspera
    Exit Sub
End If


'HTMLElementSpan.innerHTML = "pouet"
'Debug.Print HTMLElementSpan.innerText

End Sub


Public Function procesarBusiness(ByRef mDoc As MSHTML.HTMLDocument) As Boolean
Dim mObj, nodes As Object
Dim elems As IHTMLElementCollection
Dim htmlItem As IHTMLElement
Dim mTabla As IHTMLTable
Dim row As IHTMLTableRow
Dim cell As IHTMLTableCell

Dim sigPagLnk As String
Dim pos As Integer

Set mObj = mDoc.getElementById("Table2")
'Set nodes = mObj.rows(0).cells(1).childNodes

If mObj Is Nothing Then GoTo myError

Dim strContador As String

On Error GoTo myError

strContador = mObj.rows(0).cells(1).childNodes(1).firstChild.Data & " to " & mObj.rows(0).cells(1).childNodes(3).firstChild.Data _
& "(" & mObj.rows(0).cells(1).childNodes(5).firstChild.Data & ")"
  

'=============================================
' Obtengo la informacion que necesito
'=============================================
Set elems = mDoc.getElementsByName("TableCute")
logger ""
logger "Processing items " & strContador & "...."

For Each mTabla In elems
    'Set row = mTabla.rows(0)
    'Set cell = row.cells(0)
    
    procesarRegistro mTabla 'row 'cell
    
Next

'actualizo la grilla de resultado
DataGrid1.ReBind

'=============================================
'trato de identificar el link de la sig pagina
'=============================================
Set mTabla = mDoc.getElementById("Table3")

If mTabla Is Nothing Then
   
   ' proceso el siguiente link en la cola de espera
    logger "Error detecting next page, following with the next subcategorie...."
    'siguienteLinkEnEspera
    
    procesarBusiness = False ' aviso que no hay nuevas paginas
    Exit Function
    
End If
    
    Set row = mTabla.rows(0)
    Dim img As IHTMLImgElement

    For Each img In mDoc.images
    
        ' esta presente la imagen con el boton que carga la pagina sig??
        If img.nameProp = "b-sig01.gif" Then
                
            Set cell = row.cells(row.cells.length - 1)
            pos = InStr(10, cell.innerHTML, ">", vbTextCompare)
        
            If pos > 0 Then
                sigPagLnk = Mid(cell.innerHTML, 17, pos - 18)
                sigPagLnk = Replace(sigPagLnk, "&amp;", "&")
                logger "Detected next page link" '& sigPagLnk
            Else
                logger "Error trying to get the next page link access."
            End If
       
            'sigo con la siguiente pagina
            cargarPagina workUrl & sigPagLnk, "Business"
        
            procesarBusiness = True ' aviso que hay nuevas paginas
            Exit Function
        End If
    Next

    
   ' proceso el siguiente link en la cola de espera
    logger "No new page detected, following with the next subcategorie...."
    
    procesarBusiness = False ' aviso que no hay nuevas paginas
    'siguienteLinkEnEspera

Exit Function
myError:

Dim ahora As String

Dim errPag As String
    ahora = Year(Date) & "-" & Month(Date) & "-" & Day(Date) & "_" & FormatDateTime(Time(), 4)
    errPag = App.Path & "\errorPage_" & ahora & ".html"
    
    logger "Error processing result for '" & mDoc.url & "'"
    logger "Generating result error page to send to the administrator.    "
    Open errPag For Append As #5      '// open the text file
    Write #5, "<HTML>" & mDoc.body.outerHTML & "</HTML>"
    Close 5 '// close the text file
    
End Function

Sub procesarRegistro(ByRef mTabla As IHTMLTable)  'row As IHTMLTableRow)
Dim res As String
Dim strAryLines() As String
Dim str1 As String
Dim strEmpresa As String

Dim i As Integer
Dim pos, pos1 As Integer
Dim nodes As Object
    
Dim DOMObj As Object
Dim cell As IHTMLTableCell
Dim row As IHTMLTableRow
Dim direccion2 As String

'---- campos recuperados ----

Dim idEmpresa As String
Dim nombre As String
Dim direccion As String
Dim telefonos As String

Dim telefono As String
Dim fax As String
Dim celular As String

Dim pais As String
Dim distrito As String
Dim ciudad As String

Dim website As String

On Error GoTo myError

'---- campos recuperados ----
    
    Set row = mTabla.rows(0)
    Set cell = row.cells(0)
    
    ' nodos del primer TD donde estan los datos principales
    Set nodes = cell.childNodes
    
    '-----------------------------------------------------
    ' nombre de la empresa (posicion fija)
    '-----------------------------------------------------
    'res = vbNewLine & "Business name:" & vbTab
    If nodes(0).nodeName = "A" Then
        Set DOMObj = nodes(0)
        nombre = Trim(DOMObj.innerText)
'        res = res & vbTab & nombre & vbNewLine
        res = res & nombre & " - "
    Else
        res = res & "ERROR!!!" & vbNewLine
    End If
    
    '-----------------------------------------------------
    'busco id de empresa
    '-----------------------------------------------------
'   For i = 1 To nodes.length - 1
'       If nodes(i).nodeName = "A" Then '
'           If nodes(i).className = "boxResBLink2" Then '
                
 '              'lenguaje español???
  '             If LanguageOption.Item(0).Value = True Then
   '                strempresa = "ie="
    '           Else
     '              strempresa = "iem="
      '         End If
       '
        '       pos = InStr(1, nodes(i).nameProp, strempresa, vbTextCompare)
         '
          '     If pos > 0 Then
'                   str1 = Mid(nodes(i).nameProp, pos + Len(strempresa), 255)
 '                  pos = InStr(1, str1, "&", vbTextCompare)
  '
   '                idEmpresa = Val(Mid(str1, 1, pos - 1))
    '               'Debug.Print idEmpresa
     '          End If
      '
  '         End If
'       End If
 '  Next
    
                strEmpresa = mTabla.innerHTML
                pos = InStr(1, strEmpresa, "ie=", vbTextCompare)
                
                ' Lo encontre??
                If pos > 0 Then
                    str1 = Mid(strEmpresa, pos + 3, 255)
                    pos = InStr(1, str1, "&", vbTextCompare)
                    idEmpresa = Val(Mid(str1, 1, pos - 1))
                Else
                    idEmpresa = "-1"
                End If
            
    
        
    '-----------------------------------------------------
    'busco direccion (primer nodo #text)
    '-----------------------------------------------------
    For i = 1 To nodes.length - 1
        If nodes(i).nodeName = "#text" Then Exit For
    Next
    
    'res = res & "Direccion: " & vbTab
    If nodes(i).nodeName = "#text" Then
        Set DOMObj = nodes(i)
        direccion = DOMObj.nodeValue
        
        'Debug.Print direccion & "  - " & Encode_UTF8(direccion)
        
        If Not esPais(direccion) Then
            'res = res & vbTab & direccion & vbNewLine
            res = res & vbTab & " - " & direccion & " - "
        Else
            logger "Country field detected in wrong place. Fixed."
            i = 1
        End If
    Else
        res = res & "ERROR!!!" & vbNewLine
    End If
    
    
    
    '-----------------------------------------------------
    ' busco seg parte de la direccion (pais, prov, estado...)
    ' (segundo nodo #text )
    '-----------------------------------------------------
    For i = i + 1 To nodes.length - 1
        If nodes(i).nodeName = "#text" Then Exit For
    Next
    
    'res = res & "Direccion 2: " & vbTab
    If nodes(i).nodeName = "#text" Then
        Set DOMObj = nodes(i)
        direccion2 = DOMObj.nodeValue
'        res = res & vbTab & direccion2 & vbNewLine
    Else
        res = res & "ERROR!!!" & vbNewLine
    End If
    
    ' separo cada componente de la direccion 2
    
    'pais
    pos = InStr(1, direccion2, "-", vbTextCompare)
    pais = Mid(direccion2, 1, pos - 1)
    
    'distrito
    pos1 = InStr(pos, direccion2, ",", vbTextCompare)
    distrito = Mid(direccion2, pos + 1, pos1 - pos - 1)
    
    'ciudad
    ciudad = Mid(direccion2, pos1 + 1, 255)
    
    pais = Trim(pais)
    distrito = Trim(distrito)
    ciudad = Trim(ciudad)
    
    'res = res & "pais: " & vbTab & pais & vbNewLine
    'res = res & "distrito: " & vbTab & distrito & vbNewLine
    'res = res & "ciudad: " & vbTab & ciudad & vbNewLine
    
    res = res & pais & "," & distrito & " (" & ciudad & ") "
       
    '-----------------------------------------------------
    ' Todo el texto que encuentre deberia ser Telefono,
    ' Conmutador o Celular
    '-----------------------------------------------------
    'res = res & "Telefono: " & vbTab
    For i = i + 1 To nodes.length - 1
        If nodes(i).nodeName = "#text" Then
            
            If nodes(i).nodeName = "#text" And nodes(i).nodeValue <> " " Then
                Set DOMObj = nodes(i)
                str1 = Trim(DOMObj.nodeValue)
                ' armo el string de nro de telefonos separados por CrNl
'                If InStr(1, DomObj.nodeValue, "Conmutador:", vbTextCompare) > 0 Or _
 '               InStr(1, DomObj.nodeValue, "Telefonos:", vbTextCompare) > 0 Or _
  '              InStr(1, DomObj.nodeValue, "Fax:", vbTextCompare) > 0 Or _
   '             InStr(1, DomObj.nodeValue, "Celular:", vbTextCompare) > 0 Then
                    
                 If Not IsNumeric(Left(str1, 1)) And Left(str1, 1) <> "(" Then
                    
                    
                    'telefonos = telefonos & vbNewLine & DomObj.nodeValue
                    telefonos = telefonos & vbNewLine & str1
                    
                Else
                    
                    'telefonos = telefonos & DomObj.nodeValue
                    telefonos = telefonos & str1
                
                End If
                            
            End If
            
        End If
    Next
    
    ' saco las comas para que no interfiera cuando exporto a CSV
    telefonos = Replace(telefonos, ",", " ")
    
    '...........................................................................
    ' separo los distintos tipos de telefonos encontrados
    strAryLines = Split(telefonos, vbNewLine)
    
    For i = 0 To UBound(strAryLines)
        strAryLines(i) = Trim(strAryLines(i))
'        pos = InStr(1, strAryLines(i), "Conmutador:", vbTextCompare)
 '       If pos > 0 Then telefono = telefono & Mid(strAryLines(i), 12, 500)
  '
   '     pos = InStr(1, strAryLines(i), "Telefonos:", vbTextCompare)
    '    If pos > 0 Then telefono = telefono & Mid(strAryLines(i), 11, 500)
     '
''        pos = InStr(1, strAryLines(i), "Fax:", vbTextCompare)
        'If pos > 0 Then fax = fax & Mid(strAryLines(i), 5, 500)
  ''
    '    pos = InStr(1, strAryLines(i), "Celular:", vbTextCompare)
     '   If pos > 0 Then celular = celular & Mid(strAryLines(i), 9, 500)
        
        str1 = "CONMUTADOR:"
        pos = InStr(1, UCase(strAryLines(i)), str1, vbTextCompare)
        If pos > 0 Then
            telefono = telefono & Mid(strAryLines(i), Len(str1) + 1, 500)
            GoTo siguiente
        End If
                
        str1 = "PHONES:"
        pos = InStr(1, UCase(strAryLines(i)), str1, vbTextCompare)
        If pos > 0 Then
            telefono = telefono & Mid(strAryLines(i), Len(str1) + 1, 500)
            GoTo siguiente
        End If
                
        str1 = "TELEFONOS:"
        pos = InStr(1, UCase(strAryLines(i)), str1, vbTextCompare)
        If pos > 0 Then
            telefono = telefono & Mid(strAryLines(i), Len(str1) + 1, 500)
            GoTo siguiente
        End If
    
        str1 = "TELEFAX:"
        pos = InStr(1, UCase(strAryLines(i)), str1, vbTextCompare)
        If pos > 0 Then
            fax = fax & Mid(strAryLines(i), Len(str1) + 1, 500)
            GoTo siguiente
        End If
        
        str1 = "FAX:"
        pos = InStr(1, UCase(strAryLines(i)), str1, vbTextCompare)
        If pos > 0 Then
            fax = fax & Mid(strAryLines(i), Len(str1) + 1, 500)
            GoTo siguiente
        End If
    
        str1 = "CELULAR:"
        pos = InStr(1, UCase(strAryLines(i)), str1, vbTextCompare)
        If pos > 0 Then
            celular = celular & Mid(strAryLines(i), Len(str1) + 1, 500)
            GoTo siguiente
        End If

        str1 = "BEEPER:"
        pos = InStr(1, UCase(strAryLines(i)), str1, vbTextCompare)
        If pos > 0 Then
            celular = celular & Mid(strAryLines(i), Len(str1) + 1, 500)
            GoTo siguiente
        End If
                
siguiente:
    Next
    '...........................................................................
    
    'res = res & telefonos & vbNewLine
       
'    res = res & "Tel: " & telefono
 '   res = res & vbNewLine & "Fax: " & fax
  '  res = res & vbNewLine & "Cel: " & celular & vbNewLine
    
    res = res & telefono & " ; " & fax & " ; " & celular
    
    '-----------------------------------------------------
    'obtengo el website si existe
    ' Deberia estar en el segundo registro de la tabla
    '-----------------------------------------------------
    Set row = mTabla.rows(1)
    Set nodes = mTabla.rows(1).cells(0).childNodes 'row.cells(0).childNodes
    
    For i = 0 To nodes.length - 1
        If nodes(i).nodeName = "A" Then
            
            Debug.Print nodes(i).innerText
            
            If InStr(1, nodes(i).innerText, "Web", vbTextCompare) > 0 Then
                'existe!
                
                pos = InStr(1, nodes(i).nameProp, "web=", vbTextCompare)
                If pos > 0 Then website = "http://" & Mid(nodes(i).nameProp, pos + 4, 500)
                
                website = Trim(website)
                
                res = res & "Website: " & website & vbTab
                Exit For
            End If
            
        End If
    Next
    
    '-----------------------------------------------------
    '   Categoria
    '-----------------------------------------------------
    If LinksParsed.categoriaActual = "" Then LinksParsed.categoriaActual = searchtxt.Text
    
    'res = res & "Categoria: " & vbTab & LinksParsed.categoriaActual
    res = res & " [" & LinksParsed.categoriaActual & "]"
    
    With rsRes.RS
        If Not rsRes.existe(nombre) Then
'        logger "### >>>>> " & nombre & " already exists in the database. Ignored"
'        Exit Sub
'    End If
    
        '---------------------------------------------------------------
        '===============================================================
        '    R E C O R D S E T    D E   R E S U L T A D O
        '===============================================================
        '---------------------------------------------------------------
'        On Error GoTo ErrorHandler
        

        .AddNew
        
        .Fields("idBusiness") = idEmpresa
        .Fields("business") = nombre
        .Fields("address") = direccion
        .Fields("country") = pais
        .Fields("district") = distrito
        .Fields("city") = ciudad
        .Fields("tel") = telefono
        .Fields("fax") = fax
        .Fields("cel") = celular
        .Fields("website") = website
        .Fields("categories") = LinksParsed.categoriaActual
    
        .Update
        
        '===============================================================
        '---------------------------------------------------------------
    Else
           
        strAryLines = Split(.Fields("categories"), ",")
    
        'Reviso que no sea una categoria repetida.
        For i = 0 To UBound(strAryLines)
            If strAryLines(i) = LinksParsed.categoriaActual Then
                logger "### >>>>> " & nombre & " in category '" & LinksParsed.categoriaActual & "' already exists in the database. Ignored."
                Exit Sub
            End If
        Next
        
        logger "  >>>>> New categorie detected for " & nombre & "."
       
        'La funcion 'existe()' ya deberia haber posicionado el bookmark donde esta el registro
        .Fields("categories") = .Fields("categories") & "," & LinksParsed.categoriaActual
        .Update
        
    End If
    
    End With
   
    logger res
    
    logTest ""
    logTest "Inmediatamente despues de la captura:"
    logTest res
    
 Exit Sub
 
myError:

Dim ahora As String

Dim errPag As String
    ahora = Year(Date) & "-" & Month(Date) & "-" & Day(Date) & "_" & Hour(Now) & Minute(Now) & Second(Now)
    errPag = App.Path & "\errorPageBus_" & ahora & ".html"
    
    logger "Error processing result for '" & nombre & "'"
    logger "Partial process: " & res
    logger "Generating result error page '" & errPag & "' to send to the administrator."
    Open errPag For Append As #4      '// open the text file
    Write #4, "<HTML><BODY>" & mTabla.outerHTML & "</BODY></HTML>"
    Close 4 '// close the text file
    
    logger "Error " & Err.Number & " - " & Err.Description

End Sub

Function esPais(campo As String) As Boolean
Dim pos As Integer
Dim pais As String

     pos = InStr(1, campo, "-", vbTextCompare)
    If pos > 0 Then
        pais = Trim(Mid(campo, 1, pos - 1))
    
        ' convierto a url safe
        pais = URLEncode(pais)
      
        ' No importa el idioma siempre son estos porque los nombres de paises en el resultado respeta los valores en castellano
        If pais = "REPÚBLICA DOMINICANA" Or pais = "COLOMBIA" Or pais = "BOLIVIA" Or pais = "COSTA RICA" Or pais = "ECUADOR" Or _
        pais = "EL SALVADOR" Or pais = "GUATEMALA" Or pais = "HONDURAS" Or pais = "M%3FXICO" Or pais = "NICARAGUA" Or _
        pais = "PANAMÁ" Or pais = "PERÚ" Or pais = "PUERTO RICO" Or pais = "VENEZUELA" Then
        
            esPais = True
        Else
            esPais = False
        End If
    End If
    
End Function

Public Sub procesarTipoCategoria(ByRef mDoc As MSHTML.HTMLDocument)

Dim DOMObj As Object
Dim nodes As Object
Dim categorieObj As Object

'Dim mDoc As HTMLDocument
'Dim elems As IHTMLElementCollection
Dim i As Integer
Dim registro As LinkClass

'Set mDoc = WebBrowser1.document

Set categorieObj = mDoc.getElementById("Category1_resultCategory")

If categorieObj Is Nothing Then
    logger "Elements not found here, review the html code."
    Exit Sub
End If

Set nodes = categorieObj.childNodes

'Set elems = mDoc.links

Dim cont As Integer
cont = 0

'For i = 0 To elems.length - 1
'    ' links que estoy  buscando pertenecientes a la clase "boxResSecLink"
'    If elems(i).className = "boxResSecLink" Then
'        cont = cont + 1
'        Set registro = LinksParsed.Add(elems(i).nameProp, elems(i).innerText, -1, False)
'    End If
'Next

 For i = 0 + 1 To nodes.length - 1
   
    '.......................................................
    ' Solo me interesan las tablas y los links, otros obj dom pueden no tener las propiedades que utilizo y producir algun error EJ: "#text"
    If nodes(i).nodeName = "TABLE" Or nodes(i).nodeName = "A" Then
        
        If nodes(i).nodeName = "TABLE" And InStr(1, nodes(i).innerText, "Did You Mean", vbTextCompare) > 0 Then
                logger "'Did You Mean...?' link detected. No looks for more subcategories."
                Exit For
        End If
        
        ' links que estoy  buscando pertenecientes a la clase "boxResSecLink"
        If nodes(i).className = "boxResSecLink" Then
            cont = cont + 1
            Set registro = LinksParsed.Add(nodes(i).nameProp, nodes(i).innerText, -1, False)
        End If
    End If
    '.......................................................
 Next


logger "Categories detected(" & Str(cont) & ")"

If cont > 0 Then LinksParsed.nextOrder

End Sub


Private Sub WebBrowser1_NavigateError(ByVal pDisp As Object, url As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
    logger "Error (" & StatusCode & ") - " & url
End Sub

'This to make the progress bar work and to show a status message, and an image.
Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
    If Progress = -1 Then ProgressBar1.Value = 100 'the name of the progress bar is "ProgressBar1".
        Label4.Caption = "Done"
        'ProgressBar1.Visible = False 'This makes the progress bar disappear after the page is loaded.
       ' Image1.Visible = True
    If Progress > 0 And ProgressMax > 0 Then
        'ProgressBar1.Visible = True
        'Image1.Visible = False
        If (Progress * 100 / ProgressMax) > 100 Then
            ProgressBar1.Value = 100
        Else
            ProgressBar1.Value = Progress * 100 / ProgressMax
        End If
        
        Label4.Caption = "Loading " & Int(Progress * 100 / ProgressMax) & "%..."
    End If
    Exit Sub
End Sub


