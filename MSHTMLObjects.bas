Attribute VB_Name = "MSHTMLObjects"
Public Const SPECIALCHARS = 1
Public Const SPECIALNO = 0


Private Const MAX_PATH                   As Long = 260
Private Const ERROR_SUCCESS              As Long = 0

'Treat entire URL param as one URL segment
Private Const URL_ESCAPE_SEGMENT_ONLY    As Long = &H2000
Private Const URL_ESCAPE_PERCENT         As Long = &H1000
Private Const URL_UNESCAPE_INPLACE       As Long = &H100000

'escape #'s in paths
Private Const URL_INTERNAL_PATH          As Long = &H800000
Private Const URL_DONT_ESCAPE_EXTRA_INFO As Long = &H2000000
Private Const URL_ESCAPE_SPACES_ONLY     As Long = &H4000000
Private Const URL_DONT_SIMPLIFY          As Long = &H8000000

'Converts unsafe characters,
'such as spaces, into their
'corresponding escape sequences.
Public Declare Function UrlEscape Lib "shlwapi" _
   Alias "UrlEscapeA" _
  (ByVal pszURL As String, _
   ByVal pszEscaped As String, _
   pcchEscaped As Long, _
   ByVal dwFlags As Long) As Long

'Converts escape sequences back into
'ordinary characters.
Public Declare Function UrlUnescape Lib "shlwapi" _
   Alias "UrlUnescapeA" _
  (ByVal pszURL As String, _
   ByVal pszUnescaped As String, _
   pcchUnescaped As Long, _
   ByVal dwFlags As Long) As Long
Public Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Public Function YellowEncodeUrl(ByVal sUrl As String) As String
    Dim retVal As String
    Dim sChar As String
    Dim sLastChar As String
    
    retVal = ""
    sLastChar = "+"
    For i = 1 To Len(sUrl)
        sChar = Mid(sUrl, i, 1)
        If IsNumeric(sChar) Or IsCharAlpha(Asc(sChar)) = 1 Then
            sLastChar = sChar
        Else
            If sLastChar <> "+" Then
                sChar = "+"
                sLastChar = sChar
            Else
                sChar = ""
            End If
        End If
        retVal = retVal + sChar
    Next
    YellowEncodeUrl = retVal
End Function

   
'Funciona bien, pero es lo mismo que hace la funcion del la DLL que suelo usar
Public Function EncodeUrl(ByVal sUrl As String) As String

   Dim buff As String
   Dim dwSize As Long
   Dim dwFlags As Long
   
   If Len(sUrl) > 0 Then
      
      buff = Space$(MAX_PATH)
      dwSize = Len(buff)
      dwFlags = URL_ESCAPE_SEGMENT_ONLY 'URL_DONT_SIMPLIFY
      
      If UrlEscape(sUrl, _
                   buff, _
                   dwSize, _
                   dwFlags) = ERROR_SUCCESS Then
                   
         EncodeUrl = Left$(buff, dwSize)
      
      End If  'UrlEscape
   End If  'Len(sUrl)

End Function
   
Public Function DecodeUrl(ByVal sUrl As String) As String

   Dim buff As String
   Dim dwSize As Long
   Dim dwFlags As Long
   
   If Len(sUrl) > 0 Then
      
      buff = Space$(MAX_PATH)
      dwSize = Len(buff)
      dwFlags = URL_DONT_SIMPLIFY
      
      If UrlUnescape(sUrl, _
                   buff, _
                   dwSize, _
                   dwFlags) = ERROR_SUCCESS Then
                   
         DecodeUrl = Left$(buff, dwSize)
      
      End If  'UrlUnescape
   End If  'Len(sUrl)

End Function
   
   
Public Function MSTHML_obtenerDocumento3(mLink As String) As MSHTML.HTMLDocument
'Dim objLink As HTMLLinkElement
'Dim objMSHTML As New MSHTML.HTMLDocument
Dim objMSHTML As MSHTML.HTMLDocument
Dim objDoc As MSHTML.HTMLDocument
    
    Set objMSHTML = New MSHTML.HTMLDocument
    
    Set objDoc = objMSHTML.createDocumentFromUrl(mLink, vbNullString)
    Debug.Print "objDoc.readyState: " & objDoc.readyState
        
    While objDoc.readyState <> "complete"
        Debug.Print "objDoc.readyState: " & objDoc.readyState
        DoEvents
    Wend
    
   Do While objDoc.body Is Nothing
        Debug.Print "objDoc.readyState: " & objDoc.readyState
        DoEvents
    Loop
    
    For Each objLink In objDoc.links
        Debug.Print objLink
    Next
    
    Set MSTHML_obtenerDocumento3 = objDoc
    
    'get all Links
    
End Function


Public Function MSTHML_obtenerDocumento(mLink As String, ByRef objMSHTML As MSHTML.HTMLDocument)
Dim objLink As HTMLLinkElement
'Dim objMSHTML As New MSHTML.HTMLDocument
Dim objDoc As MSHTML.HTMLDocument
    
    'objMSHTML.defaultCharset = "utf-8"
    'objMSHTML.Charset = "utf-8"
    
    Set objDoc = objMSHTML.createDocumentFromUrl(mLink, vbNullString)
    Debug.Print "objDoc.readyState: " & objDoc.readyState
        
    While objDoc.readyState <> "complete"
       Debug.Print vbTab & "objDoc.readyState: " & objDoc.readyState
        DoEvents
    Wend
    
    
    Do While objDoc.body Is Nothing
        Debug.Print vbTab & "objDoc.readyState: " & objDoc.readyState
        DoEvents
    Loop
    
'    For Each objLink In objDoc.links
 '       Debug.Print objLink
  '  Next
    
    'objDoc.Charset = "utf-8"
    'objdoc.
    'MsgBox objDoc.documentElement.outerHTML
    
    Set MSTHML_obtenerDocumento = objDoc
    'get all Links
   ' Set objMSHTML = Nothing
End Function



'====================================================================================================
' URLEncode - Convert a string for using on a URL query string
'
' convert a string so that it can be used on a URL query string
' Same effect as the Server.URLEncode method in ASP

'Const SPECIALCHARS = 1
'Const SPECIALNO = 0
'====================================================================================================

Function URLEncode(ByVal Text As String, Optional special = SPECIALNO) As String
    Dim i As Integer
    Dim acode As Integer
    Dim char As String
    
    URLEncode = Text
    
    For i = Len(URLEncode) To 1 Step -1
        acode = Asc(Mid$(URLEncode, i, 1))
        Select Case acode
            
            Case Asc("-") And special = SPECIALCHARS
            Case Asc("=") And special = SPECIALCHARS
            Case Asc(".") And special = SPECIALCHARS
            Case Asc("_") And special = SPECIALCHARS
            Case Asc("&") And special = SPECIALCHARS
            
            Case 48 To 57, 65 To 90, 97 To 122
                ' don't touch alphanumeric chars
            Case 32
                ' replace space with "+"
                Mid$(URLEncode, i, 1) = "+"
            Case Else
                    ' replace punctuation chars with "%hex"
                    URLEncode = Left$(URLEncode, i - 1) & "%" & Hex$(acode) & Mid$(URLEncode, i + 1)
        End Select
    Next
    
End Function


'====================================================================================================
' Da una representacion ASCII/Unicode de los caracteres dobles de UTF-8
'====================================================================================================
Public Function Encode_UTF8(ByVal astr As String) As String
  Dim c As Long, n As Long
  Dim utftext As String
    utftext = ""
    For n = 1 To Len(astr)
        c = AscW(Mid(astr, n, 1))
        If c < 128 Then
            utftext = utftext + Mid(astr$, n, 1)
        ElseIf ((c > 127) And (c < 2048)) Then
            utftext = utftext + Chr(((c \ 64) Or 192))
            utftext = utftext + Chr(((c And 63) Or 128))
        Else
            utftext = utftext + Chr(((c \ 144) Or 234))
            utftext = utftext + Chr((((c \ 64) And 63) Or 128))
            utftext = utftext + Chr(((c And 63) Or 128))
        End If
    Next n

  Encode_UTF8 = utftext

End Function


'================================================================================================================
'                          D   O   M  -  H  T  M  L      A  C  C  E  S  S
'================================================================================================================

'---------------------------------------------------------------------------------------
' Procedure : ParametrosURL
' DateTime  : 08/11/2006 19:08
' Author    : Administrador
' Purpose   : Permite obtener un parametro determinado de una URL en formato de string
'---------------------------------------------------------------------------------------
'
Function parametrosURL(strURL As String, parametro As String) As String
Dim params As Collection
Dim resultado As String
    Set params = New Collection
    
    getURLParams strURL, params
    resultado = params(parametro)
    Set params = Nothing
    
    parametrosURL = resultado
    
End Function




Function getElementByClass(elems As Object, clsName As String, Optional Count As Integer = 1) As Object

Dim htmlItem As IHTMLElement
Dim contador As Integer
contador = 0

For Each htmlItem In elems
        
    ' links que estoy  buscando pertenecientes a la clase "link" para subcategorias en la parte superior
    ' de la pagina de resultado
    If htmlItem.className = clsName Then
        contador = contador + 1
                            
        logger clsName & "[" & Str(contador) & "] ->" & htmlItem.innerText, 4
                            
'    '-------------Write to Text-----------------
'    Open App.Path & "\logger.txt" For Append As #2     '// open the text file
 '   Print #2, clsName & "[" & Str(contador) & "] ->" & htmlItem.innerText
  '  Close 2 '// close the text file
'  '  '-------------Write to Text-----------------
                            
        If contador >= Count Then
            Set getElementByClass = htmlItem
            Exit Function
        End If
    End If
Next

Set getElementByClass = Nothing
End Function


' Me sirve para separar parametros en un http request devolviendome una coleccion con los valores encontrados
' Ej:
' http://www.goudengids.nl/contact?mfinfo.linktype=website&url=http%3a%2f%2fwww.wegwijsinreclame.nl&mfinfo.show_listingId=NL_1428347_1&mfinfo.subscriberId=211481504&mfinfo.show_name=Koppelaar%20Reclame%20Service&mfinfo.show_location=Sliedrecht|DRECHTSTEDEN&mfinfo.show_heading=Reclame&mfinfo.show_industryAssociation=&mfinfo.show_brand=&mfinfo.show_productId=ISV&mfinfo.show_zoning=null&mfinfo.show_itemCode=GPKW&mfinfo.show_listingType=BUSINESS
'
Function getURLParams(url As String, ByRef colec1 As Collection) As Boolean
Dim str1() As String
Dim pars As String
Dim pos As Integer, i As Integer

    str1 = Split(url, "?")
    
    If UBound(str1) = 0 Then
        getURLParams = False ' no tiene parametros o invalido
        Exit Function
    End If
    
    pars = str1(1)
    str1 = Split(pars, "&")
                
    For i = 0 To UBound(str1)
    
        ' un '=' puede ser parte del valor y producira una parsing erroneo
        pos = InStr(1, str1(i), "=")
        If pos = 0 Then ' no encontre separador de valores, error
            getURLParams = False
            Exit Function
        End If
                
        colec1.Add Mid(str1(i), pos + 1), Mid(str1(i), 1, pos - 1)
    
    Next
                               
    getURLParams = True
               
End Function


' igual que la anterior, pero para parsear string mas raros como los que hay en llamados a javascript
' Ej
'   javascript:void("url=_eaW5mb0BwYWtuYmFrLm5s|listingId=NL_1294098_130|name=Aanhangwagenreclame.nl")
'
'    donde deberia enviarle a esta funcion esto (Que previamente separe con split):
'   url=_eaW5mb0BwYWtuYmFrLm5s|listingId=NL_1294098_130|name=Aanhangwagenreclame.nl")
'

Function getParams(url As String, ByRef colec1 As Collection, separadorCol As String, Optional separadorValor As String = "=") As Boolean
Dim str1() As String
Dim pars As String
Dim i As Integer
Dim pos As Integer

    str1 = Split(url, separadorCol)
                
    For i = 0 To UBound(str1)
        
        ' un '=' puede ser parte del valor y producira una parsing erroneo
        pos = InStr(1, str1(i), "=")
        If pos = 0 Then ' no encontre separador de valores, error
            getParams = False
            Exit Function
        End If
                
        colec1.Add Mid(str1(i), pos + 1), Trim(Mid(str1(i), 1, pos - 1))
    Next
                               
    getParams = True
               
End Function


'===================================================================================================
' Arma un string POST con todos los elementos del FORM pasado como objecto a esta funcion
' TIENE QUE RECIBIR COMO PARAMETRO UN OBJECTO "FORM"
' Solo loguea para debugging en nivel 1.
'===================================================================================================
'INPUT hidden _dyncharset -> 'ISO-8859-1'  id:
'INPUT radio v -> '13' checked:[Verdadero] id:radiotype
'INPUT hidden _D:v -> ' '  id:
'INPUT radio v -> '12' checked:[Falso] id:radioname
'INPUT hidden _D:v -> ' '  id:
'INPUT text /smartpages/search/SearchFormHandler.searchTerms -> 'sign'  id:q
'INPUT hidden _D:/smartpages/search/SearchFormHandler.searchTerms -> ' '  id:
'INPUT text /smartpages/search/SearchFormHandler.cityZipAreaCode -> ''  id:
'INPUT hidden _D:/smartpages/search/SearchFormHandler.cityZipAreaCode -> ' '  id:
'INPUT hidden _D:/smartpages/search/SearchFormHandler.state -> ' '  id:
'SELECT select-one /smartpages/search/SearchFormHandler.state -> 'NW'  id:
'INPUT checkbox /smartpages/search/SearchFormHandler.saveDefaultLocation -> 'true'  id:
'INPUT hidden _D:/smartpages/search/SearchFormHandler.saveDefaultLocation -> ' '  id:
'INPUT submit Go! -> 'Go!'  id:
'INPUT hidden /smartpages/search/SearchFormHandler.search -> ''  id:
'INPUT hidden _D:/smartpages/search/SearchFormHandler.search -> ' '  id:
'INPUT hidden /smartpages/search/SearchFormHandler.successURL -> 'ypresults.jsp;jsessionid=PEILO15GWG43DQFI21GRNWQ'  id:
'INPUT hidden _D:/smartpages/search/SearchFormHandler.successURL -> ' '  id:
'INPUT hidden /smartpages/search/SearchFormHandler.failureURL -> 'yptransition.jsp;jsessionid=PEILO15GWG43DQFI21GRNWQ'  id:
'INPUT hidden _D:/smartpages/search/SearchFormHandler.failureURL -> ' '  id:
'INPUT hidden _DARGS -> '/sp/index1.jsp'  id:

'_dyncharset=ISO-8859-1&v=13&_D%3Av=+&_D%3Av=+&%2Fsmartpages%2Fsearch%2FSearchFormHandler.searchTerms=sign&_D%3A%2Fsmartpages%2Fsearch%2FSearchFormHandler.searchTerms=+&%2Fsmartpages%2Fsearch%2FSearchFormHandler.cityZipAreaCode=&_D%3A%2Fsmartpages%2Fsearch%2FSearchFormHandler.cityZipAreaCode=+&_D%3A%2Fsmartpages%2Fsearch%2FSearchFormHandler.state=+&%2Fsmartpages%2Fsearch%2FSearchFormHandler.state=NW&%2Fsmartpages%2Fsearch%2FSearchFormHandler.saveDefaultLocation=true&_D%3A%2Fsmartpages%2Fsearch%2FSearchFormHandler.saveDefaultLocation=+&Go%21=Go%21&%2Fsmartpages%2Fsearch%2FSearchFormHandler.search=&_D%3A%2Fsmartpages%2Fsearch%2FSearchFormHandler.search=+&%2Fsmartpages%2Fsearch%2FSearchFormHandler.successURL=ypresults.jsp%3Bjsessionid=PEILO15GWG43DQFI21GRNWQ&_D%3A%2Fsmartpages%2Fsearch%2FSearchFormHandler.successURL=+&%2Fsmartpages%2Fsearch%2FSearchFormHandler.failureURL=yptrnsition.jsp%3Bjsessionid=PEILO15GWG43DQFI21GRNWQ&_D%3A%2Fsmartpages%2Fsearch%2FSearchFormHandler.failureURL=+&_DARGS=%2Fsp%2Findex1.jsp

Public Function extraerForm(mObj As Object) As String
Dim elem As Object
Dim temp As String
    For Each elem In mObj.All
        If elem.tagName = "INPUT" Then
            
            Select Case elem.Type
            
            Case "radio"
                logger "INPUT " & elem.Type & " " & elem.Name & " -> '" & elem.Value & "' checked:[" & elem.checked & "] id:" & elem.id, 3
                        
                If elem.checked = "true" Or elem.checked = "Verdadero" Then temp = temp & elem.Name & "=" & elem.Value & "&"
            
            Case Else
                
                logger "INPUT " & elem.Type & " " & elem.Name & " -> '" & elem.Value & "'  id:" & elem.id, 3
                temp = temp & elem.Name & "=" & elem.Value & "&"
            
            End Select
        End If
            
        If elem.tagName = "SELECT" Then
        
            logger "SELECT " & elem.Type & " " & elem.Name & " -> '" & elem.Value & "'  id:" & elem.id, 3
            temp = temp & elem.Name & "=" & elem.Value & "&"
        
        End If
    
    
    Next
    
extraerForm = Left$(temp, Len(temp) - 1)

End Function

