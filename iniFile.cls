VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"iniFile"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Private m_cIni As CIniFile
'Private Sub Form_Load()
'  Set m_cIni = New CIniFile
'End Sub


Option Explicit

'// Private member that holds a reference to
'// the path of our ini file
Private strInI As String

'// Win API Declares
Private Declare Function WritePrivateProfileString _
  Lib "kernel32" Alias "WritePrivateProfileStringA" _
 (ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpString As Any, _
  ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString _
  Lib "kernel32" Alias "GetPrivateProfileStringA" _
 (ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpDefault As String, _
  ByVal lpReturnedString As String, _
  ByVal nSize As Long, _
  ByVal lpFileName As String) As Long

Private Function MakePath(ByVal strDrv As String, _
  ByVal strDir As String) As String
 
 '// Makes an INI file: Guarantees a sub dir
 Do While Right$(strDrv, 1) = "\"
   strDrv = Left$(strDrv, Len(strDrv) - 1)
 Loop

 Do While Left$(strDir, 1) = "\"
   strDir = Mid$(strDir, 2)
 Loop

 '// Return the path
 MakePath = strDrv & "\" & strDir
End Function

Public Sub CreateIni(strDrv As String, strDir As String)


 '// Make a new ini file
 strInI = MakePath(strDrv, strDir)
End Sub

Public Sub WriteFile(strSection As String, _
 strKey As String, _
 strValue As String)
 
 '// Write to strINI
 WritePrivateProfileString strSection, _
   strKey, strValue, strInI
End Sub

Public Function GetFile(strSection As String, _
  strKey As String, default As String) As String

  Dim strTmp As String
  Dim lngRet As String

  strTmp = String$(1000, 0)
  lngRet = GetPrivateProfileString(strSection, strKey, default, strTmp, Len(strTmp), strInI)
  
  'lngRet = GetPrivateProfileString(strSection, strKey, default, strTmp, 100, strInI)
    
  'GetFile = strTmp
  GetFile = Mid(strTmp, 1, Val(lngRet))

 
End Function

Public Property Let iniFile(ByVal New_IniPath As String)

 '// Sets the new ini path
 strInI = New_IniPath
 
End Property

Public Property Get iniFile() As String

 '// Returns the current ini path
 iniFile = strInI
 
End Property



