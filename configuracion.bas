Attribute VB_Name = "configuracion"
Option Explicit

Public Const appTitle = "YellowPages.com Bot"

'----------------------------------------------------------------------------------
'        C o n f i g u r a t i o n C l s        s e t t i n g s
'----------------------------------------------------------------------------------
Public Config As ConfigurationCls

Public Sub configInicial()

Set Config = New ConfigurationCls

With Config

    .setIniFile App.Path & "\config.ini"
    
    .Add "DebugLevel", "Settings", "1"
    .Add "retryAfterError", "Settings", "10"
    .Add "connectionTimeout", "Settings", "40"
    .Add "DebugPwd", "Settings", ""
        
    .Add "connectionString", "Sql", ""
    
    .Add "lastAnalizedLink", "session", ""
    .Add "lastAnalizedCat", "session", ""
    .Add "persistentSession", "session", "1"

        
    .Add "startX", "session", frmMain.Left
    .Add "startY", "session", frmMain.Top
    .Add "Width", "session", frmMain.Width
    .Add "height", "session", frmMain.Height
    .Add "state", "session", frmMain.WindowState
    
    .getFromArchive
 
     If Config("DebugLevel") >= 3 And Config("debugPwd") <> debugPassword Then
        Config("DebugLevel") = 2
        logger "Debug level: " & Config("debuglevel")
    End If

    ' set connection string
    Config("connectionstring") = ConnString("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & defaultDBName & ";")
    If Config("connectionstring") = "" Then End
    
End With


End Sub
'----------------------------------------------------------------------------------
'        C o n f i g u r a t i o n C l s        s e t t i n g s
'----------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : checkConnString
' DateTime  : 29/03/2007 10:34
' Author    : Administrador
' Purpose   : Checkea que este configurado correctamente un string de conexion
'
'           NECESITA QUE EXISTE EL OBJETO Config("connectionString")
'
'  defauldb para una base access en el directorio actual:
'  "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & defaultDB & ";"
'---------------------------------------------------------------------------------------
'
Function ConnString(Optional defDB As String) As String

Dim cadena As String
   On Error GoTo checkConnString_Error
   '---------------------------------------------------------------------------------------------------------------

    cadena = Config("connectionString")

    If cadena = "" Then

        'cadena de conexion default para que no tenga que buscar el archivo en otro lado
        If defDB <> "" Then cadena = defDB

        Config("connectionString") = GetConnectString(cadena)
    
    End If
    
    ConnString = cadena
    
  '---------------------------------------------------------------------------------------------------------------
   On Error GoTo 0
   Exit Function

checkConnString_Error:

'   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkConnString of Módulo configuracion"
    ErrorHandler Err, "checkConnString", "Módulo", "configuracion"


End Function
    




Public Sub logger(msg As String, Optional level As Integer = 1, Optional overwrite As Boolean = False)
Dim hora As String
Dim msgTmp As String

If level > Config("DebugLevel") Then Exit Sub

With frmMain
    msgTmp = msg
    msgTmp = Replace(msgTmp, vbNewLine, "<nl>")
    msgTmp = Replace(msgTmp, vbTab, "<tab>")
    msgTmp = Replace(msgTmp, vbTab, "<tab>")
    msgTmp = Replace(msgTmp, vbLf, "<lf>")
       
   
    If .List1.ListCount > 1000 Then .List1.RemoveItem (0)

    'If level > Config("DebugLevel") Then Exit Sub
    hora = FormatDateTime(Now(), vbLongTime)
    
    If Len(msg) > 255 Then msgTmp = Left$(msg, 255)
    
        .List1.AddItem hora & " " & msgTmp
        .List1.Selected(.List1.ListCount - 1) = True
    
    
End With

'    '-------------Write to Text-----------------
    If Not overwrite Then
        Open App.Path & "\logger.txt" For Append As #2     '// open the text file
    Else
        Open App.Path & "\logger.txt" For Output As #2     '// open the text file
    End If
    Print #2, hora & " " & msg
    Close 2 '// close the text file
'    '-------------Write to Text-----------------

End Sub








'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Procedure : ErrorHandler
' DateTime  : 16/02/2007 12:37
' Author    : Administrador
' Purpose   :  Funcion principal de manejo de errores
'               -errores normales
'               -informa la coleccion de errores ado
'               -log de mensajes a archivo
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'



Public Sub ErrorHandler(ObjErr As ErrObject, proc As String, tipo As String, nombre As String, Optional showMsg As Boolean = True)

   Dim errLoop As Error
    Dim strError As String
    Dim strTmp As String
    Dim Errs1 As Errors
    Dim i As Integer
    
    i = 1

    'strTmp = "Error [" & Err.Number & "] en el procedimiento " & proc & "  del " & tipo & " '" & nombre & "'" & vbNewLine
    strTmp = "Error [" & Err.Number & "] in procedure " & proc & " - " & tipo & " '" & nombre & "'" & vbNewLine
    
       ' Process
     'strTmp = strTmp & vbCrLf & "Generado por: " & Err.Source
     'strTmp = strTmp & vbCrLf & "Descripcion:  " & Err.Description
     strTmp = strTmp & vbCrLf & "Generated by: " & Err.Source
     strTmp = strTmp & vbCrLf & "Description : " & Err.Description
     strTmp = strTmp & vbNewLine
     
'   ' Enumerate Errors collection and display properties of
'   ' each Error object.
'     Set Errs1 = DEnv.Con1.Errors
'     For Each errLoop In Errs1
'          With errLoop
'            strTmp = strTmp & vbCrLf & "Error #" & i & ":"
'            strTmp = strTmp & vbCrLf & "   ADO Error   #" & .Number
''            strTmp = strTmp & vbCrLf & "   Descripcion  " & .Description
'            strTmp = strTmp & vbCrLf & "   Description  " & .Description
'            strTmp = strTmp & vbCrLf & "   Source       " & .Source
'            i = i + 1
'       End With
'    Next
    
    Open App.Path & "\" & archivoLog For Output As #3
    'Print #3, "Error reportado : " & Now
    Print #3, "Error reported : " & Now
    Print #3, strTmp
    Print #3, ""
    Close 3
                
    
    logger strTmp
    
    
    'strTmp = strTmp & vbNewLine & "Se ha generado un registro de este error en el archivo:" & vbNewLine & App.Path & "\" & archivoLog & vbNewLine & MSGADMINLOG
    strTmp = strTmp & vbNewLine & "A error report file had been generated con the previous message:" & vbNewLine & App.Path & "\" & archivoLog & vbNewLine & "Keep this file safe, it can be requested when you report this error."
    
    If showMsg Then MsgBox strTmp, vbExclamation

End Sub


'*************************************************
' Name: GetConnectString
' Description:Build or modify a connection string
' By: Mikey
'
' Inputs:optional existing connection string
'
' Returns:string containing a connection  string
'
' Assumes:n/a
'--------------------------------------------------------------------------------------------------------------
' REFERENCIAS NECESARIAS:
'--------------------------------------------------------------------------------------------------------------
' [x] Microsoft OLE DB Service Component 1.0 Type Library
' [x] Microsoft Active Data Object 2.5 Library
'--------------------------------------------------------------------------------------------------------------

Public Function GetConnectString(ConString As String) As String
    'Purpose of Function: Build or modify a  Connection String
    'References required: oledb32.dll, msado15.dll
    'Usage:
        'Constring = GetConnectString(ExistingConnectionString) '---- To Modify Or
        'Constring = GetConnectString("")                       '---- To Create New
    'Returns a standard connection string

   On Error GoTo GetConnectString_Error
   '---------------------------------------------------------------------------------------------------------------

    GetConnectString = ""
    Dim varDataLink As MSDASC.DataLinks
    On Error GoTo GetConnectString_Error
    Set varDataLink = New MSDASC.DataLinks


    If ConString = "" Then
        GetConnectString = varDataLink.PromptNew
    Else
        Dim b As Boolean
        Dim o As Object
        Set o = New ADODB.Connection
        o.ConnectionString = ConString
        b = varDataLink.PromptEdit(o)
        GetConnectString = o.ConnectionString
    End If

  '---------------------------------------------------------------------------------------------------------------
   On Error GoTo 0
   Exit Function


GetConnectString_Error:

    GetConnectString = ""
    Set varDataLink = Nothing

'   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetConnectString of Módulo configuracion"
    ErrorHandler Err, "GetConnectString", "Módulo", "configuracion"

End Function
