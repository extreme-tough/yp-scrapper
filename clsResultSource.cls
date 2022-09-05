VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResultSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' PENDIENTE: declarar objeto Recordset ADO local. Por ejemplo:
'Private WithEvents rs As ADODB.Recordset

Public RS As ADODB.Recordset
Public Function existe(strBuscado As String) As Boolean

On Error GoTo ErrorHandler

 If (RS.BOF <> True) Then RS.MoveFirst
    
    RS.Find "business = '" & Replace(strBuscado, "'", "''") & "'"
    
    'rs.Find "business='" & strBuscado & "'"
    
    'rs.Find "Description = '" & Replace(strbuscado, "'", "''")

    ' existe el id?
    If (RS.BOF <> True) And (RS.EOF <> True) Then
        
        existe = True
    
    Else
    
        existe = False

    End If
    
Exit Function

ErrorHandler:
    MsgBox "Analizing [" & strBuscado & "] ADO Error " & Err.Number & " - " & Err.Description, vbInformation

End Function


Public Sub exportToCSV(file As String)
Dim sExportLine As String
Dim hFile As Long
Dim oField As ADODB.Field
On Error GoTo PROC_ERR
Dim rs2 As ADODB.Recordset
    Set rs2 = RS.Clone
    With rs2
        If (.State = adStateOpen) Then
            hFile = FreeFile
            'Open file For Output As hFile
            
            Open file For Binary As hFile ' esto es para soporte internacional, escribe puenteando la conversion a ANSI
           ' Open file For Output As hFile ' esto es para soporte internacional, escribe puenteando la conversion a ANSI
            
            sExportLine = ""
            
            '----Escribo los encabezados ----
            For Each oField In .Fields
'                sExportLine = sExportLine & oField.Name & ","
                sExportLine = sExportLine & """" & oField.Name & ""","
            Next
            sExportLine = Left$(sExportLine, Len(sExportLine) - 1)
            'Print #hFile, sExportLine
           Put #hFile, , sExportLine & vbNewLine ' esto es para soporte internacional, escribe puenteando la conversion a ANSI

           '  Print #hFile, sExportLine
 
            '----Escribo los registros ----
            Do Until .EOF
                sExportLine = ""
                
                For Each oField In .Fields
'                    sExportLine = sExportLine & oField.Value & ","
                     sExportLine = sExportLine & """" & oField.Value & ""","
               '     MsgBox oField.Value
                    'convert back to Unicode
'                    logTest oField.Value & " - " & StrConv(oField.Value, vbUnicode)
                Next
            
  '              sExportLine = Left$(sExportLine, Len(sExportLine) - 1)
                
                'Print #hFile, sExportLine
'               Put #hFile, , sExportLine & vbNewLine ' esto es para soporte internacional, escribe puenteando la conversion a ANSI
           '      Print #hFile, sExportLine & " --- " & StrConv(sExportLine, vbUnicode)
               
               Put #hFile, , StrConv(sExportLine, vbUnicode) & vbNewLine        ' esto es para soporte internacional, escribe puenteando la conversion a ANSI
                .MoveNext
            Loop
        End If
    
    End With
    
PROC_EXIT:
    
    If (hFile <> 0) Then
        Close hFile
    End If
    
PROC_ERR:

  Select Case Err.Number
  Case Is <> 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExportToCVS of Form frmMain"
    Err.clear
    Resume PROC_EXIT
  End Select
End Sub


Private Sub Class_GetDataMember(DataMember As String, Data As Object)
    ' PENDIENTE:  devolver el recordset apropiado según DataMember. Por ejemplo:
    
    Select Case DataMember
    Case ""             ' Predeterminado
        Set Data = Nothing
    Case Else           ' Predeterminado
        Set Data = RS
    End Select
End Sub

Private Sub Class_Initialize()
    'create the disconnected recordset
    Set RS = New ADODB.Recordset

   With RS
        
        .CursorLocation = adUseClient      'Enable index creation
        
        '  Construyo el recordset
        
        .Fields.Append "business", adBSTR ' adWChar, 250, adFldIsNullable
        .Fields.Append "address", adBSTR ', 150
        .Fields.Append "city", adBSTR ', 50
        .Fields.Append "state", adBSTR ', 50
        .Fields.Append "zip", adBSTR ', 350
        .Fields.Append "tel", adBSTR ', 100
        .Fields.Append "email", adBSTR ', 30
        .Fields.Append "website", adBSTR ', 250
     '   .Fields.Append "keyWord", adBSTR ', 250
        .Fields.Append "category", adBSTR ', 250
                
    End With
    
    RS.open
    
'    esto es para indexar un campo (no utilizado en este projecto)
'    RS.Fields("business").Properties("OPTIMIZE") = True
End Sub


'' Grabo el resulset como archivo XLS basado en tags HTML
'Public Function save2XLS_HTML() As Boolean
'Dim sFilter As String
'Dim NomArchivo As String
'Dim res As Integer
'
'    sFilter = "Excel files (*.xls)|*.xls|Text files (*.txt)|*.txt|All files (*.*)|*.*"
'    NomArchivo = ShowSaveFileDialog(sFilter, "xls", App.Path, OFN_OVERWRITEPROMPT)
'
'    If NomArchivo <> "" Then
'
'        logger "Exporting results to target file '" & NomArchivo & "'...."
'        WriteXlsHTMLFileADO RS, NomArchivo
'
'        logger "Exporting tasks complete."
'
'        save2XLS_HTML = True
'    Else
'        save2XLS_HTML = False
'    End If
'End Function
'
'
'' Grabo el resultset por medio de automatizacion de un objeto MSExcel
'Public Function save2XLSAutomation() As Boolean
'
'Dim sFilter As String
'Dim NomArchivo As String
'Dim res As Integer
'
'    sFilter = "Excel files (*.xls)|*.xls|Text files (*.txt)|*.txt|All files (*.*)|*.*"
'    NomArchivo = ShowSaveFileDialog(sFilter, "xls", App.Path, OFN_OVERWRITEPROMPT)
'
'    If NomArchivo <> "" Then
'
'        logger "Exporting results to target file '" & NomArchivo & "'...."
'        WriteXlsFileADO RS, NomArchivo
'
'        logger "Exporting tasks complete."
'
'        save2XLSAutomation = True
'    Else
'       save2XLSAutomation = False
'    End If
'
'End Function

