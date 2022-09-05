Attribute VB_Name = "ExcelRoutines"
Public Enum saveExcelType
    XlSByAutomation = 1
    XLSAsHMTLTable = 2
End Enum

'===================================================================================
'   Genera un xls real usando automatizacion
'
'===================================================================================
Public Sub CreateExcelFromRS(ByRef rst As ADODB.Recordset, arch As String, Optional crearHeaders As Boolean = True)
On Error GoTo ErrorHandler

    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
 
    Dim recArray As Variant
    
    Dim strDB As String
    Dim fldCount As Integer
    Dim recCount As Long
    Dim iCol As Integer
    Dim iRow As Integer
    
    ' Create an instance of Excel and add a workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets(1)
  
    ' Display Excel and give user control of Excel's lifetime
    xlApp.Visible = False ' True
    xlApp.UserControl = True
    xlApp.DisplayAlerts = False
    
    ' Copy field names to the first row of the worksheet
    fldCount = rst.Fields.Count
    For iCol = 1 To fldCount
        xlWs.cells(1, iCol).Value = rst.Fields(iCol - 1).Name
    Next
        
    ' Check version of Excel
    If Val(Mid(xlApp.version, 1, InStr(1, xlApp.version, ".") - 1)) > 8 Then
        'EXCEL 2000 or 2002: Use CopyFromRecordset
         
        ' Copy the recordset to the worksheet, starting in cell A2
        xlWs.cells(2, 1).CopyFromRecordset rst
        'Note: CopyFromRecordset will fail if the recordset
        'contains an OLE object field or array data such
        'as hierarchical recordsets
        
    Else
        'EXCEL 97 or earlier: Use GetRows then copy array to Excel
    
        ' Copy recordset to an array
        recArray = rst.GetRows
        'Note: GetRows returns a 0-based array where the first
        'dimension contains fields and the second dimension
        'contains records. We will transpose this array so that
        'the first dimension contains records, allowing the
        'data to appears properly when copied to Excel
        
        ' Determine number of records

        recCount = UBound(recArray, 2) + 1 '+ 1 since 0-based array
        

        ' Check the array for contents that are not valid when
        ' copying the array to an Excel worksheet
        For iCol = 0 To fldCount - 1
            For iRow = 0 To recCount - 1
                ' Take care of Date fields
                If IsDate(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = format(recArray(iCol, iRow))
                ' Take care of OLE object fields or array fields
                ElseIf IsArray(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = "Array Field"
                End If
            Next iRow 'next record
        Next iCol 'next field
            
        ' Transpose and Copy the array to the worksheet,
        ' starting in cell A2
        xlWs.cells(2, 1).Resize(recCount, fldCount).Value = _
            TransposeDim(recArray)
    End If

    ' Auto-fit the column widths and row heights
    xlApp.selection.CurrentRegion.Columns.AutoFit
    xlApp.selection.CurrentRegion.rows.AutoFit

    'Save the Workbook and Quit Excel
    xlWs.SaveAs arch
    xlApp.DisplayAlerts = True
    
    xlApp.Quit
    
    ' Release Excel references
    Set xlWs = Nothing
    Set xlWb = Nothing

    Set xlApp = Nothing

Exit Sub

ErrorHandler:
    MsgBox "Excel Error " & Err.Number & " - " & Err.Description, vbInformation
    Resume Next
End Sub


Function TransposeDim(v As Variant) As Variant
' Custom Function to Transpose a 0-based array (v)
    
    Dim X As Long, Y As Long, Xupper As Long, Yupper As Long
    Dim tempArray As Variant
    
    Xupper = UBound(v, 2)
    Yupper = UBound(v, 1)
    
    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = v(Y, X)
        Next Y
    Next X
    
    TransposeDim = tempArray

End Function

        


'================================================================================================
' Este es un CSV mentiroso, porque en realidad es una tabla html que el excel abre directamente
' como si fuera un archivo xls
' CUIDADO: Pasar una copia porque mueve el cursor original
'================================================================================================

Public Sub RecordsetToHTMLXLS(ByRef RS As ADODB.Recordset, ByVal CSVFilePath, Optional IncludeFieldNames As Boolean = True)

Dim sCSV As String
Dim headers As String
RS.MoveFirst

' Copy field names to the first row of the worksheet
Dim fld As ADODB.Field

headers = "<b>"
If IncludeFieldNames Then
    For Each fld In RS.Fields
        headers = headers & fld.Name & "</b></TD>" & vbNewLine & "<TD><b>"
    Next
    
    headers = Left$(headers, Len(headers) - 8)
    headers = headers & "</b></TD></TR><TR><TD>" & vbNewLine & vbNewLine

End If

Debug.Print headers

' convert the recordset to a table
sCSV = RS.GetString(, , "</TD><TD>", "</TD></TR><TR><TD>", "&nbsp;")
' to complete the conversion you must add the opening TR and TD tags
' for the very first cell, and drop the closing TD and TR tags
' after the very last cell
sCSV = "<TABLE BORDER=1><TR><TD>" & headers & Left(sCSV, Len(sCSV) - 8) & "</TABLE>"

'sCSV = RS.GetString(, , ",", vbCrLf)

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFile = oFSO.CreateTextFile(CSVFilePath, True, True)
oFile.Write sCSV
oFile.Close: Set oFile = Nothing
Set oFSO = Nothing

End Sub
'===================================================================================================



'===================================================================================================
' Crea un archivo unicode basado en el contenido de un recordset
'===================================================================================================
Private Sub CreateFile(ByVal pstrFile As String, ByVal pstrData As String)

  Dim objStream As Object
 
  'Create the stream
  Set objStream = CreateObject("ADODB.Stream")

  'Initialize the stream
  objStream.open

  'Reset the position and indicate the charactor encoding
  objStream.position = 0
'  objStream.Charset = "UTF-8"
 objStream.Charset = "UTF-8"
  'Write to the steam
  objStream.WriteText pstrData
 
  'Save the stream to a file
  objStream.SaveToFile pstrFile
 
End Sub




'=============================================================================
'=============================================================================
'-----------------------------------------------------------------------------
'          FUNCIONES DE CONVERSION DE RECORDSET A CSV O XLS
'-----------------------------------------------------------------------------
'=============================================================================
'=============================================================================

'Public Sub saveAsCSV(file As String, Optional headers As Boolean = True)
'---------------------------------------------------------------------------------------
' Procedure : exportToCSV
' DateTime  : 02/04/2007 09:17
' Author    : Administrador
' Purpose   : Rutina general de exportacion a CSV file con Dialogs para especificar el
'               nombre de la archivo
'---------------------------------------------------------------------------------------
'
Function exportToCSV(mRst As ADODB.Recordset, Optional headers As Boolean = True, Optional quoted As Boolean = True, Optional defaultName As String = "", Optional AsUnicode As Boolean = False) As Boolean
Dim sFilter As String
Dim NomArchivo As String
Dim res As Integer

   On Error GoTo exportToCSV_Error
   '---------------------------------------------------------------------------------------------------------------

    sFilter = "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|All files (*.*)|*.*"
    NomArchivo = ShowSaveFileDialog(sFilter, "csv", App.Path, OFN_OVERWRITEPROMPT, , defaultName)
   
    If NomArchivo <> "" Then
        
        logger "Exporting results to target file '" & NomArchivo & "'...."
        
        Dim rs2 As ADODB.Recordset
        Set rs2 = mRst.Clone
        
        saveAsCSV rs2, NomArchivo, headers, quoted, AsUnicode
        rs2.Close
    
       logger "Exporting tasks complete."
            
        exportToCSV = True
    Else
        'MsgBox "The operation was canceled", vbInformation, appTitle
        exportToCSV = False
    End If

  '---------------------------------------------------------------------------------------------------------------
   On Error GoTo 0
   Exit Function

exportToCSV_Error:

'   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure exportToCSV of Módulo ExcelRoutines"
    ErrorHandler Err, "exportToCSV", "Módulo", "ExcelRoutines"


End Function

'---------------------------------------------------------------------------------------
' Procedure : exportToCSV
' DateTime  : 01/04/2007 15:27
' Author    : Administrador
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub saveAsCSV(rs2 As ADODB.Recordset, file As String, Optional headers As Boolean = True, Optional quoted As Boolean = True, Optional AsUnicode As Boolean = False)
Dim sExportLine As String
Dim hFile As Long
Dim oField As ADODB.Field
Dim cont As Long
   On Error GoTo exportToCSV_Error
   '---------------------------------------------------------------------------------------------------------------
    Dim fExp As frmExport
    Set fExp = New frmExport
    
    fExp.progBar.Max = rs2.RecordCount
    fExp.Show
    
    'Dim rs2 As ADODB.Recordset
    'Set rs2 = RS.Clone
    
    cont = 0
    With rs2
        If (.State = adStateOpen) Then
            hFile = FreeFile
            If AsUnicode Then
                Open file For Binary As hFile ' esto es para soporte internacional, escribe puenteando la conversion a ANSI
            Else
                Open file For Output As hFile
            End If
            
            sExportLine = ""
            
            If headers Then
                '----Escribo los encabezados ----
                For Each oField In .Fields
                    If quoted Then
                        sExportLine = sExportLine & """" & oField.Name & ""","
                    Else
                        sExportLine = sExportLine & oField.Name & ","
                    End If
                Next
            End If
        
            sExportLine = Left$(sExportLine, Len(sExportLine) - 1)
            
            If AsUnicode Then
                Put #hFile, , sExportLine & vbNewLine ' esto es para soporte internacional, escribe puenteando la conversion a ANSI
            Else
                Print #hFile, sExportLine
            End If

            
 
            '----Escribo los registros ----
            Do Until .EOF
                sExportLine = ""
                
                For Each oField In .Fields
                    If quoted Then
                        sExportLine = sExportLine & """" & oField.Value & ""","
                    Else
                        sExportLine = sExportLine & oField.Value & ","
                    End If
                Next
            
                sExportLine = Left$(sExportLine, Len(sExportLine) - 1)
                
                If AsUnicode Then
                    Put #hFile, , StrConv(sExportLine, vbUnicode) & vbNewLine        ' esto es para soporte internacional, escribe puenteando la conversion a ANSI
                Else
                    Print #hFile, sExportLine
                End If
                
                .MoveNext
                
                If cont Mod 20 = 0 Then
                    DoEvents
                    fExp.progBar.Value = .AbsolutePosition
                    fExp.Show
                End If
                cont = cont + 1

            Loop
            '----Escribo los registros ----
        End If
    End With
    
    Unload fExp
    
    If (hFile <> 0) Then
        Close hFile
    End If
    
    
  '---------------------------------------------------------------------------------------------------------------
   On Error GoTo 0
   Exit Sub

exportToCSV_Error:

'   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure exportToCSV of Módulo ExcelRoutines"
    ErrorHandler Err, "exportToCSV", "Módulo", "ExcelRoutines"
    
    If (hFile <> 0) Then
        Close hFile
    End If

End Sub



'---------------------------------------------------------------------------------------
' Procedure : saveToXLS
' DateTime  : 01/04/2007 15:41
' Author    : Administrador
' Purpose   : Funcion de exportacion de recordset a excel file
'---------------------------------------------------------------------------------------
'
Function exportToXLS(mRst As ADODB.Recordset, format As saveExcelType, Optional headers As Boolean = True, Optional defaulName As String = "") As Boolean
Dim sFilter As String
Dim NomArchivo As String
Dim res As Integer
    
   On Error GoTo saveToXLS_Error
   '---------------------------------------------------------------------------------------------------------------

    sFilter = "Excel files (*.xls)|*.xls|Text files (*.txt)|*.txt|All files (*.*)|*.*"
    NomArchivo = ShowSaveFileDialog(sFilter, "xls", App.Path, OFN_OVERWRITEPROMPT, , defaulName)
   
    If NomArchivo <> "" Then
        
        logger "Exporting results to target file '" & NomArchivo & "'...."
        
        Dim rs2 As ADODB.Recordset
        Set rs2 = RS.Clone
        
        If format = XlSByAutomation Then
            CreateExcelFromRS rs2, NomArchivo
        Else
            RecordsetToHTMLXLS rs2, NomArchivo
        End If
            
        rs2.Close
    
        logger "Exporting tasks complete."
            
        saveToXLS = True
    Else
        MsgBox "The operation was canceled", vbInformation, appTitle
        saveToXLS = False
    End If
   

  '---------------------------------------------------------------------------------------------------------------
   On Error GoTo 0
   Exit Function

saveToXLS_Error:

'   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure saveToXLS of Módulo ExcelRoutines"
    ErrorHandler Err, "saveToXLS", "Módulo", "ExcelRoutines"


End Function

