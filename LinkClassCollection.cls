VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LinkClassCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"LinkClass"
Attribute VB_Ext_KEY = "Member0" ,"LinkClass"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Public RS As ADODB.Recordset

Private mLista As ListBox

'variable local para contener colección
'Private mCol As Collection
Private orderActual As Long
'variables locales para almacenar los valores de las propiedades
Private mvarenabled As Boolean 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarcategoriaActual As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarallowDuplicates As Boolean 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarpersistent As Boolean 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarpersistentFileName As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarlastAnalizedLink As String 'copia local
Public Property Let lastAnalizedLink(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.lastAnalizedLink = 5
    mvarlastAnalizedLink = vData
End Property


Public Property Get lastAnalizedLink() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.lastAnalizedLink
    lastAnalizedLink = mvarlastAnalizedLink
End Property



Public Property Let persistentFileName(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.persistentFileName = 5
    mvarpersistentFileName = vData
End Property


Public Property Get persistentFileName() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.persistentFileName
    persistentFileName = mvarpersistentFileName
End Property




'---------------------------------------------------------------------------------------
' Procedure : setPersistentStorage
' DateTime  : 07/03/2007 03:14
' Author    : Administrador
' Purpose   : Inicializa el archivo de almacenamiento persistente en formato XML
'---------------------------------------------------------------------------------------
'
Public Sub setPersistentStorage(FileName As String)

Dim XMLFile As String
    
   On Error GoTo setPersistentStorage_Error
   '---------------------------------------------------------------------------------------------------------------
    
    persistent = True
    
    XMLFile = FileName
    
    If Not RS Is Nothing Then
            
        If Len(Dir(XMLFile)) = 0 Then
            RS.Save XMLFile, adPersistXML
        End If
           
        RS.Close
        Set RS = Nothing
            
        Set RS = New ADODB.Recordset
        RS.open XMLFile, "PROVIDER=MSPersist;", adOpenKeyset, adLockOptimistic, adCmdFile
                        
            '------ dump recordset-----
            
            Dim mRs As ADODB.Recordset
            Set mRs = RS.Clone
    
            With mRs
                
                logger "Reading data from a previous session.", 2
                logger ""
                
                If Not .EOF Then
                    
                    .MoveFirst
                    While Not .EOF
                        cont = cont + 1
                        logger "     > " & .Fields("text").Value
                        .MoveNext
                    Wend
                    
                    .MoveFirst
                    
                    lastAnalizedLink = Config("lastAnalizedLink")
                    Me.categoriaActual = mRs.Fields("text").Value
                    
                End If
                logger ""
                logger "Categories recovered from a previous session [" & .RecordCount & "].", 2
                logger ""
            
            End With
            
        persistentFileName = XMLFile
 
    End If
    
   On Error GoTo 0
   Exit Sub

   '---------------------------------------------------------------------------------------------------------------
   
setPersistentStorage_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en el procedimiento setPersistentStorage del Módulo de clase LinkClassCollection"
    'ErrorHandler Err, "setPersistentStorage", "Módulo de clase", "LinkClassCollection"
End Sub

Public Property Let persistent(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.persistent = 5
    mvarpersistent = vData
End Property


Public Property Get persistent() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.persistent
    persistent = mvarpersistent
End Property



Public Property Let allowDuplicates(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.allowDuplicates = 5
    mvarallowDuplicates = vData
End Property


Public Property Get allowDuplicates() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.allowDuplicates
    allowDuplicates = mvarallowDuplicates
End Property



Public Property Let categoriaActual(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.categoriaActual = 5
    mvarcategoriaActual = vData
End Property


Public Property Get categoriaActual() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.categoriaActual
    categoriaActual = mvarcategoriaActual
End Property

Public Sub asociarLista(ByRef List1 As ListBox)

    Set mLista = List1

End Sub

Public Sub clear()
    'vacio la coleccion de links
    While Me.Count > 0
        Me.Remove (0)
    Wend
'    RS.Close
'    RS.open
'    RS.Requery
    
    'si tiene una lista, la borro tambien
    If Not mLista Is Nothing Then
      While mLista.ListCount > 0
        mLista.RemoveItem (0)
      Wend
    End If
    
End Sub

'
'Public Sub dump()
'Dim l As LinkClass
'Dim cont As Integer
'    logger ""
'    logger "Categories collected [" & mCol.Count & "].", 3
'
'    cont = 0
'    For Each l In mCol
'        cont = cont + 1
'        logger vbTab & "(" & cont & ") {" & l.orden & "}" & l.texto & " -> " & l.checked, 3
'    Next
'
'End Sub

Public Sub dump()
Dim cont As Integer
    
    Dim mRs As ADODB.Recordset
    Set mRs = RS.Clone
    
    With mRs
        logger ""
        logger "Categories collected [" & .RecordCount & "].", 3
            
        .MoveFirst
        cont = 0
        While Not .EOF
            cont = cont + 1
            logger vbTab & "(" & cont & ") {" & .Fields("order").Value & "}" & .Fields("text").Value & " -> " & .Fields("checked").Value, 3
            
            .MoveNext
        Wend
    End With
    logger ""
    
End Sub

Public Property Let enabled(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.enabled = 5
    mvarenabled = vData
End Property


Public Property Get enabled() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.enabled
    enabled = mvarenabled
End Property


'Public Function getFirst() As LinkClass
'
'Dim link As LinkClass
'Dim order As Long
'
'      ' esta habilitado?
'    ' Lo uso para que ningun evennto asincronico intente agregar entradas a la coleccion
'    If enabled = False Then
'        Set getFirst = Nothing
'        Exit Function
'    End If
'
'    order = orderActual
'    While order >= 0
'        For Each link In mCol
'
'            ' el primer link unchecked de orden mas alto
'            If link.orden = order And link.checked = False Then
'                categoriaActual = link.texto
'                logger ""
'                logger "> > > processing order " & "[" & link.orden & "] - " & link.texto
'                link.checked = True
'                Set getFirst = link '.link
'
'                ' tildo el check en la lista de links
''                checkListItem link.texto
'
'                Exit Function
'            End If
'        Next
'        order = order - 1
'    Wend
'
'    Set getFirst = Nothing 'no se encontro un link valido que cumpliera esta condicion
'
'End Function

Public Sub checkItem(item As String)
    
    Dim tRs As ADODB.Recordset
    Set tRs = RS.Clone
    
    With tRs
        
            .Filter = "text='" & item & "'"
            '.Find "text='" & item & "'"
            If Not .EOF And Not .BOF Then
                
                .MoveFirst
                .Fields("checked").Value = True
                logger "  <<Category: " & .Fields("text").Value & " checked....>>", 3
                .Update
                
                If persistent = True Then
            '       RS.Requery
                    .Filter = "checked=false"
                    .Save
                End If
            End If
    End With

End Sub


Public Function getFirstLink(ByRef link As LinkClass) As Boolean
Dim order As Long
    
    ' esta habilitado?
    ' Lo uso para que ningun evennto asincronico intente agregar entradas a la coleccion
    If enabled = False Then
        'Set getFirst = Nothing
        getFirst = False
        Exit Function
    End If
    
    Dim tRs As ADODB.Recordset
    Set tRs = RS.Clone
    
    order = orderActual
    While order >= 0
        
        
        With tRs
        
            .Filter = "order=" & order & " AND checked=false"
            
            If Not .EOF And Not .BOF Then
                
                .MoveFirst
                loadRSValues tRs, link
                
                categoriaActual = link.texto
                logger ""
                logger "> > > processing order " & "[" & link.orden & "] - " & link.texto
                
'                link.checked = True
'                .Fields("checked").Value = True
'                .Update
'                If persistent = True Then .Save
                
                getFirstLink = True

                Exit Function
                
            End If
        End With
        
        order = order - 1
    Wend
    
    'Set getFirst = Nothing 'no se encontro un link valido que cumpliera esta condicion
    getFirstLink = False
  
End Function


Private Sub loadRSValues(mRs As ADODB.Recordset, ByRef obj As LinkClass)
'        .Fields.Append "link", adBSTR
'        .Fields.Append "text", adBSTR
'        .Fields.Append "parent", adInteger
'        .Fields.Append "checked", adBoolean
'        .Fields.Append "order", adBSTR
'        .Fields.Append "category", adBSTR

With obj
    .texto = mRs.Fields("text").Value
    .checked = mRs.Fields("checked").Value
    .link = mRs.Fields("link").Value
    .orden = mRs.Fields("order").Value
    .parent = mRs.Fields("parent").Value
    .categoria = mRs.Fields("category").Value
    
End With
End Sub

Private Function checkListItem(valorBuscado As String) As String
Dim strListbox As Object
Dim mValor As String
        
        If mLista Is Nothing Then
            checkListItem = False
            Exit Function
        End If
        
        For Each strListbox In mLista  '.Items
            mValor = Mid(strListbox, 3, Len(strListbox + 1))
            If mValor = valorBuscado Then
                strListbox.Check = True
                checkListItem = True
            Else
                checkListItem = False
            End If
        Next

End Function


Public Sub nextOrder()
    orderActual = orderActual + 1
End Sub

Public Function unprocesedItems() As Long
Dim a As LinkClass
Dim Count As Long
    Count = 0
    
    Dim tRs As ADODB.Recordset
    Set tRs = RS.Clone
    With tRs
        If Not .EOF Then .MoveFirst
        
        While Not .EOF
            If .Fields("checked").Value = False Then Count = Count + 1
            .MoveNext
        Wend
   
   End With
   unprocesedItems = Count

'    For Each a In mCol
'        If a.checked = False Then Count = Count + 1
'    Next

End Function

'Public Function linkVisited() As Boolean
'Dim obj As LinkClass
'
'    For Each obj In mCol
'        If obj.texto = texto Then
'
'            If obj.checked = True Then
'                linkVisited = True
'            Else
'                linkVisited = False
'            End If
'
'            Exit Function
'
'        End If
'    Next
'
'End Function

'Public Function Add(link As String, texto As String, parent As Long, checked As Boolean, orden As Long, Optional sKey As String) As LinkClass
'---------------------------------------------------------------------------------------
' Procedure : Add
' DateTime  : 06/03/2007 17:42
' Author    : Administrador
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Add(link As String, texto As String, parent As Long, Optional checked As Boolean, Optional sKey As String = "", Optional categoria As String) As Boolean 'LinkClass
Dim existe As Boolean
   On Error GoTo Add_Error
   '---------------------------------------------------------------------------------------------------------------
    
    existe = False ' existe??
    
    ' esta habilitado?
    ' Lo uso para que ningun evennto asincronico intente agregar entradas a la coleccion
    If enabled = False Then
        'Set
        Add = False
        Exit Function
    End If
    
    'crear un nuevo objeto
'    Dim objNewMember As LinkClass
'    Set objNewMember = New LinkClass
'
'    'establecer las propiedades que se transfieren al método
'    objNewMember.link = link
'    objNewMember.texto = texto
'    objNewMember.parent = parent
'    objNewMember.checked = checked
'    objNewMember.orden = orderActual ' orden
'    objNewMember.categoria = categoria
    
'    Dim obj As LinkClass
'    For Each obj In mCol
'        If obj.texto = texto Then
'            logger vbTab & " >>>>>> Item [" & texto & "] exists in the list. Ignored."
'            objNewMember.checked = True
'            existe = True
'            Exit For
'        End If
'    Next
    
    Dim tRs As ADODB.Recordset
    Set tRs = RS.Clone
    With tRs
        
        If Not .EOF Then .MoveFirst
        
        While Not .EOF
            If .Fields("text").Value = texto Then
                logger vbTab & " >>>>>> Item [" & texto & "] exists in the list. Ignored."
                existe = True
            End If
            .MoveNext
        Wend
   
    End With
    
        
    If (Not existe) Or (existe And allowDuplicates) Then
        logger " >>" & texto, 3 '& " - " & link,3
 
'        If Len(sKey) = 0 Then
'            mCol.Add objNewMember
'        Else
'            mCol.Add objNewMember, sKey
'        End If
'
    '-----insert into the internal recordset---
    
    If Not RS Is Nothing Then
        With RS
            
            .AddNew
            
            .Fields("link").Value = link
            .Fields("text").Value = texto
            .Fields("parent").Value = parent
            .Fields("checked").Value = checked
            .Fields("order").Value = orderActual
            .Fields("category").Value = categoria
            
            .Update
            
            If persistent = True Then .Save
            
        End With
    End If
    
    '-----insert into the internal recordset---

    
 
        'agrego a la lista si tiene alguna asociada
        If Not mLista Is Nothing Then
             If mLista.ListCount > 2500 Then mLista.RemoveItem (0)
             'mLista.AddItem "[" & objNewMember.orden & "] - " & objNewMember.texto 'objNewMember.link
             
             mLista.AddItem objNewMember.texto 'objNewMember.link
             mLista.Selected(mLista.ListCount - 1) = Not (objNewMember.checked)
             
             'frmMain.LabelLinks = "Listed links: " & mLista.ListCount
        End If
        
    End If
    
    'devolver el objeto creado
    'Set objNewMember = Nothing
    
    'Set
    Add = True 'objNewMember
    

   On Error GoTo 0
   Exit Function

   '---------------------------------------------------------------------------------------------------------------
   
Add_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") en el procedimiento Add del Módulo de clase LinkClassCollection"
    ErrorHandler Err, "Add", "Módulo de clase", "LinkClassCollection"

    'Set Add = objNewMember
    'Set objNewMember = Nothing

    Add = False
    
End Function

'Public Property Get Item(vntIndexKey As Variant) As LinkClass
'    'se usa al hacer referencia a un elemento de la colección
'    'vntIndexKey contiene el índice o la clave de la colección,
'    'por lo que se declara como un Variant
'    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
'  Set Item = mCol(vntIndexKey)
'End Property



Public Property Get Count() As Long
    'se usa al obtener el número de elementos de la
    'colección. Sintaxis: Debug.Print x.Count
    Count = RS.RecordCount ' mCol.Count
End Property


'---------------------------------------------------------------------------------------
' Procedure : Remove
' DateTime  : 07/03/2007 01:02
' Author    : Administrador
' Purpose   : Cuidado porque cambia la posicion del registro actual!!!!
'---------------------------------------------------------------------------------------
'
Public Sub Remove(vntIndexKey As Variant)
    'se usa al quitar un elemento de la colección
    'vntIndexKey contiene el índice o la clave, por lo que se
    'declara como un Variant
    'Sintaxis: x.Remove(xyz)

    With RS
        .MoveFirst
        .Move vntIndexKey ', 0
        .Delete adAffectCurrent
        .Update
        
        If persistent = True Then .Save
        '.Resync
    End With
    'mCol.Remove vntIndexKey
End Sub


'Public Property Get NewEnum() As IUnknown
'    'esta propiedad permite enumerar
'    'esta colección con la sintaxis For...Each
'    Set NewEnum = mCol.[_NewEnum]
'End Property


Private Sub Class_Initialize()
    'crea la colección cuando se crea la clase
  '  Set mCol = New Collection
    Set mLista = Nothing
    orderActual = 1
    enabled = True
    allowDuplicates = False
    persistent = False
    
    ' persistent storage
    
    Set RS = New ADODB.Recordset

    With RS
    '    .CursorLocation = adUseClient      'Enable index creation
        
        .Fields.Append "link", adBSTR
        .Fields.Append "text", adBSTR
        .Fields.Append "parent", adInteger
        .Fields.Append "checked", adBoolean
        .Fields.Append "order", adBSTR
        .Fields.Append "category", adBSTR
        
        .LockType = adLockOptimistic
        .CursorType = adOpenKeyset
        .open
    End With


End Sub


Private Sub Class_Terminate()
    'destruye la colección cuando se termina la clase
   ' Set mCol = Nothing
    'saveRStoCSV RS, App.Path & "\perStor.csv"
 '   RS.Save App.Path & "\perStor.xml", adPersistXML
    Set RS = Nothing
End Sub

'Public Property Get unprocesedItems() As Long
'Dim a As LinkClass
'Dim Count As Long
'    Count = 0
 '   For Each a In mCol
  '      If a.checked = False Then Count = Count + 1
   ' Next
'   unprocesedItems = Count
'End Property


'Public Sub nextOrder()
    'orderActual = orderActual + 1
'End Sub


'Public Property Get getFistLink() As LinkClass
   
'Dim link As LinkClass
'Dim order As Long
    
 '   order = orderActual
  '  While order >= 0
   '     For Each link In mCol
    '
            ' el primer link unchecked de orden mas alto
'            If link.orden = order And link.checked = False Then
 '               link.checked = True
  '              Set getFistLink = link '.link
   '             Exit Property
    '        End If
     '   Next
      '  order = order - 1
    'Wend
    
'    Set getFistLink = Nothing 'no se encontro un link valido que cumpliera esta condicion
  
'End Property


