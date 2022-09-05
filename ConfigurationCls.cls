VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigurationCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CfgCls"
Attribute VB_Ext_KEY = "Member0" ,"CfgCls"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variable local para contener colecci?n
Private mCol As Collection
Private objIni As iniFile

Public Sub setIniFile(archivo As String)
    objIni.iniFile = archivo
End Sub

Public Sub getFromArchive()
Dim Key As CfgCls
Dim iniValue As String

    For Each Key In mCol
        Key.value = objIni.GetFile(Key.Section, Key.Key, Key.defaultValue)
    Next

End Sub

Public Sub saveToArchive()
Dim Key As CfgCls
Dim iniValue As String

    For Each Key In mCol
         objIni.WriteFile Key.Section, Key.Key, Key.value
    Next

End Sub


Public Function Add(Key As String, Section As String, value As String) As CfgCls
    'crear un nuevo objeto
    Dim objNewMember As CfgCls
    Set objNewMember = New CfgCls


    'establecer las propiedades que se transfieren al m?todo
    objNewMember.Key = Key
    objNewMember.Section = Section
    objNewMember.defaultValue = value
    
    mCol.Add objNewMember, Key
   

    'devolver el objeto creado
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As String 'CfgCls
Attribute Item.VB_UserMemId = 0
    'se usa al hacer referencia a un elemento de la colecci?n
    'vntIndexKey contiene el ?ndice o la clave de la colecci?n,
    'por lo que se declara como un Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  
  'Set Item = mCol(vntIndexKey)
  
   Item = mCol(vntIndexKey).value
End Property

Public Property Let Item(vntIndexKey As Variant, value As String)
    
'   If value = True Then value = 1
'   If value = False Then value = 0
   
   mCol(vntIndexKey).value = value
End Property



Public Property Get Count() As Long
    'se usa al obtener el n?mero de elementos de la
    'colecci?n. Sintaxis: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'se usa al quitar un elemento de la colecci?n
    'vntIndexKey contiene el ?ndice o la clave, por lo que se
    'declara como un Variant
    'Sintaxis: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'esta propiedad permite enumerar
    'esta colecci?n con la sintaxis For...Each
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'crea la colecci?n cuando se crea la clase
    Set mCol = New Collection
    Set objIni = New iniFile
End Sub


Private Sub Class_Terminate()
    'destruye la colecci?n cuando se termina la clase
    Set mCol = Nothing
    Set objIni = Nothing
End Sub

