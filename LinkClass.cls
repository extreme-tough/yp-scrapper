VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LinkClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales para almacenar los valores de las propiedades
Private mvarlink As String 'copia local
Private mvartexto As String 'copia local
Private mvarparent As Long 'copia local
Private mvarchecked As Boolean 'copia local
Private mvarorden As Long 'copia local
Private mvarcategoria As String 'copia local

Public Property Let categoria(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.orden = 5
    mvarcategoria = vData
End Property


Public Property Get categoria() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.orden
    categoria = mvarcategoria
End Property


Public Property Let orden(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.orden = 5
    mvarorden = vData
End Property


Public Property Get orden() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.orden
    orden = mvarorden
End Property



Public Property Let checked(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.checked = 5
    mvarchecked = vData
End Property


Public Property Get checked() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.checked
    checked = mvarchecked
End Property



Public Property Let parent(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.parent = 5
    mvarparent = vData
End Property


Public Property Get parent() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.parent
    parent = mvarparent
End Property



Public Property Let texto(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.texto = 5
    mvartexto = vData
End Property


Public Property Get texto() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.texto
    texto = mvartexto
End Property



Public Property Let link(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.link = 5
    mvarlink = vData
End Property


Public Property Get link() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.link
    link = mvarlink
End Property



