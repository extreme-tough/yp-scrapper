VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "linkInfoINVALIDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'variables locales para almacenar los valores de las propiedades
Private mvarverificado As Integer 'copia local
Private mvarlink As String 'copia local
Private mvartexto As String 'copia local
Private mvarorden As Integer 'copia local
Private mvarparent As Integer 'copia local

Public Property Let parent(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.parent = 5
    mvarparent = vData
End Property


Public Property Get parent() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.parent
    parent = mvarparent
End Property



Public Property Let orden(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.orden = 5
    mvarorden = vData
End Property


Public Property Get orden() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.orden
    orden = mvarorden
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



Public Property Let verificado(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.verificado = 5
    mvarverificado = vData
End Property


Public Property Get verificado() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.verificado
    verificado = mvarverificado
End Property



