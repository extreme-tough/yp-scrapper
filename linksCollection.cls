VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INVALIDOlinksCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mColX1 As Collection
Private orderActual As Long


Private Sub Class_Initialize()
   orderActual = 5
   Set mColX1 = New Collection
End Sub

Private Sub Class_Terminate()
    Set mColX1 = Nothing
End Sub

'Public Property Let actual(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.actual = 5
 '   mvaractual = vData
'End Property


'Public Property Get actual() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.actual
'    actual = mvaractual
'End Property

Public Property Get Item(index As Integer) As linkInfo
   Set Item = mColX1.Item(index)
End Property

Public Function ItemCount() As Long
   ItemCount = mColX1.Count
End Function

Public Sub RemoveItem(index As Long)
   mColX1.Remove index
End Sub

Public Function Add( _
                    verificado As Integer, _
                    link As String, _
                    texto As String, _
                    parent As Integer, _
                    Optional sKey As String _
                    ) As linkInfo
    
    'créer un nouvel objet
    Dim objNewMember As linkInfo
    Set objNewMember = New linkInfo

    With objNewMember
        .verificado = verificado
        .link = link
        .texto = texto
        .orden = orderActual ' orden
        .parent = parent
   End With
   
    If Len(sKey) = 0 Then
        mColX1.Add objNewMember
    Else
        mColX1.Add objNewMember, sKey
    End If


    'renvoyer l'objet créé
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Sub AddItem(linkInfo As linkInfo)
    linkInfo.orden = orderActual
    mColX1.Add linkInfo
End Sub

'Public Property Get NewEnum() As IUnknown
'   Set NewEnum = mColX1.[_NewEnum]
'End Property

Public Property Get unprocesedItems() As Long
Dim a As linkInfo
Dim Count As Long
    Count = 0
    For Each a In mColX1
        If a.verificado = False Then Count = Count + 1
    Next
   unprocesedItems = Count
End Property

Public Sub nextOrder()
    orderActual = orderActual + 1
End Sub


Public Property Get getFistLink() As linkInfo
   
Dim link As linkInfo
Dim order As Long
    
    order = orderActual
    While order >= 0
        For Each link In mColX1
        
            ' el primer link unchecked de orden mas alto
            If link.orden = order And link.verificado = False Then
                link.verificado = True
                Set getFistLink = link '.link
                Exit Sub
            End If
        Next
        order = order - 1
    Wend
    
    Set getFistLink = Nothing 'no se encontro un link valido que cumpliera esta condicion
  
End Property

