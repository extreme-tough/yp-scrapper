VERSION 5.00
Begin VB.Form frmSelect 
   Caption         =   "Select the values...."
   ClientHeight    =   5115
   ClientLeft      =   18960
   ClientTop       =   11505
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "U&nselect All"
      Height          =   465
      Left            =   4995
      TabIndex        =   4
      Top             =   4455
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S&elect All"
      Height          =   465
      Left            =   4995
      TabIndex        =   3
      Top             =   3870
      Width           =   1050
   End
   Begin VB.ListBox Lista 
      Height          =   4785
      Left            =   135
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   135
      Width           =   4650
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   645
      Left            =   4995
      TabIndex        =   1
      Top             =   900
      Width           =   1050
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "Ok"
      Height          =   645
      Left            =   4995
      TabIndex        =   0
      Top             =   135
      Width           =   1050
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mRst As ADODB.Recordset

Dim listCol As String
Dim checkCol As String

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnOK_Click()
Dim i As Integer

With mRst
    If Not .EOF Or Not .BOF Then .MoveFirst
    
    For i = 0 To Lista.ListCount - 1
        Debug.Print i & "  - " & .AbsolutePosition
        .Fields(checkCol).Value = Lista.Selected(i)
        If Not .EOF Then .Move 1
    Next
    .Requery
End With
Unload Me
End Sub

Private Sub Form_Load()

'    Set rst = New ADODB.Recordset
'    With rst
'        .CursorLocation = adUseClient
'        .CursorType = adOpenDynamic
'        .LockType = adLockOptimistic
'        .open "select *, texto + ' (' + state +')' as msg from statesUSA", Config("connectionstring")
'
'        If Not .EOF And Not .BOF Then .MoveFirst
'        While Not .EOF
'            With Lista
'            .AddItem rst.Fields("msg")
'            .Selected(.ListCount - 1) = rst.Fields("parsed").Value
'            rst.MoveNext
'            End With
'        Wend
'    End With
'
'
'    With DataGrid1
'        Set .DataSource = rst
'        '.Columns(2).DataFormat
'    End With
   
End Sub


Public Sub setData(miRst As ADODB.Recordset, listColumn As String, checkColumn As String)
Set mRst = miRst

    listCol = listColumn
    checkCol = checkColumn
    
    With mRst
        If Not .EOF And Not .BOF Then .MoveFirst
        While Not .EOF
            With Lista
            .AddItem mRst.Fields(listCol)
            .Selected(.ListCount - 1) = mRst.Fields(checkCol).Value
            mRst.MoveNext
            End With
        Wend
   End With

End Sub


Sub setGrid(allowCheckCol As String, caption As String)
    ' Variable for Format
    Dim stdYesNo As StdFormat.StdDataFormat
   
    'DataGridFormat
    Set stdYesNo = New StdFormat.StdDataFormat
    stdYesNo.Type = fmtCheckbox ' fmtBoolean
'    stdYesNo.TrueValue = "YES" 'display value for True ( -1)
'    stdYesNo.FalseValue = "NO" 'display value for False (0)
'
    stdYesNo.TrueValue = True
    stdYesNo.FalseValue = False
    
    
    With DataGrid1.Columns(allowCheckCol)
        Set .DataFormat = stdYesNo
        .Locked = False '.Columns(2).DataFormat
        .caption = caption
        '.Refresh
    End With

    DataGrid1.Refresh
End Sub

Private Sub Lista_ItemCheck(Item As Integer)
    'Lista.Selected(Item) = Not Lista.Selected(Item)

End Sub

Private Sub Command1_Click()
    Dim i As Integer
    For i = 0 To Lista.ListCount - 1
        Lista.Selected(i) = True
    Next
        
    'Lista.ListIndex(Lista.ListCount - 1) = True

End Sub

Private Sub Command2_Click()
   Dim i As Integer
   For i = 0 To Lista.ListCount - 1
        Lista.Selected(i) = False
    Next

End Sub

