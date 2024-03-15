VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mapItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public itemType As String
Public itemIndex As Single
Public width As Long
Public height As Long

Public enabled As Boolean

Public x As Single
Public y As Single

Public collisionEnable As Boolean
Public isLadder As Boolean

Private myRect As rect
Private Sub Class_Initialize()
    x = 0
    y = 0
        enabled = False
            isLadder = False
            collisionEnable = True
End Sub
Sub SetType(myType As String, myIndex As Single)
    itemType = myType
    itemIndex = myIndex
End Sub

Sub SetPosition(posX As Single, posY As Single)
    x = posX
    y = posY
    
          myRect.Top = y
            myRect.Left = x
            myRect.Right = x + width
            myRect.Bottom = y + height
End Sub

