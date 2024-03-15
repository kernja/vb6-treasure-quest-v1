VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mapObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public objType As String
Public itemIndex As Single

Public width As Long
Public height As Long
Public collisionEnable As Boolean
Private Sub Class_Initialize()
    objType = ""
End Sub

Private Sub Class_Terminate()
    objType = ""
End Sub

Public Sub create(objType As String)
    Me.objType = objType
    collisionEnable = True
    
        If Val(objType) = 0 Then
            frmItems.picItemSprite(itemIndex).Picture = LoadPicture(App.Path & "\items\mario\block_brick.bmp")
        ElseIf Val(objType) = 1 Then
           frmItems.picItemSprite(itemIndex).Picture = LoadPicture(App.Path & "\items\mario\block_question.bmp")
        Else
            frmItems.picItemSprite(itemIndex).Picture = LoadPicture(App.Path & "\items\mario\block_empty.bmp")
        End If
        
       frmItems.picItemMask(itemIndex).Picture = LoadPicture(App.Path & "\items\mario\full_mask.bmp")
End Sub
