VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public x As Single
Public y As Single

Public width As Long
Public height As Long

Public gasTank As Single
Public enabled As Boolean
Public gravityTime As Single
Public upwardsVelocity As Single
Private myRect As rect

Private Sub Class_Initialize()
    x = 0
    y = 0
    gravityTime = 0
    upwardsVelocity = 0
    width = 32
    height = 32
    gasTank = 100
End Sub

Sub setPosition(posX As Single, posY As Single, Optional absolute As Boolean = False)
    x = x + posX
    y = y + posY

    If (absolute = True) Then
        x = posX
        y = posY
    End If
    
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    
    If x > 768 Then x = 768
    If y > 544 Then y = 544
      myRect.Top = y
        myRect.Left = x
        myRect.Right = x + width
        myRect.Bottom = y + height
End Sub
