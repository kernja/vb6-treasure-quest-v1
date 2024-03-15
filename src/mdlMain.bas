Attribute VB_Name = "mdlMain"
Public Declare Function StretchBlt Lib "gdi32" _
     (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
     ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, _
     ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'**************************************
'Windows API/Global Declarations for :Fo
'     rmat GetTickCount
'**************************************

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
          
   'You'll also need these
   Public Const KEY_TOGGLED As Integer = &H1
   Public Const KEY_PRESSED As Integer = &H1000
   
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Type rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function IntersectRect Lib "user32" (lpDestRect As rect, lpSrc1Rect As rect, lpSrc2Rect As rect) As Long
Public targetHDC

Public myMap As map
Public doRender As Boolean

Public gtcDesiredTime As Single
Public gtcStart As Single


Public Sub Main(newHDC)
    Set myMap = New map
        doRender = False
        'targetHDC = newHDC
        myMap.loadMap ("./map.txt")
            doRender = True
            
        gtcDesiredTime = 10
        gtcStart = 0
        render
        
End Sub
Public Sub render()
     Do While doRender = True
    
            If gtcDesiredTime < GetTickCount - gtcStart Then
                gtcStart = GetTickCount
                    getInput
                    myMap.renderToOtherHDC
            End If
    
            DoEvents
    Loop

End Sub

Sub getInput()
    'If GetKeyState(vbKeyA) And KEY_PRESSED Then
    '    myMap.moveCharacter 0, -8, False
    'ElseIf GetKeyState(vbKeyB) And KEY_PRESSED Then
    '    myMap.moveCharacter 0, 8, False
    'End If
    If GetKeyState(vbKeyRight) And KEY_PRESSED Then
        myMap.moveCharacter 2, 0, False
    ElseIf GetKeyState(vbKeyLeft) And KEY_PRESSED Then
        myMap.moveCharacter -2, 0, False
    End If
    
    If mdlMain.myMap.onFloor = True Then
        If GetKeyState(vbKeyShift) And KEY_PRESSED Then
            myMap.character.gravityTime = 0
            myMap.character.upwardsVelocity = -12
        End If
    End If
    
    If GetKeyState(vbKeyF1) And KEY_PRESSED Then
        myMap.character.gasTank = 100
    End If
    
    If GetKeyState(vbKeyZ) And KEY_PRESSED And myMap.character.gasTank > 0 Then
        myMap.character.gravityTime = 0
        myMap.character.upwardsVelocity = -5
            myMap.character.gasTank = myMap.character.gasTank - 0.5
    End If
End Sub
