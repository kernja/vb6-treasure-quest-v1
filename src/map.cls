VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mapItems() As mapItem
Public itemCount As Single

Public objectCount As Single
Private objectNames() As String
Public myItems As New frmItems
Private renderReady As Boolean
Public character As clsPlayer

Private Sub Class_Initialize()
Dim i As Single
    Me.itemCount = -1
        ReDim objectNames(0 To 24)
    
        objectNames(0) = "0"
        objectNames(1) = "1"
        objectNames(2) = "2"
        objectNames(3) = "3"
        objectNames(4) = "4"
        objectNames(5) = "5"
        objectNames(6) = "6"
        objectCount = 6
            
            
            For i = 1 To objectCount
                Load myItems.picItemMask(i)
                  '  myItems.picItemMask(i).Top = (myItems.picItemMask(i).Top + (32 * i))
                Load myItems.picItemSprite(i)
                  '  myItems.picItemSprite(i).Top = (myItems.picItemSprite(i).Top + (32 * i))
            Next i
            
            'Load myItems.picItemMask(70)
            '      Load myItems.picItemSprite(70)
            
                  myItems.picItemSprite(0).Picture = LoadPicture(App.path & "\items\mario\block_brick.bmp")
                    myItems.picItemMask(0).Picture = LoadPicture(App.path & "\items\mario\full_mask.bmp")
                  myItems.picItemSprite(1).Picture = LoadPicture(App.path & "\items\mario\block_question.bmp")
                    myItems.picItemMask(1).Picture = LoadPicture(App.path & "\items\mario\full_mask.bmp")
                  myItems.picItemSprite(2).Picture = LoadPicture(App.path & "\items\mario\block_empty.bmp")
                    myItems.picItemMask(2).Picture = LoadPicture(App.path & "\items\mario\full_mask.bmp")
                    
                    myItems.picItemSprite(3).Picture = LoadPicture(App.path & "\items\gems\blue_gem_sprite.bmp")
                      myItems.picItemMask(3).Picture = LoadPicture(App.path & "\items\gems\blue_gem_mask.bmp")

                    myItems.picItemSprite(4).Picture = LoadPicture(App.path & "\items\gems\no_gem_sprite.bmp")
                      myItems.picItemMask(4).Picture = LoadPicture(App.path & "\items\gems\no_gem_mask.bmp")

                    myItems.picItemSprite(5).Picture = LoadPicture(App.path & "\items\objects\key_sprite.bmp")
                      myItems.picItemMask(5).Picture = LoadPicture(App.path & "\items\objects\key_mask.bmp")

                    myItems.picItemSprite(6).Picture = LoadPicture(App.path & "\items\objects\hole_sprite.bmp")
                      myItems.picItemMask(6).Picture = LoadPicture(App.path & "\items\objects\hole_mask.bmp")
      Set character = New clsPlayer
      renderReady = True
End Sub
Public Sub addItem(objectType As Single, posX As Single, posY As Single)
Dim tempItem As mapItem

    renderReady = False
    For i = 0 To objectCount
        If Str$(objectNames(i)) = Str$(objectType) Then
            Set tempItem = New mapItem
            tempItem.itemType = objectType
            tempItem.setPosition posX, posY
            tempItem.enabled = True
            tempItem.width = 32
            tempItem.height = 32
            tempItem.itemIndex = i
            
                If Val(tempItem.itemType) <= 2 Then
                    tempItem.collisionEnable = True
                ElseIf Val(tempItem.itemType) = 3 Then
                    tempItem.collisionEnable = False
                End If
                
                If itemCount = -1 Then
                    ReDim mapItems(0 To 0)
                        Set mapItems(0) = tempItem
                Else
                    ReDim Preserve mapItems(0 To itemCount + 1)
                        Set mapItems(itemCount + 1) = tempItem
                End If
                
            Me.itemCount = Me.itemCount + 1
        End If
    Next i
    
    renderReady = True
End Sub
Sub render()
    Dim i As Single
    
    If renderReady = True Then
        frmEditor.picPlayfield.Cls
        For i = 0 To Me.itemCount
            
            With mapItems(i)
                If .itemType <= 2 Then
                If .enabled = True Then
                    StretchBlt frmEditor.picPlayfield.hdc, .x, .y, 32, 32, myItems.picItemSprite(.itemIndex).hdc, 0, 0, 32, 32, vbSrcCopy
                End If
                Else
                    StretchBlt frmEditor.picPlayfield.hdc, .x, .y, 16, 32, myItems.picItemMask(.itemIndex).hdc, 0, 0, 16, 32, vbSrcAnd
                    StretchBlt frmEditor.picPlayfield.hdc, .x, .y, 16, 32, myItems.picItemSprite(.itemIndex).hdc, 0, 0, 16, 32, vbSrcPaint
                End If
                
            End With
        Next i
        
        frmEditor.picPlayfield.Refresh
    End If
End Sub
Sub renderToOtherHDC()
    Dim i As Single
    
    If renderReady = True Then
        frmPlayArea.picPlayfield.Cls
        moveCharacter 0, 0, True
        For i = 0 To itemCount
            With mapItems(i)
            If .enabled = True Then
                StretchBlt frmPlayArea.picPlayfield.hdc, .x, .y, 32, 32, myItems.picItemSprite(.itemIndex).hdc, 0, 0, 32, 32, vbSrcCopy
            End If
            End With
        Next i
        
        'draw the player
        With character
            If .gasTank < 0 Then .gasTank = 0
            If .gasTank > 100 Then .gasTank = 100
        frmPlayArea.progressGas.Value = .gasTank
        StretchBlt frmPlayArea.picPlayfield.hdc, .x, .y, 32, 32, myItems.picItemSprite(2).hdc, 0, 0, 32, 32, vbSrcCopy
        End With
        frmPlayArea.picPlayfield.Refresh
    End If
End Sub
Private Sub Class_Terminate()
    'destroy myItems
        Erase mapItems
        itemCount = -1
        
        Set myItems = Nothing
        
End Sub
Public Sub loadMap(path As String)

Dim i As Single, J As Single, FS As FileSystemObject, MF As TextStream
Dim myLine As String
Dim mySplitResult() As String
Dim tempCount As Integer

renderReady = False
Set FS = CreateObject("Scripting.FileSystemObject")
    Set MF = FS.OpenTextFile("./map.txt")
        MF.SkipLine
        MF.SkipLine
            tempCount = Val(MF.ReadLine)
        
            For i = 0 To tempCount
                myLine = MF.ReadLine 'Str(myMap.getItem(I).x) & "," & Str(myMap.getItem(I).y) & "," & Str(myMap.getItem(I).width) & "," & Str(myMap.getItem(I).height) & "," & Str(myMap.getItem(I).itemType)
                    mySplitResult = Split(myLine, ",")
                    
                        addItem Val(mySplitResult(4)), Val(mySplitResult(0)), Val(mySplitResult(1))
            Next i
    renderReady = True
End Sub
Public Function getItem(Index As Single) As mapItem
    Set getItem = mapItems(Index)
End Function

Function isCollision(ByVal obj As mapItem, Optional shiftX As Single = 0, Optional shiftY As Single = 0) As Boolean
    Dim newRect As rect
    Dim targetRect As rect
        
        targetRect.Top = obj.y
        targetRect.Bottom = obj.y + obj.height
        targetRect.Left = obj.x
        targetRect.Right = obj.x + obj.width
        
    Dim colResult As rect
            newRect.Left = character.x + shiftX
            newRect.Right = character.x + shiftX + character.width
            newRect.Top = character.y + shiftY
            newRect.Bottom = character.y + shiftY + character.height
            
                
    
    If IntersectRect(colResult, newRect, targetRect) = 0 Then
        isCollision = False
    Else
        isCollision = True
    End If
End Function
Function onFloor() As Boolean
    Dim newRect As rect
    Dim targetRect As rect
    Dim collision As Boolean
    
    If (character.y = 544) Then
        onFloor = True
        Exit Function
    End If
    
    collision = False
       For i = 0 To UBound(mapItems)
            If collision = False Then
                collision = isCollision(mapItems(i), 0, 1)
            End If
        Next i
    onFloor = collision
End Function
Sub moveCharacter(difX As Single, difY As Single, gravity As Boolean)
Dim i As Single
Dim collision As Boolean

    collision = False
    upwardcollision = True
    
    If gravity = True Then
        If character.upwardsVelocity < 0 Then
            character.upwardsVelocity = character.upwardsVelocity + (1 * (character.gravityTime ^ 2))
                difY = difY + character.upwardsVelocity
                 character.gravityTime = character.gravityTime + (GetTickCount * 0.000000004)
                 
                    If character.upwardsVelocity > 0 Then
                        character.upwardsVelocity = 0
                       ' character.gravityTime = 0.001
                    End If
        Else
            difY = 1 * (character.gravityTime ^ 2)
        End If
    End If
    
    For i = 0 To UBound(mapItems) - 1
        If collision = False Then
            collision = isCollision(mapItems(i), difX, difY)
        End If
    Next i
        
          
    If collision = False Then
        With character
            .setPosition x + difX, y + difY
        End With
        If gravity = True Then
            If character.upwardsVelocity >= 0 Then
                character.gravityTime = character.gravityTime + (GetTickCount * 0.000000004)
                    character.upwardsVelocity = 0
            End If
        End If
    ElseIf collision = True Then
        If character.upwardsVelocity < 0 And (difX = 0) Then
            character.upwardsVelocity = 0
            difY = 0
                 character.setPosition 0, 1
        End If
    End If
        
    If gravity = True Then
        If collision = True Then
            upwardcollision = False
            character.gravityTime = 0.001
            For i = 0 To UBound(mapItems) - 1
                upwardcollision = isCollision(mapItems(i), difX, difY)
                If upwardcollision = True Then character.y = mapItems(i).y - 32
            Next i
            
            With character
                .setPosition x + difX, y + difY
            End With
        End If
    
        If (onFloor() = True) Then
            With character
                .gravityTime = 0
                .setPosition .x, Int(.y / 32) * 32, True
            End With
        End If
    End If
End Sub


