VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   592
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   935
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog commonDialog 
      Left            =   11640
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmTools 
      Caption         =   "Tools"
      Height          =   8655
      Left            =   12240
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton optTool 
         Caption         =   "Select Object"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "full copy sprite"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   8160
         Width           =   1455
      End
      Begin VB.CheckBox chkFlag 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   7920
         Width           =   1455
      End
      Begin VB.CheckBox chkFlag 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   7680
         Width           =   1455
      End
      Begin VB.CheckBox chkFlag 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   7440
         Width           =   1455
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "ladder"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   7200
         Width           =   1455
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "toggle"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   6960
         Width           =   1455
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "solid"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   6720
         Width           =   1455
      End
      Begin VB.ListBox lstObjects 
         Height          =   4740
         ItemData        =   "frmEditor.frx":0000
         Left            =   120
         List            =   "frmEditor.frx":0002
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton btnDelete 
         Caption         =   "Delete Object"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   6120
         Width           =   1455
      End
      Begin VB.OptionButton optTool 
         Caption         =   "Add Object"
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.Line lneBorder 
         Index           =   1
         X1              =   120
         X2              =   1560
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Line lneBorder 
         Index           =   0
         X1              =   120
         X2              =   1560
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin VB.PictureBox picPlayfield 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   8640
      Left            =   120
      ScaleHeight     =   576
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   120
      Width           =   12000
      Begin VB.Shape shapeCursor 
         BackColor       =   &H00FFFFFF&
         BorderWidth     =   3
         DrawMode        =   6  'Mask Pen Not
         FillColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   480
         Top             =   3360
         Width           =   480
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuNewMap 
         Caption         =   "New Map"
      End
      Begin VB.Menu mnuLoadMap 
         Caption         =   "Load Map"
      End
      Begin VB.Menu mnuSaveMap 
         Caption         =   "Save Map"
      End
      Begin VB.Menu mnuBorderA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuMapProperties 
      Caption         =   "Map Properties"
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gtcDesiredTime As Long
Public gtcStart As Long

Public doRender As Boolean

Public myMap As New map
Public CurX, CurY As Single

Private Sub Form_Load()
    gtcStgart = 0
    gtcDesiredTime = 10
    selectedItem = -1
    doRender = True
        'frmProperties.Top = Me.Top
        frmProperties.Left = Me.Left + Me.width
        frmProperties.Show
        'myMap.myItems.Show
        Me.Show
        render
'        myMap = New map
End Sub

Public Sub render()

     Do While doRender = True
    
            If gtcDesiredTime < GetTickCount - gtcStart Then
                gtcStart = GetTickCount
                    myMap.render
            End If
    
            DoEvents
    Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
    doRender = False
        End
End Sub

Private Sub lstObjects_Click()
    If lstObjects.ListIndex > -1 Then
        optTool(0).Value = True
    End If
End Sub

Private Sub mnuLoadMap_Click()
   Set myMap = Nothing
   Set myMap = New map
   myMap.loadMap "./map.txt"
End Sub

Private Sub mnuNewMap_Click()
   Set myMap = Nothing
   Set myMap = New map
End Sub

Private Sub mnuSaveMap_Click()
'On Error GoTo errorhandler

Dim i As Single, J As Single, FS As FileSystemObject, MF As TextStream
Set FS = CreateObject("Scripting.FileSystemObject")
    Set MF = FS.CreateTextFile("./map.txt", True)
        MF.WriteLine "//OUTPUT BLOCK DATA AND AMOUNT OF BLOCKS TO LOAD, STARTING AT -1 (NO BLOCK)"
        MF.WriteLine "//BLOCK X COORD, BLOCK Y COORD, BLOCK WIDTH, BLOCK HEIGHT, ITEM TYPE"
        MF.WriteLine Str(myMap.itemCount)
        
            For i = 0 To myMap.itemCount
                MF.WriteLine Str(myMap.getItem(i).x) & "," & Str(myMap.getItem(i).y) & "," & Str(myMap.getItem(i).width) & "," & Str(myMap.getItem(i).height) & "," & Str(myMap.getItem(i).itemType)
            Next i
        
    Exit Sub
'errorhandler:
 '   MsgBox "Error outputting HTML file.", 16, "Whoops!"
End Sub

Private Sub picPlayfield_Click()
    If optTool(0).Value = True Then
        'update selected item
    ElseIf optTool(1).Value = True Then
        If frmProperties.selectedItem > -1 Then
            
            myMap.addItem Str(frmProperties.selectedItem), shapeCursor.Left, shapeCursor.Top
        'myMap.itemCount = myMap.itemCount + 1
            lstObjects.addItem "block"
            lstObjects.ItemData(lstObjects.ListCount - 1) = myMap.itemCount
            
                'MsgBox lstObjects.ItemData(myMap.itemCount)
        End If
    End If
End Sub

Private Sub picPlayfield_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    shapeCursor.Left = Int(x / 32) * 32
    shapeCursor.Top = Int(y / 32) * 32
End Sub
