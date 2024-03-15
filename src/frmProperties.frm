VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmOptions 
      Caption         =   "Map Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3735
      Begin VB.ListBox lstMusic 
         Height          =   645
         Left            =   1560
         TabIndex        =   12
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtRGB 
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   10
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtRGB 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtRGB 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtMapName 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label lblInfo 
         Caption         =   "Default BG Music:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label lblInfo 
         Caption         =   "Map BG Color (RGB, 0-255)"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label lblInfo 
         Caption         =   "Map Name"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame frmItems 
      Caption         =   "Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   3735
      Begin VB.OptionButton optItem 
         Height          =   495
         Index           =   6
         Left            =   2520
         Picture         =   "frmProperties.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optItem 
         Height          =   495
         Index           =   5
         Left            =   2040
         Picture         =   "frmProperties.frx":0C42
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optItem 
         Height          =   495
         Index           =   4
         Left            =   1800
         Picture         =   "frmProperties.frx":1884
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optItem 
         Height          =   495
         Index           =   3
         Left            =   1560
         Picture         =   "frmProperties.frx":1EC6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optItem 
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "frmProperties.frx":2508
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optItem 
         Height          =   495
         Index           =   1
         Left            =   600
         Picture         =   "frmProperties.frx":314A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optItem 
         Height          =   495
         Index           =   2
         Left            =   1080
         Picture         =   "frmProperties.frx":3D8C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public selectedItem As String
Private Sub Form_Load()
    selectedItem = -1
End Sub

Private Sub optItem_Click(Index As Integer)
   Select Case Index
        Case Index = 0
            selectedItem = "marioA"
        Case Index = 1
            selectedItem = "marioB"
        Case Index = 2
            selectedItem = "marioC"
        Case Index = 3
            selectedItem = ""
   End Select
    If Index <= 2 Or Index > 3 Then
        frmEditor.shapeCursor.width = 32
        frmEditor.shapeCursor.height = 32
    ElseIf Index = 3 Then
        frmEditor.shapeCursor.width = 16
        frmEditor.shapeCursor.height = 32
    End If
End Sub
