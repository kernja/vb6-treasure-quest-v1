VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Object Editor"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picStats 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5040
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load Mask"
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load Sprite"
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Object"
      Height          =   615
      Left            =   4200
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Object"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   3255
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "solid"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   9
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "toggle"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "ladder"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   7
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CheckBox chkFlag 
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   6
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CheckBox chkFlag 
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   5
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CheckBox chkFlag 
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   4
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "full copy sprite"
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
   End
   Begin VB.PictureBox picPreview 
      Height          =   615
      Index           =   1
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.PictureBox picPreview 
      Height          =   615
      Index           =   0
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.ListBox lstItems 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoadSet 
         Caption         =   "Load Object Set"
      End
      Begin VB.Menu mnuSaveSet 
         Caption         =   "Save Object Set"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
