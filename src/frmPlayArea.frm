VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlayArea 
   Caption         =   "Form1"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   12210
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar progressGas 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   8880
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
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
   End
End
Attribute VB_Name = "frmPlayArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    mdlMain.targetHDC = picPlayfield.hdc
        frmPlayArea.Show
            mdlMain.main picPlayfield.hdc
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdlMain.doRender = False
    End
End Sub

Private Sub picPlayfield_KeyDown(KeyCode As Integer, Shift As Integer)
'If gtcDesiredTime < GetTickCount - gtcStart Then
 
'End If
End Sub
