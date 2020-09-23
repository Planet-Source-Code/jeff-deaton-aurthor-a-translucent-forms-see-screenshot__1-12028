VERSION 5.00
Begin VB.Form frmTrans 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Trans Form"
   ClientHeight    =   7725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13950
   Icon            =   "frmTrans.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   13950
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picColorBox 
      AutoRedraw      =   -1  'True
      Height          =   3015
      Left            =   3420
      ScaleHeight     =   2955
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   900
      Width           =   285
   End
   Begin VB.PictureBox picColors 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3015
      Left            =   360
      Picture         =   "frmTrans.frx":08CA
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   900
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Use the colors below to select the translucent color..."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   330
      TabIndex        =   3
      Top             =   510
      Width           =   3750
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Translucent Form (Aurthor-Jeff Deaton)"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   2745
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6420
      TabIndex        =   0
      Top             =   4035
      Width           =   255
   End
   Begin VB.Image imgImages 
      Height          =   225
      Index           =   5
      Left            =   7980
      Picture         =   "frmTrans.frx":150A
      Top             =   1110
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgImages 
      Height          =   225
      Index           =   4
      Left            =   7710
      Picture         =   "frmTrans.frx":1926
      Top             =   1110
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgImages 
      Height          =   225
      Index           =   3
      Left            =   7980
      Picture         =   "frmTrans.frx":1D2A
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgImages 
      Height          =   225
      Index           =   2
      Left            =   7710
      Picture         =   "frmTrans.frx":216E
      Top             =   840
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgImages 
      Height          =   300
      Index           =   1
      Left            =   7710
      Picture         =   "frmTrans.frx":259E
      Top             =   510
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image imgImages 
      Height          =   300
      Index           =   0
      Left            =   7710
      Picture         =   "frmTrans.frx":2F6E
      Top             =   180
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image imgMinimize 
      Height          =   225
      Left            =   6960
      Picture         =   "frmTrans.frx":3710
      Top             =   60
      Width           =   225
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   7200
      Picture         =   "frmTrans.frx":3B14
      Top             =   60
      Width           =   225
   End
   Begin VB.Image imgButton 
      Height          =   300
      Left            =   5850
      Picture         =   "frmTrans.frx":3F44
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Image Image2 
      Height          =   4170
      Index           =   1
      Left            =   7440
      Picture         =   "frmTrans.frx":46E6
      Top             =   330
      Width           =   75
   End
   Begin VB.Image Image2 
      Height          =   4170
      Index           =   0
      Left            =   0
      Picture         =   "frmTrans.frx":495A
      Top             =   330
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   0
      Picture         =   "frmTrans.frx":4BE7
      Top             =   0
      Width           =   7500
   End
   Begin VB.Image Image2 
      Height          =   75
      Index           =   2
      Left            =   30
      Picture         =   "frmTrans.frx":5359
      Top             =   4440
      Width           =   7500
   End
End
Attribute VB_Name = "frmTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()

'Make the form translucent...
Call MakeTranslucent(Me, tColor)

End Sub
Private Sub Form_Load()

'Width and Height...
Me.Width = 7500
Me.Height = 4500

'Make the form translucent...
Call MakeTranslucent(Me, tColor)

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    Call DragForm(Me)
    Call MakeTranslucent(Me, tColor)
End If

End Sub
Private Sub Form_Resize()

If iRecursion Then Exit Sub

'Make the form translucent...
Call MakeTranslucent(Me, tColor)

End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    Call Form_MouseMove(vbLeftButton, Shift, X, Y)
End If

End Sub

Private Sub imgButton_Click()

lblExit_Click

End Sub
Private Sub imgButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgButton.Picture = imgImages(1).Picture
End If

End Sub
Private Sub imgButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgButton.Picture = imgImages(0).Picture

End Sub
Private Sub imgClose_Click()

Unload Me

End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgClose.Picture = imgImages(3).Picture
End If

End Sub
Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgClose.Picture = imgImages(2).Picture

End Sub

Private Sub imgMinimize_Click()

iRecursion = True
Me.WindowState = vbMinimized
iRecursion = False

End Sub
Private Sub imgMinimize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgMinimize.Picture = imgImages(5).Picture
End If

End Sub
Private Sub imgMinimize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgMinimize.Picture = imgImages(4).Picture

End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    Call DragForm(Me)
    Call MakeTranslucent(Me, tColor)
End If

End Sub
Private Sub lblExit_Click()

Unload Me

End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgButton.Picture = imgImages(1).Picture
End If

End Sub

Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgButton.Picture = imgImages(0).Picture

End Sub

Private Sub picColors_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Local Error Resume Next

If Button = vbLeftButton Then
    picColorBox.BackColor = picColors.Point(X, Y)
End If

End Sub
Private Sub picColors_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

tColor = picColors.Point(X, Y)
Call MakeTranslucent(Me, tColor)

End Sub
