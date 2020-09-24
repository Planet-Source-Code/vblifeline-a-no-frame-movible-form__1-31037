VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   Picture         =   "noframmovibleform.frx":0000
   ScaleHeight     =   2550
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image5 
      Height          =   315
      Left            =   5100
      Top             =   75
      Width           =   390
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"noframmovibleform.frx":2EAAA
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   675
      TabIndex        =   2
      Top             =   600
      Width           =   4515
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3375
      TabIndex        =   1
      Top             =   1575
      Width           =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   825
      TabIndex        =   0
      Top             =   1575
      Width           =   1440
   End
   Begin VB.Image Image3 
      Height          =   555
      Left            =   3300
      Picture         =   "noframmovibleform.frx":2EB34
      Top             =   1425
      Width           =   1590
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   750
      Picture         =   "noframmovibleform.frx":319B6
      Top             =   1425
      Width           =   1590
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   750
      Picture         =   "noframmovibleform.frx":34838
      Top             =   1425
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Image Image4 
      Height          =   600
      Left            =   3300
      Picture         =   "noframmovibleform.frx":37A7A
      Top             =   1425
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private MouseDownForm
Private MouseDownFormX
Private MouseDownFormY
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Sub form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDownForm = 1
MouseDownFormX = X
MouseDownFormY = Y
End Sub
Private Sub form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDownForm = 0
End Sub
Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseDownForm <> 1 Then
Exit Sub
End If
Dim Z As POINTAPI
Call GetCursorPos(Z)
Form1.Top = (Z.Y * 15) - MouseDownFormY
Form1.Left = (Z.X * 15) - MouseDownFormX
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
Image1.Visible = True
End Sub
Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
Image1.Visible = False
End Sub
Private Sub Image5_Click()
End
End Sub
Private Sub Label1_Click()
MsgBox "You CLick Yes"
End Sub
Private Sub Label2_Click()
MsgBox "You CLick No"
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
Image1.Visible = True
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
Image1.Visible = False
End Sub
Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = False
Image4.Visible = True
End Sub
Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = True
Image4.Visible = False
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = False
Image4.Visible = True
End Sub
Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = True
Image4.Visible = False
End Sub
