VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mengklik Menu Start Button Windows"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Klik Start"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Const MENU_KEYCODE = 91
  keybd_event MENU_KEYCODE, 0, 0, 0
  keybd_event MENU_KEYCODE, 0, KEYEVENTF_KEYUP, 0
End Sub

