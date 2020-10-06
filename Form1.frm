VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menghindari Input Karakter Tertentu"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1440
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim sTemplate As String
  'Ganti '!@#$%^&*()_+=' dengan karakter yang Anda
  'inginkan untuk dihindari diinput pada Text1
  sTemplate = "!@#$%^&*()_+="
If InStr(1, sTemplate, Chr(KeyAscii)) > 0 Then _
   KeyAscii = 0
End Sub


