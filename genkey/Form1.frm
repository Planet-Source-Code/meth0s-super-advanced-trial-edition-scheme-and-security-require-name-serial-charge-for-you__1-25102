VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Security Generater key."
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gen Key"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
x = 0
Do While x < 256
Number = Int(255 * Rnd + 1)
Text1.Text = Text1.Text & "If currentchar = Chr(" & x & ") Then currentchar = Chr(" & Number & ") & chr("
Number = Int(255 * Rnd + 1)
Text1.Text = Text1.Text & Number & ")" & vbCrLf
x = x + 1
DoEvents
Loop
End Sub

Private Sub Form_Load()
MsgBox "Coded by meth0s in vb 6" & vbCrLf & _
"Just push Gen Key button on the form. and it will create" _
& vbCrLf & "a completly random key code. place the code in your form for conversion!", vbCritical, "Hey!"
Randomize
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Text1_Change()

End Sub
