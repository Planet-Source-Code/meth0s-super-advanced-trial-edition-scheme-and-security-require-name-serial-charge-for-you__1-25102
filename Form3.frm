VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Registration"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   LinkTopic       =   "Form3"
   ScaleHeight     =   1575
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "End"
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Register"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4320
      TabIndex        =   0
      Top             =   70
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Serial:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   870
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   500
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run -X- Registration. - Give me your money!"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   100
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   0
      X2              =   4680
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      Height          =   1550
      Left            =   15
      Top             =   15
      Width           =   4665
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form3.Hide
Form1.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Text1.Text = "" Then Exit Sub
If Text2.Text = "" Then Exit Sub

    Text1.Tag = "" 'clears tag Each click
    For i = 1 To Len(Text1)
        currentchar = Mid(Text1, i, 1)
If currentchar = Chr(0) Then currentchar = Chr(56) & Chr(207)
If currentchar = Chr(1) Then currentchar = Chr(81) & Chr(100)
If currentchar = Chr(2) Then currentchar = Chr(17) & Chr(76)
If currentchar = Chr(3) Then currentchar = Chr(3) & Chr(22)
If currentchar = Chr(4) Then currentchar = Chr(232) & Chr(63)
If currentchar = Chr(5) Then currentchar = Chr(1) & Chr(218)
If currentchar = Chr(6) Then currentchar = Chr(157) & Chr(91)
If currentchar = Chr(7) Then currentchar = Chr(63) & Chr(167)
If currentchar = Chr(8) Then currentchar = Chr(219) & Chr(233)
If currentchar = Chr(9) Then currentchar = Chr(240) & Chr(180)
If currentchar = Chr(10) Then currentchar = Chr(147) & Chr(111)
If currentchar = Chr(11) Then currentchar = Chr(158) & Chr(56)
If currentchar = Chr(12) Then currentchar = Chr(169) & Chr(138)
If currentchar = Chr(13) Then currentchar = Chr(118) & Chr(115)
If currentchar = Chr(14) Then currentchar = Chr(236) & Chr(29)
If currentchar = Chr(15) Then currentchar = Chr(229) & Chr(241)
If currentchar = Chr(16) Then currentchar = Chr(76) & Chr(173)
If currentchar = Chr(17) Then currentchar = Chr(104) & Chr(14)
If currentchar = Chr(18) Then currentchar = Chr(72) & Chr(11)
If currentchar = Chr(19) Then currentchar = Chr(159) & Chr(181)
If currentchar = Chr(20) Then currentchar = Chr(165) & Chr(51)
If currentchar = Chr(21) Then currentchar = Chr(164) & Chr(110)
If currentchar = Chr(22) Then currentchar = Chr(101) & Chr(116)
If currentchar = Chr(23) Then currentchar = Chr(156) & Chr(152)
If currentchar = Chr(24) Then currentchar = Chr(247) & Chr(209)
If currentchar = Chr(25) Then currentchar = Chr(140) & Chr(242)
If currentchar = Chr(26) Then currentchar = Chr(171) & Chr(167)
If currentchar = Chr(27) Then currentchar = Chr(109) & Chr(97)
If currentchar = Chr(28) Then currentchar = Chr(106) & Chr(145)
If currentchar = Chr(29) Then currentchar = Chr(130) & Chr(236)
If currentchar = Chr(30) Then currentchar = Chr(164) & Chr(137)
If currentchar = Chr(31) Then currentchar = Chr(233) & Chr(12)
If currentchar = Chr(32) Then currentchar = Chr(137) & Chr(80)
If currentchar = Chr(33) Then currentchar = Chr(113) & Chr(35)
If currentchar = Chr(34) Then currentchar = Chr(124) & Chr(15)
If currentchar = Chr(35) Then currentchar = Chr(171) & Chr(74)
If currentchar = Chr(36) Then currentchar = Chr(196) & Chr(62)
If currentchar = Chr(37) Then currentchar = Chr(68) & Chr(80)
If currentchar = Chr(38) Then currentchar = Chr(134) & Chr(197)
If currentchar = Chr(39) Then currentchar = Chr(145) & Chr(254)
If currentchar = Chr(40) Then currentchar = Chr(239) & Chr(97)
If currentchar = Chr(41) Then currentchar = Chr(108) & Chr(160)
If currentchar = Chr(42) Then currentchar = Chr(182) & Chr(74)
If currentchar = Chr(43) Then currentchar = Chr(63) & Chr(194)
If currentchar = Chr(44) Then currentchar = Chr(194) & Chr(19)
If currentchar = Chr(45) Then currentchar = Chr(94) & Chr(56)
If currentchar = Chr(46) Then currentchar = Chr(40) & Chr(208)
If currentchar = Chr(47) Then currentchar = Chr(154) & Chr(97)
If currentchar = Chr(48) Then currentchar = Chr(88) & Chr(127)
If currentchar = Chr(49) Then currentchar = Chr(19) & Chr(171)
If currentchar = Chr(50) Then currentchar = Chr(151) & Chr(161)
If currentchar = Chr(51) Then currentchar = Chr(77) & Chr(90)
If currentchar = Chr(52) Then currentchar = Chr(175) & Chr(40)
If currentchar = Chr(53) Then currentchar = Chr(134) & Chr(134)
If currentchar = Chr(54) Then currentchar = Chr(230) & Chr(150)
If currentchar = Chr(55) Then currentchar = Chr(71) & Chr(99)
If currentchar = Chr(56) Then currentchar = Chr(46) & Chr(98)
If currentchar = Chr(57) Then currentchar = Chr(58) & Chr(197)
If currentchar = Chr(58) Then currentchar = Chr(155) & Chr(158)
If currentchar = Chr(59) Then currentchar = Chr(59) & Chr(227)
If currentchar = Chr(60) Then currentchar = Chr(26) & Chr(215)
If currentchar = Chr(61) Then currentchar = Chr(177) & Chr(92)
If currentchar = Chr(62) Then currentchar = Chr(95) & Chr(61)
If currentchar = Chr(63) Then currentchar = Chr(31) & Chr(120)
If currentchar = Chr(64) Then currentchar = Chr(31) & Chr(4)
If currentchar = Chr(65) Then currentchar = Chr(245) & Chr(175)
If currentchar = Chr(66) Then currentchar = Chr(128) & Chr(10)
If currentchar = Chr(67) Then currentchar = Chr(173) & Chr(111)
If currentchar = Chr(68) Then currentchar = Chr(205) & Chr(184)
If currentchar = Chr(69) Then currentchar = Chr(21) & Chr(26)
If currentchar = Chr(70) Then currentchar = Chr(111) & Chr(47)
If currentchar = Chr(71) Then currentchar = Chr(229) & Chr(79)
If currentchar = Chr(72) Then currentchar = Chr(25) & Chr(156)
If currentchar = Chr(73) Then currentchar = Chr(157) & Chr(106)
If currentchar = Chr(74) Then currentchar = Chr(66) & Chr(235)
If currentchar = Chr(75) Then currentchar = Chr(138) & Chr(78)
If currentchar = Chr(76) Then currentchar = Chr(216) & Chr(166)
If currentchar = Chr(77) Then currentchar = Chr(37) & Chr(98)
If currentchar = Chr(78) Then currentchar = Chr(51) & Chr(21)
If currentchar = Chr(79) Then currentchar = Chr(160) & Chr(217)
If currentchar = Chr(80) Then currentchar = Chr(72) & Chr(164)
If currentchar = Chr(81) Then currentchar = Chr(195) & Chr(55)
If currentchar = Chr(82) Then currentchar = Chr(31) & Chr(143)
If currentchar = Chr(83) Then currentchar = Chr(244) & Chr(17)
If currentchar = Chr(84) Then currentchar = Chr(134) & Chr(184)
If currentchar = Chr(85) Then currentchar = Chr(150) & Chr(20)
If currentchar = Chr(86) Then currentchar = Chr(8) & Chr(215)
If currentchar = Chr(87) Then currentchar = Chr(148) & Chr(72)
If currentchar = Chr(88) Then currentchar = Chr(25) & Chr(214)
If currentchar = Chr(89) Then currentchar = Chr(62) & Chr(150)
If currentchar = Chr(90) Then currentchar = Chr(147) & Chr(123)
If currentchar = Chr(91) Then currentchar = Chr(83) & Chr(136)
If currentchar = Chr(92) Then currentchar = Chr(102) & Chr(72)
If currentchar = Chr(93) Then currentchar = Chr(95) & Chr(81)
If currentchar = Chr(94) Then currentchar = Chr(138) & Chr(159)
If currentchar = Chr(95) Then currentchar = Chr(71) & Chr(13)
If currentchar = Chr(96) Then currentchar = Chr(57) & Chr(43)
If currentchar = Chr(97) Then currentchar = Chr(34) & Chr(73)
If currentchar = Chr(98) Then currentchar = Chr(92) & Chr(122)
If currentchar = Chr(99) Then currentchar = Chr(74) & Chr(197)
If currentchar = Chr(100) Then currentchar = Chr(67) & Chr(238)
If currentchar = Chr(101) Then currentchar = Chr(179) & Chr(122)
If currentchar = Chr(102) Then currentchar = Chr(152) & Chr(215)
If currentchar = Chr(103) Then currentchar = Chr(124) & Chr(215)
If currentchar = Chr(104) Then currentchar = Chr(150) & Chr(217)
If currentchar = Chr(105) Then currentchar = Chr(196) & Chr(81)
If currentchar = Chr(106) Then currentchar = Chr(118) & Chr(149)
If currentchar = Chr(107) Then currentchar = Chr(191) & Chr(27)
If currentchar = Chr(108) Then currentchar = Chr(42) & Chr(133)
If currentchar = Chr(109) Then currentchar = Chr(9) & Chr(49)
If currentchar = Chr(110) Then currentchar = Chr(78) & Chr(39)
If currentchar = Chr(111) Then currentchar = Chr(58) & Chr(154)
If currentchar = Chr(112) Then currentchar = Chr(92) & Chr(93)
If currentchar = Chr(113) Then currentchar = Chr(186) & Chr(238)
If currentchar = Chr(114) Then currentchar = Chr(31) & Chr(20)
If currentchar = Chr(115) Then currentchar = Chr(214) & Chr(24)
If currentchar = Chr(116) Then currentchar = Chr(106) & Chr(36)
If currentchar = Chr(117) Then currentchar = Chr(21) & Chr(86)
If currentchar = Chr(118) Then currentchar = Chr(9) & Chr(119)
If currentchar = Chr(119) Then currentchar = Chr(196) & Chr(134)
If currentchar = Chr(120) Then currentchar = Chr(247) & Chr(110)
If currentchar = Chr(121) Then currentchar = Chr(216) & Chr(163)
If currentchar = Chr(122) Then currentchar = Chr(211) & Chr(128)
If currentchar = Chr(123) Then currentchar = Chr(245) & Chr(142)
If currentchar = Chr(124) Then currentchar = Chr(142) & Chr(37)
If currentchar = Chr(125) Then currentchar = Chr(202) & Chr(11)
If currentchar = Chr(126) Then currentchar = Chr(102) & Chr(241)
If currentchar = Chr(127) Then currentchar = Chr(161) & Chr(11)
If currentchar = Chr(128) Then currentchar = Chr(24) & Chr(5)
If currentchar = Chr(129) Then currentchar = Chr(52) & Chr(48)
If currentchar = Chr(130) Then currentchar = Chr(80) & Chr(163)
If currentchar = Chr(131) Then currentchar = Chr(193) & Chr(142)
If currentchar = Chr(132) Then currentchar = Chr(100) & Chr(32)
If currentchar = Chr(133) Then currentchar = Chr(97) & Chr(175)
If currentchar = Chr(134) Then currentchar = Chr(66) & Chr(255)
If currentchar = Chr(135) Then currentchar = Chr(149) & Chr(218)
If currentchar = Chr(136) Then currentchar = Chr(167) & Chr(91)
If currentchar = Chr(137) Then currentchar = Chr(34) & Chr(149)
If currentchar = Chr(138) Then currentchar = Chr(146) & Chr(133)
If currentchar = Chr(139) Then currentchar = Chr(31) & Chr(105)
If currentchar = Chr(140) Then currentchar = Chr(248) & Chr(238)
If currentchar = Chr(141) Then currentchar = Chr(76) & Chr(229)
If currentchar = Chr(142) Then currentchar = Chr(186) & Chr(73)
If currentchar = Chr(143) Then currentchar = Chr(165) & Chr(229)
If currentchar = Chr(144) Then currentchar = Chr(212) & Chr(233)
If currentchar = Chr(145) Then currentchar = Chr(55) & Chr(21)
If currentchar = Chr(146) Then currentchar = Chr(215) & Chr(112)
If currentchar = Chr(147) Then currentchar = Chr(51) & Chr(176)
If currentchar = Chr(148) Then currentchar = Chr(153) & Chr(171)
If currentchar = Chr(149) Then currentchar = Chr(65) & Chr(141)
If currentchar = Chr(150) Then currentchar = Chr(43) & Chr(183)
If currentchar = Chr(151) Then currentchar = Chr(23) & Chr(93)
If currentchar = Chr(152) Then currentchar = Chr(12) & Chr(105)
If currentchar = Chr(153) Then currentchar = Chr(73) & Chr(47)
If currentchar = Chr(154) Then currentchar = Chr(155) & Chr(236)
If currentchar = Chr(155) Then currentchar = Chr(98) & Chr(53)
If currentchar = Chr(156) Then currentchar = Chr(210) & Chr(172)
If currentchar = Chr(157) Then currentchar = Chr(53) & Chr(201)
If currentchar = Chr(158) Then currentchar = Chr(50) & Chr(115)
If currentchar = Chr(159) Then currentchar = Chr(110) & Chr(178)
If currentchar = Chr(160) Then currentchar = Chr(249) & Chr(210)
If currentchar = Chr(161) Then currentchar = Chr(108) & Chr(164)
If currentchar = Chr(162) Then currentchar = Chr(156) & Chr(195)
If currentchar = Chr(163) Then currentchar = Chr(83) & Chr(7)
If currentchar = Chr(164) Then currentchar = Chr(113) & Chr(140)
If currentchar = Chr(165) Then currentchar = Chr(92) & Chr(249)
If currentchar = Chr(166) Then currentchar = Chr(172) & Chr(230)
If currentchar = Chr(167) Then currentchar = Chr(112) & Chr(150)
If currentchar = Chr(168) Then currentchar = Chr(140) & Chr(95)
If currentchar = Chr(169) Then currentchar = Chr(245) & Chr(119)
If currentchar = Chr(170) Then currentchar = Chr(214) & Chr(253)
If currentchar = Chr(171) Then currentchar = Chr(232) & Chr(120)
If currentchar = Chr(172) Then currentchar = Chr(131) & Chr(37)
If currentchar = Chr(173) Then currentchar = Chr(45) & Chr(191)
If currentchar = Chr(174) Then currentchar = Chr(182) & Chr(186)
If currentchar = Chr(175) Then currentchar = Chr(35) & Chr(249)
If currentchar = Chr(176) Then currentchar = Chr(240) & Chr(136)
If currentchar = Chr(177) Then currentchar = Chr(122) & Chr(231)
If currentchar = Chr(178) Then currentchar = Chr(135) & Chr(227)
If currentchar = Chr(179) Then currentchar = Chr(73) & Chr(26)
If currentchar = Chr(180) Then currentchar = Chr(85) & Chr(141)
If currentchar = Chr(181) Then currentchar = Chr(91) & Chr(250)
If currentchar = Chr(182) Then currentchar = Chr(173) & Chr(214)
If currentchar = Chr(183) Then currentchar = Chr(202) & Chr(14)
If currentchar = Chr(184) Then currentchar = Chr(147) & Chr(7)
If currentchar = Chr(185) Then currentchar = Chr(207) & Chr(119)
If currentchar = Chr(186) Then currentchar = Chr(44) & Chr(255)
If currentchar = Chr(187) Then currentchar = Chr(217) & Chr(188)
If currentchar = Chr(188) Then currentchar = Chr(115) & Chr(31)
If currentchar = Chr(189) Then currentchar = Chr(220) & Chr(205)
If currentchar = Chr(190) Then currentchar = Chr(46) & Chr(100)
If currentchar = Chr(191) Then currentchar = Chr(236) & Chr(67)
If currentchar = Chr(192) Then currentchar = Chr(33) & Chr(211)
If currentchar = Chr(193) Then currentchar = Chr(11) & Chr(229)
If currentchar = Chr(194) Then currentchar = Chr(129) & Chr(27)
If currentchar = Chr(195) Then currentchar = Chr(62) & Chr(113)
If currentchar = Chr(196) Then currentchar = Chr(171) & Chr(116)
If currentchar = Chr(197) Then currentchar = Chr(230) & Chr(153)
If currentchar = Chr(198) Then currentchar = Chr(22) & Chr(205)
If currentchar = Chr(199) Then currentchar = Chr(78) & Chr(75)
If currentchar = Chr(200) Then currentchar = Chr(134) & Chr(39)
If currentchar = Chr(201) Then currentchar = Chr(128) & Chr(54)
If currentchar = Chr(202) Then currentchar = Chr(131) & Chr(60)
If currentchar = Chr(203) Then currentchar = Chr(93) & Chr(137)
If currentchar = Chr(204) Then currentchar = Chr(10) & Chr(101)
If currentchar = Chr(205) Then currentchar = Chr(236) & Chr(253)
If currentchar = Chr(206) Then currentchar = Chr(130) & Chr(186)
If currentchar = Chr(207) Then currentchar = Chr(242) & Chr(23)
If currentchar = Chr(208) Then currentchar = Chr(241) & Chr(123)
If currentchar = Chr(209) Then currentchar = Chr(196) & Chr(167)
If currentchar = Chr(210) Then currentchar = Chr(112) & Chr(174)
If currentchar = Chr(211) Then currentchar = Chr(90) & Chr(149)
If currentchar = Chr(212) Then currentchar = Chr(220) & Chr(11)
If currentchar = Chr(213) Then currentchar = Chr(164) & Chr(220)
If currentchar = Chr(214) Then currentchar = Chr(207) & Chr(21)
If currentchar = Chr(215) Then currentchar = Chr(33) & Chr(215)
If currentchar = Chr(216) Then currentchar = Chr(206) & Chr(136)
If currentchar = Chr(217) Then currentchar = Chr(173) & Chr(188)
If currentchar = Chr(218) Then currentchar = Chr(196) & Chr(250)
If currentchar = Chr(219) Then currentchar = Chr(154) & Chr(102)
If currentchar = Chr(220) Then currentchar = Chr(176) & Chr(189)
If currentchar = Chr(221) Then currentchar = Chr(4) & Chr(87)
If currentchar = Chr(222) Then currentchar = Chr(154) & Chr(5)
If currentchar = Chr(223) Then currentchar = Chr(93) & Chr(253)
If currentchar = Chr(224) Then currentchar = Chr(203) & Chr(71)
If currentchar = Chr(225) Then currentchar = Chr(79) & Chr(53)
If currentchar = Chr(226) Then currentchar = Chr(62) & Chr(232)
If currentchar = Chr(227) Then currentchar = Chr(196) & Chr(13)
If currentchar = Chr(228) Then currentchar = Chr(81) & Chr(23)
If currentchar = Chr(229) Then currentchar = Chr(62) & Chr(206)
If currentchar = Chr(230) Then currentchar = Chr(192) & Chr(242)
If currentchar = Chr(231) Then currentchar = Chr(109) & Chr(58)
If currentchar = Chr(232) Then currentchar = Chr(212) & Chr(241)
If currentchar = Chr(233) Then currentchar = Chr(255) & Chr(18)
If currentchar = Chr(234) Then currentchar = Chr(216) & Chr(130)
If currentchar = Chr(235) Then currentchar = Chr(187) & Chr(219)
If currentchar = Chr(236) Then currentchar = Chr(205) & Chr(240)
If currentchar = Chr(237) Then currentchar = Chr(201) & Chr(225)
If currentchar = Chr(238) Then currentchar = Chr(95) & Chr(138)
If currentchar = Chr(239) Then currentchar = Chr(85) & Chr(126)
If currentchar = Chr(240) Then currentchar = Chr(23) & Chr(1)
If currentchar = Chr(241) Then currentchar = Chr(84) & Chr(149)
If currentchar = Chr(242) Then currentchar = Chr(209) & Chr(17)
If currentchar = Chr(243) Then currentchar = Chr(164) & Chr(97)
If currentchar = Chr(244) Then currentchar = Chr(113) & Chr(99)
If currentchar = Chr(245) Then currentchar = Chr(92) & Chr(116)
If currentchar = Chr(246) Then currentchar = Chr(209) & Chr(177)
If currentchar = Chr(247) Then currentchar = Chr(90) & Chr(250)
If currentchar = Chr(248) Then currentchar = Chr(254) & Chr(45)
If currentchar = Chr(249) Then currentchar = Chr(33) & Chr(63)
If currentchar = Chr(250) Then currentchar = Chr(165) & Chr(29)
If currentchar = Chr(251) Then currentchar = Chr(230) & Chr(113)
If currentchar = Chr(252) Then currentchar = Chr(201) & Chr(198)
If currentchar = Chr(253) Then currentchar = Chr(231) & Chr(166)
If currentchar = Chr(254) Then currentchar = Chr(183) & Chr(147)
If currentchar = Chr(255) Then currentchar = Chr(1) & Chr(33)
        'add new character one at a time:
        Text1.Tag = Text1.Tag + currentchar
    Next i
'Clipboard.Clear
'Clipboard.SetText Text1.Tag
If Text1.Tag = Text2.Text Then
MsgBox Chr(84) & Chr(104) & Chr(97) & Chr(110) & Chr(107) & Chr(115) & Chr(32) & Chr(102) & Chr(111) & Chr(114) & Chr(32) & Chr(121) & Chr(111) & Chr(117) & Chr(114) & Chr(32) & Chr(109) & Chr(111) & Chr(110) & Chr(101) & Chr(121), vbExclamation, Chr(78) & Chr(111) & Chr(32) & Chr(109) & Chr(111) & Chr(114) & Chr(101) & Chr(32) & Chr(110) & Chr(97) & Chr(103) & Chr(115)
whatsalllength = Text1.Text + Text2.Text + Chr(84) & Chr(114) & Chr(117) & Chr(101)
        'Save the text file.  This writes over the the information from the last save.
        Open Chr(105) & Chr(115) & Chr(114) & Chr(101) & Chr(103) & Chr(46) & Chr(107) & Chr(101) & Chr(121) For Output As #1
        Print #1, Text1.Text
        Print #1, Text2.Text
        Print #1, Chr(84) & Chr(114) & Chr(117) & Chr(101)
        Print #1, Len(whatsalllength) * 7
        Print #1, Chr(65) & Chr(108) & Chr(108) & Chr(32) & Chr(121) & Chr(117) & Chr(114) & Chr(32) & Chr(98) & Chr(97) & Chr(115) & Chr(101) & Chr(32) & Chr(97) & Chr(114) & Chr(101) & Chr(32) & Chr(98) & Chr(101) & Chr(108) & Chr(111) & Chr(110) & Chr(103) & Chr(32) & Chr(116) & Chr(111) & Chr(32) & Chr(117) & Chr(115) & Chr(33)
        Print #1, Chr(67) & Chr(111) & Chr(100) & Chr(101) & Chr(100) & Chr(32) & Chr(98) & Chr(121) & Chr(32) & Chr(109) & Chr(101) & Chr(116) & Chr(104) & Chr(48) & Chr(115) & Chr(32) & Chr(119) & Chr(119) & Chr(119) & Chr(46) & Chr(109) & Chr(101) & Chr(116) & Chr(104) & Chr(48) & Chr(115) & Chr(46) & Chr(99) & Chr(111) & Chr(109) & Chr(32) & Chr(73) & Chr(67) & Chr(81) & Chr(35) & Chr(52) & Chr(49) & Chr(53) & Chr(57) & Chr(52) & Chr(52) & Chr(48)
        Close #1 'Close the file
Form3.Hide
Form1.Show
Else
MsgBox Chr(87) & Chr(114) & Chr(111) & Chr(110) & Chr(103) & Chr(32) & Chr(115) & Chr(101) & Chr(114) & Chr(105) & Chr(97) & Chr(108) & Chr(32) & Chr(111) & Chr(114) & Chr(32) & Chr(110) & Chr(97) & Chr(109) & Chr(101) & Chr(46) & Chr(32) & Chr(82) & Chr(101) & Chr(108) & Chr(111) & Chr(97) & Chr(100) & Chr(32) & Chr(116) & Chr(104) & Chr(101) & Chr(32) & Chr(112) & Chr(114) & Chr(111) & Chr(103) & Chr(114) & Chr(97) & Chr(109) & Chr(32) & Chr(97) & Chr(110) & Chr(100) & Chr(32) & Chr(116) & Chr(114) & Chr(121) & Chr(32) & Chr(97) & Chr(103) & Chr(97) & Chr(105) & Chr(110) & Chr(33), vbCritical, Chr(69) & Chr(114) & Chr(114) & Chr(111) & Chr(114)
End
End If

Text1.Tag = ""
currentchar = ""
End Sub

Private Sub Command3_Click()
Form3.Hide
Form1.Show
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
DragObject Form3
End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Form1.Hide
Form3.Hide
MsgBox Chr(82) & Chr(105) & Chr(103) & Chr(104) & Chr(116) & Chr(32) & Chr(99) & Chr(108) & Chr(105) & Chr(99) & Chr(107) & Chr(32) & Chr(100) & Chr(105) & Chr(115) & Chr(97) & Chr(98) & Chr(108) & Chr(101) & Chr(100), vbCritical, Chr(87) & Chr(97) & Chr(114) & Chr(110) & Chr(105) & Chr(110) & Chr(103) & Chr(33)
Form1.Show
Form3.Show
End If
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Form1.Hide
Form3.Hide
MsgBox Chr(82) & Chr(105) & Chr(103) & Chr(104) & Chr(116) & Chr(32) & Chr(99) & Chr(108) & Chr(105) & Chr(99) & Chr(107) & Chr(32) & Chr(100) & Chr(105) & Chr(115) & Chr(97) & Chr(98) & Chr(108) & Chr(101) & Chr(100), vbCritical, Chr(87) & Chr(97) & Chr(114) & Chr(110) & Chr(105) & Chr(110) & Chr(103) & Chr(33)
Form1.Show
Form3.Show
End If
End Sub
