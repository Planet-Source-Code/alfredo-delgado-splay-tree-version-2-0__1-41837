Attribute VB_Name = "ModSplay"
'***********************************************************************
' Module: ModSplay
' Author: Alfredo Delgado
' Date Created: December 19, 2002
' Modified: December 23, 2002 by Alfredo Delgado
' Description: These is an implementation of Splay Tree compression in
'              Visual Basic. These are optimize for faster compression
'              and decompression routines. It is a translation from pascal
'              routines. The constants are not declared due to memory operations
'              which takes longer time than assigning it directly in the procedures
'              and functions. A Run Length Encoding is added to preprocess the
'              file and apply splay compression thus yields a good compression ratio.
'              It is good for files with repeated sequences.
'
' Bugs: Please report any bugs found in the program.
' Any Modifications of the source code please give the author a copy of
' modifications
' Any Comments please email: fred72ph@yahoo.com
'
Option Explicit
'Variables for File Operations
Dim InBuffer() As Byte
Dim OutBuffer() As Byte
Dim RLEBuffer() As Byte
Dim InFile As Integer
Dim OutFile As Integer
Dim InFileName As String
Dim OutFileName As String
Dim Path As String
Dim Index As Long
Dim InByte As Byte
Dim OutByte As Byte
Dim FileLength As Long
Dim ExpandedFileLength As Long
Dim FileHeader(1 To 6) As Byte
Dim ExpFileLen(1 To 4) As Byte
Dim OutSize As Long
Dim RLEIndex As Long
Dim RLELength As Long
Dim RLEOutIndex As Long

'Variables used in SplayTree operations
Dim LeftTree(0 To 255) As Integer
Dim RightTree(0 To 255) As Integer
Dim UpTree(0 To 512) As Byte

'Variable used to mask Bits
Dim BitMask(0 To 7) As Byte
Dim BitPos As Byte

Private Function Expand() As Integer
   'Return next character from compressed input
   Dim A As Integer
   Dim GetB As Byte
   'Scan the tree to a leaf, which determines the character
   A = 0
   Do
      If BitPos = 7 Then
         'Used up the bits in current byte, get another
         InByte = GetByte
         BitPos = 0
      Else
         BitPos = BitPos + 1  'Increment BitPos
      End If
      GetB = InByte And BitMask(BitPos)
      If GetB = 0 Then        'If InByte And BitMask(BitPos) = 0 Then
         A = LeftTree(A)
      Else
         A = RightTree(A)
      End If
   Loop Until A > 255

   'Update the code tree
   A = A - 256 'Decrement A by 256
   Call Splay(A)

   'Return the character
   Expand = A
End Function

Private Function GetByte() As Byte
   '-Return next byte from compressed input
   If Index < FileLength Then
      Index = Index + 1         'Increment Index
      GetByte = InBuffer(Index) 'Get a Byte from Input File Buffer
   End If
End Function

Private Sub Compress(Plain As Byte)
   '-Compress a single character
   Dim A As Integer
   Dim U As Integer
   Dim Sp As Integer   'Stack Pointer
   Dim Stack(0 To 255) As Boolean 'Stack
   
   A = Plain + 256
   Sp = 0

   'Walk up the tree pushing bits onto stack
   Do
      U = UpTree(A)
      If RightTree(U) = A Then
         Stack(Sp) = True
      Else
         Stack(Sp) = False
      End If
      Sp = Sp + 1 'Increment Sp
      A = U
   Loop Until A = 0

   'Unstack to transmit bits in correct order
   Do
      Sp = Sp - 1 'Decrement Sp
      If Stack(Sp) = True Then
        OutByte = OutByte Or BitMask(BitPos)
      End If
      If BitPos = 7 Then
         'Byte filled with bits, write it out
         Put OutFile, , OutByte  'Write the Byte to OutputFile
         BitPos = 0
         OutByte = 0
      Else
         BitPos = BitPos + 1 'Increment BitPos
      End If
   Loop Until Sp = 0

   'Update the tree
   Call Splay(CInt(Plain))
End Sub

Private Sub CompressFile()
   'Compress Infile, writing to OutFile
   Dim TempStr As String
   Dim cnt As Integer
   Dim I As Integer
   Dim TempByte(1 To 4) As Byte
   
   'Write header to output
   Put OutFile, , CByte(65)
   Put OutFile, , CByte(66)
   Put OutFile, , CByte(68)
   Put OutFile, , CByte(90)
   Put OutFile, , CByte(65)
   Put OutFile, , CByte(80)
   I = 1
   TempStr = Hex$(RLELength)
   For cnt = 1 To (8 - Len(TempStr))
      TempStr = "0" & TempStr
   Next cnt
   For cnt = 1 To 8 Step 2
       TempByte(I) = CByte(Val("&H" & Mid$(TempStr, cnt, 2)))
       I = I + 1
   Next cnt
   For I = 1 To 4
      Put OutFile, , (TempByte(I)) 'Put Outfile, , TempByte(1): Put Outfile, , TempByte(2): Put Outfile, , TempByte(3): Put Outfile, , TempByte(4)
   Next I
   'Compress file
   BitPos = 0
   OutByte = 0
   For Index = 1 To RLELength
      Call Compress(CInt(RLEBuffer(Index)))
   Next Index
   
   'Mark end of file
   'Call Compress(EofChar) 'This is handled by ExpandedFileLength when decompressing a file

   'Flush buffers
   If BitPos <> 0 Then
      Put OutFile, , OutByte 'Write Byte to file
   End If
End Sub

Private Sub ExpandFile()
   'Uncompress the input file and write to output file
   'Force buffer load first time
   Index = 0
   ' {Nothing in output buffer}
   OutSize = 1
   'Force bit buffer load first time}
   BitPos = 7
   'Read and expand the compressed input
   OutByte = CByte(Expand)
   Do While OutSize <> ExpandedFileLength ' OutByte <> EofChar
      OutBuffer(OutSize) = OutByte 'Put OutFile, , OutByte
      OutSize = OutSize + 1
      OutByte = CByte(Expand)
   Loop
End Sub

Private Sub InitSplayTree()
   'Initialize Bit Mask
   BitMask(0) = 1: BitMask(1) = 2: BitMask(2) = 4: _
   BitMask(3) = 8: BitMask(4) = 16: BitMask(5) = 32: _
   BitMask(6) = 64: BitMask(7) = 128
   
   'These initial Tree values are generated by the original procedures
   'which take longer time process than assigning values to the Trees
     
   'Initialize UpTree
   UpTree(1) = 0: UpTree(2) = 0: UpTree(3) = 1: _
   UpTree(4) = 1: UpTree(5) = 2: UpTree(6) = 2: _
   UpTree(7) = 3: UpTree(8) = 3: UpTree(9) = 4: _
   UpTree(10) = 4: UpTree(11) = 5: UpTree(12) = 5: _
   UpTree(13) = 6: UpTree(14) = 6: UpTree(15) = 7: _
   UpTree(16) = 7: UpTree(17) = 8: UpTree(18) = 8: _
   UpTree(19) = 9: UpTree(20) = 9: UpTree(21) = 10: _
   UpTree(22) = 10: UpTree(23) = 11: UpTree(24) = 11: _
   UpTree(25) = 12: UpTree(26) = 12: UpTree(27) = 13: _
   UpTree(28) = 13: UpTree(29) = 14: UpTree(30) = 14: _
   UpTree(31) = 15: UpTree(32) = 15: UpTree(33) = 16: _
   UpTree(34) = 16: UpTree(35) = 17: UpTree(36) = 17: _
   UpTree(37) = 18: UpTree(38) = 18: UpTree(39) = 19: _
   UpTree(40) = 19: UpTree(41) = 20: UpTree(42) = 20: _
   UpTree(43) = 21: UpTree(44) = 21: UpTree(45) = 22: _
   UpTree(46) = 22: UpTree(47) = 23: UpTree(48) = 23: _
   UpTree(49) = 24: UpTree(50) = 24: UpTree(51) = 25: _
   UpTree(52) = 25: UpTree(53) = 26: UpTree(54) = 26: _
   UpTree(55) = 27: UpTree(56) = 27: UpTree(57) = 28: _
   UpTree(58) = 28: UpTree(59) = 29: UpTree(60) = 29: _
   UpTree(61) = 30: UpTree(62) = 30: UpTree(63) = 31: _
   UpTree(64) = 31: UpTree(65) = 32: UpTree(66) = 32: _
   UpTree(67) = 33: UpTree(68) = 33: UpTree(69) = 34: _
   UpTree(70) = 34: UpTree(71) = 35: UpTree(72) = 35: _
   UpTree(73) = 36: UpTree(74) = 36: UpTree(75) = 37
   UpTree(76) = 37: UpTree(77) = 38: UpTree(78) = 38: _
   UpTree(79) = 39: UpTree(80) = 39: UpTree(81) = 40: _
   UpTree(82) = 40: UpTree(83) = 41: UpTree(84) = 41: _
   UpTree(85) = 42: UpTree(86) = 42: UpTree(87) = 43: _
   UpTree(88) = 43: UpTree(89) = 44: UpTree(90) = 44: _
   UpTree(91) = 45: UpTree(92) = 45: UpTree(93) = 46: _
   UpTree(94) = 46: UpTree(95) = 47: UpTree(96) = 47: _
   UpTree(97) = 48: UpTree(98) = 48: UpTree(99) = 49: _
   UpTree(100) = 49: UpTree(101) = 50: UpTree(102) = 50: _
   UpTree(103) = 51: UpTree(104) = 51: UpTree(105) = 52: _
   UpTree(106) = 52: UpTree(107) = 53: UpTree(108) = 53: _
   UpTree(109) = 54: UpTree(110) = 54: UpTree(111) = 55: _
   UpTree(112) = 55: UpTree(113) = 56: UpTree(114) = 56: _
   UpTree(115) = 57: UpTree(116) = 57: UpTree(117) = 58: _
   UpTree(118) = 58: UpTree(119) = 59: UpTree(120) = 59: _
   UpTree(121) = 60: UpTree(122) = 60: UpTree(123) = 61: _
   UpTree(124) = 61: UpTree(125) = 62: UpTree(126) = 62: _
   UpTree(127) = 63: UpTree(128) = 63: UpTree(129) = 64: _
   UpTree(130) = 64: UpTree(131) = 65: UpTree(132) = 65: _
   UpTree(133) = 66: UpTree(134) = 66: UpTree(135) = 67: _
   UpTree(136) = 67: UpTree(137) = 68: UpTree(138) = 68: _
   UpTree(139) = 69: UpTree(140) = 69: UpTree(141) = 70: _
   UpTree(142) = 70: UpTree(143) = 71: UpTree(144) = 71: _
   UpTree(145) = 72: UpTree(146) = 72: UpTree(147) = 73: _
   UpTree(148) = 73: UpTree(149) = 74: UpTree(150) = 74
   UpTree(151) = 75: UpTree(152) = 75: UpTree(153) = 76: _
   UpTree(154) = 76: UpTree(155) = 77: UpTree(156) = 77: _
   UpTree(157) = 78: UpTree(158) = 78: UpTree(159) = 79: _
   UpTree(160) = 79: UpTree(161) = 80: UpTree(162) = 80: _
   UpTree(163) = 81: UpTree(164) = 81: UpTree(165) = 82: _
   UpTree(166) = 82: UpTree(167) = 83: UpTree(168) = 83: _
   UpTree(169) = 84: UpTree(170) = 84: UpTree(171) = 85: _
   UpTree(172) = 85: UpTree(173) = 86: UpTree(174) = 86: _
   UpTree(175) = 87: UpTree(176) = 87: UpTree(177) = 88: _
   UpTree(178) = 88: UpTree(179) = 89: UpTree(180) = 89: _
   UpTree(181) = 90: UpTree(182) = 90: UpTree(183) = 91: _
   UpTree(184) = 91: UpTree(185) = 92: UpTree(186) = 92: _
   UpTree(187) = 93: UpTree(188) = 93: UpTree(189) = 94: _
   UpTree(190) = 94: UpTree(191) = 95: UpTree(192) = 95: _
   UpTree(193) = 96: UpTree(194) = 96: UpTree(195) = 97: _
   UpTree(196) = 97: UpTree(197) = 98: UpTree(198) = 98: _
   UpTree(199) = 99: UpTree(200) = 99: UpTree(201) = 100: _
   UpTree(202) = 100: UpTree(203) = 101: UpTree(204) = 101: _
   UpTree(205) = 102: UpTree(206) = 102: UpTree(207) = 103: _
   UpTree(208) = 103: UpTree(209) = 104: UpTree(210) = 104: _
   UpTree(211) = 105: UpTree(212) = 105: UpTree(213) = 106: _
   UpTree(214) = 106: UpTree(215) = 107: UpTree(216) = 107: _
   UpTree(217) = 108: UpTree(218) = 108: UpTree(219) = 109: _
   UpTree(220) = 109: UpTree(221) = 110: UpTree(222) = 110: _
   UpTree(223) = 111: UpTree(224) = 111: UpTree(225) = 112
   UpTree(226) = 112: UpTree(227) = 113: UpTree(228) = 113: _
   UpTree(229) = 114: UpTree(230) = 114: UpTree(231) = 115: _
   UpTree(232) = 115: UpTree(233) = 116: UpTree(234) = 116: _
   UpTree(235) = 117: UpTree(236) = 117: UpTree(237) = 118: _
   UpTree(238) = 118: UpTree(239) = 119: UpTree(240) = 119: _
   UpTree(241) = 120: UpTree(242) = 120: UpTree(243) = 121: _
   UpTree(244) = 121: UpTree(245) = 122: UpTree(246) = 122: _
   UpTree(247) = 123: UpTree(248) = 123: UpTree(249) = 124: _
   UpTree(250) = 124: UpTree(251) = 125: UpTree(252) = 125: _
   UpTree(253) = 126: UpTree(254) = 126: UpTree(255) = 127: _
   UpTree(256) = 127: UpTree(257) = 128: UpTree(258) = 128: _
   UpTree(259) = 129: UpTree(260) = 129: UpTree(261) = 130: _
   UpTree(262) = 130: UpTree(263) = 131: UpTree(264) = 131: _
   UpTree(265) = 132: UpTree(266) = 132: UpTree(267) = 133: _
   UpTree(268) = 133: UpTree(269) = 134: UpTree(270) = 134: _
   UpTree(271) = 135: UpTree(272) = 135: UpTree(273) = 136: _
   UpTree(274) = 136: UpTree(275) = 137: UpTree(276) = 137: _
   UpTree(277) = 138: UpTree(278) = 138: UpTree(279) = 139: _
   UpTree(280) = 139: UpTree(281) = 140: UpTree(282) = 140: _
   UpTree(283) = 141: UpTree(284) = 141: UpTree(285) = 142: _
   UpTree(286) = 142: UpTree(287) = 143: UpTree(288) = 143: _
   UpTree(289) = 144: UpTree(290) = 144: UpTree(291) = 145: _
   UpTree(292) = 145: UpTree(293) = 146: UpTree(294) = 146: _
   UpTree(295) = 147: UpTree(296) = 147: UpTree(297) = 148: _
   UpTree(298) = 148: UpTree(299) = 149: UpTree(300) = 149
   UpTree(301) = 150: UpTree(302) = 150: UpTree(303) = 151: _
   UpTree(304) = 151: UpTree(305) = 152: UpTree(306) = 152: _
   UpTree(307) = 153: UpTree(308) = 153: UpTree(309) = 154: _
   UpTree(310) = 154: UpTree(311) = 155: UpTree(312) = 155: _
   UpTree(313) = 156: UpTree(314) = 156: UpTree(315) = 157: _
   UpTree(316) = 157: UpTree(317) = 158: UpTree(318) = 158: _
   UpTree(319) = 159: UpTree(320) = 159: UpTree(321) = 160: _
   UpTree(322) = 160: UpTree(323) = 161: UpTree(324) = 161: _
   UpTree(325) = 162: UpTree(326) = 162: UpTree(327) = 163: _
   UpTree(328) = 163: UpTree(329) = 164: UpTree(330) = 164: _
   UpTree(331) = 165: UpTree(332) = 165: UpTree(333) = 166: _
   UpTree(334) = 166: UpTree(335) = 167: UpTree(336) = 167: _
   UpTree(337) = 168: UpTree(338) = 168: UpTree(339) = 169: _
   UpTree(340) = 169: UpTree(341) = 170: UpTree(342) = 170: _
   UpTree(343) = 171: UpTree(344) = 171: UpTree(345) = 172: _
   UpTree(346) = 172: UpTree(347) = 173: UpTree(348) = 173: _
   UpTree(349) = 174: UpTree(350) = 174: UpTree(351) = 175: _
   UpTree(352) = 175: UpTree(353) = 176: UpTree(354) = 176: _
   UpTree(355) = 177: UpTree(356) = 177: UpTree(357) = 178: _
   UpTree(358) = 178: UpTree(359) = 179: UpTree(360) = 179: _
   UpTree(361) = 180: UpTree(362) = 180: UpTree(363) = 181: _
   UpTree(364) = 181: UpTree(365) = 182: UpTree(366) = 182: _
   UpTree(367) = 183: UpTree(368) = 183: UpTree(369) = 184: _
   UpTree(370) = 184: UpTree(371) = 185: UpTree(372) = 185: _
   UpTree(373) = 186: UpTree(374) = 186: UpTree(375) = 187
   UpTree(376) = 187: UpTree(377) = 188: UpTree(378) = 188: _
   UpTree(379) = 189: UpTree(380) = 189: UpTree(381) = 190: _
   UpTree(382) = 190: UpTree(383) = 191: UpTree(384) = 191: _
   UpTree(385) = 192: UpTree(386) = 192: UpTree(387) = 193: _
   UpTree(388) = 193: UpTree(389) = 194: UpTree(390) = 194: _
   UpTree(391) = 195: UpTree(392) = 195: UpTree(393) = 196: _
   UpTree(394) = 196: UpTree(395) = 197: UpTree(396) = 197: _
   UpTree(397) = 198: UpTree(398) = 198: UpTree(399) = 199: _
   UpTree(400) = 199: UpTree(401) = 200: UpTree(402) = 200: _
   UpTree(403) = 201: UpTree(404) = 201: UpTree(405) = 202: _
   UpTree(406) = 202: UpTree(407) = 203: UpTree(408) = 203: _
   UpTree(409) = 204: UpTree(410) = 204: UpTree(411) = 205: _
   UpTree(412) = 205: UpTree(413) = 206: UpTree(414) = 206: _
   UpTree(415) = 207: UpTree(416) = 207: UpTree(417) = 208: _
   UpTree(418) = 208: UpTree(419) = 209: UpTree(420) = 209: _
   UpTree(421) = 210: UpTree(422) = 210: UpTree(423) = 211: _
   UpTree(424) = 211: UpTree(425) = 212: UpTree(426) = 212: _
   UpTree(427) = 213: UpTree(428) = 213: UpTree(429) = 214: _
   UpTree(430) = 214: UpTree(431) = 215: UpTree(432) = 215: _
   UpTree(433) = 216: UpTree(434) = 216: UpTree(435) = 217: _
   UpTree(436) = 217: UpTree(437) = 218: UpTree(438) = 218: _
   UpTree(439) = 219: UpTree(440) = 219: UpTree(441) = 220: _
   UpTree(442) = 220: UpTree(443) = 221: UpTree(444) = 221: _
   UpTree(445) = 222: UpTree(446) = 222: UpTree(447) = 223: _
   UpTree(448) = 223: UpTree(449) = 224: UpTree(450) = 224
   UpTree(451) = 225: UpTree(452) = 225: UpTree(453) = 226: _
   UpTree(454) = 226: UpTree(455) = 227: UpTree(456) = 227: _
   UpTree(457) = 228: UpTree(458) = 228: UpTree(459) = 229: _
   UpTree(460) = 229: UpTree(461) = 230: UpTree(462) = 230: _
   UpTree(463) = 231: UpTree(464) = 231: UpTree(465) = 232: _
   UpTree(466) = 232: UpTree(467) = 233: UpTree(468) = 233: _
   UpTree(469) = 234: UpTree(470) = 234: UpTree(471) = 235: _
   UpTree(472) = 235: UpTree(473) = 236: UpTree(474) = 236: _
   UpTree(475) = 237: UpTree(476) = 237: UpTree(477) = 238: _
   UpTree(478) = 238: UpTree(479) = 239: UpTree(480) = 239: _
   UpTree(481) = 240: UpTree(482) = 240: UpTree(483) = 241: _
   UpTree(484) = 241: UpTree(485) = 242: UpTree(486) = 242: _
   UpTree(487) = 243: UpTree(488) = 243: UpTree(489) = 244: _
   UpTree(490) = 244: UpTree(491) = 245: UpTree(492) = 245: _
   UpTree(493) = 246: UpTree(494) = 246: UpTree(495) = 247: _
   UpTree(496) = 247: UpTree(497) = 248: UpTree(498) = 248: _
   UpTree(499) = 249: UpTree(500) = 249: UpTree(501) = 250: _
   UpTree(502) = 250: UpTree(503) = 251: UpTree(504) = 251: _
   UpTree(505) = 252: UpTree(506) = 252: UpTree(507) = 253: _
   UpTree(508) = 253: UpTree(509) = 254: UpTree(510) = 254: _
   UpTree(511) = 255: UpTree(512) = 255
   
   'Initialize LeftTree
   LeftTree(0) = 1: LeftTree(1) = 3: LeftTree(2) = 5: _
   LeftTree(3) = 7: LeftTree(4) = 9: LeftTree(5) = 11: _
   LeftTree(6) = 13: LeftTree(7) = 15: LeftTree(8) = 17: _
   LeftTree(9) = 19: LeftTree(10) = 21: LeftTree(11) = 23: _
   LeftTree(12) = 25: LeftTree(13) = 27: LeftTree(14) = 29: _
   LeftTree(15) = 31: LeftTree(16) = 33: LeftTree(17) = 35: _
   LeftTree(18) = 37: LeftTree(19) = 39: LeftTree(20) = 41: _
   LeftTree(21) = 43: LeftTree(22) = 45: LeftTree(23) = 47: _
   LeftTree(24) = 49: LeftTree(25) = 51: LeftTree(26) = 53: _
   LeftTree(27) = 55: LeftTree(28) = 57: LeftTree(29) = 59: _
   LeftTree(30) = 61: LeftTree(31) = 63: LeftTree(32) = 65: _
   LeftTree(33) = 67: LeftTree(34) = 69: LeftTree(35) = 71: _
   LeftTree(36) = 73: LeftTree(37) = 75: LeftTree(38) = 77: _
   LeftTree(39) = 79: LeftTree(40) = 81: LeftTree(41) = 83: _
   LeftTree(42) = 85: LeftTree(43) = 87: LeftTree(44) = 89: _
   LeftTree(45) = 91: LeftTree(46) = 93: LeftTree(47) = 95: _
   LeftTree(48) = 97: LeftTree(49) = 99: LeftTree(50) = 101: _
   LeftTree(51) = 103: LeftTree(52) = 105: LeftTree(53) = 107: _
   LeftTree(54) = 109: LeftTree(55) = 111: LeftTree(56) = 113: _
   LeftTree(57) = 115: LeftTree(58) = 117: LeftTree(59) = 119: _
   LeftTree(60) = 121: LeftTree(61) = 123: LeftTree(62) = 125: _
   LeftTree(63) = 127: LeftTree(64) = 129: LeftTree(65) = 131: _
   LeftTree(66) = 133: LeftTree(67) = 135: LeftTree(68) = 137: _
   LeftTree(69) = 139: LeftTree(70) = 141: LeftTree(71) = 143: _
   LeftTree(72) = 145: LeftTree(73) = 147: LeftTree(74) = 149
   LeftTree(75) = 151: LeftTree(76) = 153: LeftTree(77) = 155: _
   LeftTree(78) = 157: LeftTree(79) = 159: LeftTree(80) = 161: _
   LeftTree(81) = 163: LeftTree(82) = 165: LeftTree(83) = 167: _
   LeftTree(84) = 169: LeftTree(85) = 171: LeftTree(86) = 173: _
   LeftTree(87) = 175: LeftTree(88) = 177: LeftTree(89) = 179: _
   LeftTree(90) = 181: LeftTree(91) = 183: LeftTree(92) = 185: _
   LeftTree(93) = 187: LeftTree(94) = 189: LeftTree(95) = 191: _
   LeftTree(96) = 193: LeftTree(97) = 195: LeftTree(98) = 197: _
   LeftTree(99) = 199: LeftTree(100) = 201: LeftTree(101) = 203: _
   LeftTree(102) = 205: LeftTree(103) = 207: LeftTree(104) = 209: _
   LeftTree(105) = 211: LeftTree(106) = 213: LeftTree(107) = 215: _
   LeftTree(108) = 217: LeftTree(109) = 219: LeftTree(110) = 221: _
   LeftTree(111) = 223: LeftTree(112) = 225: LeftTree(113) = 227: _
   LeftTree(114) = 229: LeftTree(115) = 231: LeftTree(116) = 233: _
   LeftTree(117) = 235: LeftTree(118) = 237: LeftTree(119) = 239: _
   LeftTree(120) = 241: LeftTree(121) = 243: LeftTree(122) = 245: _
   LeftTree(123) = 247: LeftTree(124) = 249: LeftTree(125) = 251: _
   LeftTree(126) = 253: LeftTree(127) = 255: LeftTree(128) = 257: _
   LeftTree(129) = 259: LeftTree(130) = 261: LeftTree(131) = 263: _
   LeftTree(132) = 265: LeftTree(133) = 267: LeftTree(134) = 269: _
   LeftTree(135) = 271: LeftTree(136) = 273: LeftTree(137) = 275: _
   LeftTree(138) = 277: LeftTree(139) = 279: LeftTree(140) = 281: _
   LeftTree(141) = 283: LeftTree(142) = 285: LeftTree(143) = 287: _
   LeftTree(144) = 289: LeftTree(145) = 291: LeftTree(146) = 293: _
   LeftTree(147) = 295: LeftTree(148) = 297: LeftTree(149) = 299
   LeftTree(150) = 301: LeftTree(151) = 303: LeftTree(152) = 305: _
   LeftTree(153) = 307: LeftTree(154) = 309: LeftTree(155) = 311: _
   LeftTree(156) = 313: LeftTree(157) = 315: LeftTree(158) = 317: _
   LeftTree(159) = 319: LeftTree(160) = 321: LeftTree(161) = 323: _
   LeftTree(162) = 325: LeftTree(163) = 327: LeftTree(164) = 329: _
   LeftTree(165) = 331: LeftTree(166) = 333: LeftTree(167) = 335: _
   LeftTree(168) = 337: LeftTree(169) = 339: LeftTree(170) = 341: _
   LeftTree(171) = 343: LeftTree(172) = 345: LeftTree(173) = 347: _
   LeftTree(174) = 349: LeftTree(175) = 351: LeftTree(176) = 353: _
   LeftTree(177) = 355: LeftTree(178) = 357: LeftTree(179) = 359: _
   LeftTree(180) = 361: LeftTree(181) = 363: LeftTree(182) = 365: _
   LeftTree(183) = 367: LeftTree(184) = 369: LeftTree(185) = 371: _
   LeftTree(186) = 373: LeftTree(187) = 375: LeftTree(188) = 377: _
   LeftTree(189) = 379: LeftTree(190) = 381: LeftTree(191) = 383: _
   LeftTree(192) = 385: LeftTree(193) = 387: LeftTree(194) = 389: _
   LeftTree(195) = 391: LeftTree(196) = 393: LeftTree(197) = 395: _
   LeftTree(198) = 397: LeftTree(199) = 399: LeftTree(200) = 401: _
   LeftTree(201) = 403: LeftTree(202) = 405: LeftTree(203) = 407: _
   LeftTree(204) = 409: LeftTree(205) = 411: LeftTree(206) = 413: _
   LeftTree(207) = 415: LeftTree(208) = 417: LeftTree(209) = 419: _
   LeftTree(210) = 421: LeftTree(211) = 423: LeftTree(212) = 425: _
   LeftTree(213) = 427: LeftTree(214) = 429: LeftTree(215) = 431: _
   LeftTree(216) = 433: LeftTree(217) = 435: LeftTree(218) = 437: _
   LeftTree(219) = 439: LeftTree(220) = 441: LeftTree(221) = 443: _
   LeftTree(222) = 445: LeftTree(223) = 447: LeftTree(224) = 449
   LeftTree(225) = 451: LeftTree(226) = 453: LeftTree(227) = 455: _
   LeftTree(228) = 457: LeftTree(229) = 459: LeftTree(230) = 461: _
   LeftTree(231) = 463: LeftTree(232) = 465: LeftTree(233) = 467: _
   LeftTree(234) = 469: LeftTree(235) = 471: LeftTree(236) = 473: _
   LeftTree(237) = 475: LeftTree(238) = 477: LeftTree(239) = 479: _
   LeftTree(240) = 481: LeftTree(241) = 483: LeftTree(242) = 485: _
   LeftTree(243) = 487: LeftTree(244) = 489: LeftTree(245) = 491: _
   LeftTree(246) = 493: LeftTree(247) = 495: LeftTree(248) = 497: _
   LeftTree(249) = 499: LeftTree(250) = 501: LeftTree(251) = 503: _
   LeftTree(252) = 505: LeftTree(253) = 507: LeftTree(254) = 509: _
   LeftTree(255) = 511
   
   'Initialize RightTree
   RightTree(0) = 2: RightTree(1) = 4: RightTree(2) = 6: _
   RightTree(3) = 8: RightTree(4) = 10: RightTree(5) = 12: _
   RightTree(6) = 14: RightTree(7) = 16: RightTree(8) = 18: _
   RightTree(9) = 20: RightTree(10) = 22: RightTree(11) = 24: _
   RightTree(12) = 26: RightTree(13) = 28: RightTree(14) = 30: _
   RightTree(15) = 32: RightTree(16) = 34: RightTree(17) = 36: _
   RightTree(18) = 38: RightTree(19) = 40: RightTree(20) = 42: _
   RightTree(21) = 44: RightTree(22) = 46: RightTree(23) = 48: _
   RightTree(24) = 50: RightTree(25) = 52: RightTree(26) = 54: _
   RightTree(27) = 56: RightTree(28) = 58: RightTree(29) = 60: _
   RightTree(30) = 62: RightTree(31) = 64: RightTree(32) = 66: _
   RightTree(33) = 68: RightTree(34) = 70: RightTree(35) = 72: _
   RightTree(36) = 74: RightTree(37) = 76: RightTree(38) = 78: _
   RightTree(39) = 80: RightTree(40) = 82: RightTree(41) = 84: _
   RightTree(42) = 86: RightTree(43) = 88: RightTree(44) = 90: _
   RightTree(45) = 92: RightTree(46) = 94: RightTree(47) = 96: _
   RightTree(48) = 98: RightTree(49) = 100: RightTree(50) = 102: _
   RightTree(51) = 104: RightTree(52) = 106: RightTree(53) = 108: _
   RightTree(54) = 110: RightTree(55) = 112: RightTree(56) = 114: _
   RightTree(57) = 116: RightTree(58) = 118: RightTree(59) = 120: _
   RightTree(60) = 122: RightTree(61) = 124: RightTree(62) = 126: _
   RightTree(63) = 128: RightTree(64) = 130: RightTree(65) = 132: _
   RightTree(66) = 134: RightTree(67) = 136: RightTree(68) = 138: _
   RightTree(69) = 140: RightTree(70) = 142: RightTree(71) = 144: _
   RightTree(72) = 146: RightTree(73) = 148: RightTree(74) = 150
   RightTree(75) = 152: RightTree(76) = 154: RightTree(77) = 156: _
   RightTree(78) = 158: RightTree(79) = 160: RightTree(80) = 162: _
   RightTree(81) = 164: RightTree(82) = 166: RightTree(83) = 168: _
   RightTree(84) = 170: RightTree(85) = 172: RightTree(86) = 174: _
   RightTree(87) = 176: RightTree(88) = 178: RightTree(89) = 180: _
   RightTree(90) = 182: RightTree(91) = 184: RightTree(92) = 186: _
   RightTree(93) = 188: RightTree(94) = 190: RightTree(95) = 192: _
   RightTree(96) = 194: RightTree(97) = 196: RightTree(98) = 198: _
   RightTree(99) = 200: RightTree(100) = 202: RightTree(101) = 204: _
   RightTree(102) = 206: RightTree(103) = 208: RightTree(104) = 210: _
   RightTree(105) = 212: RightTree(106) = 214: RightTree(107) = 216: _
   RightTree(108) = 218: RightTree(109) = 220: RightTree(110) = 222: _
   RightTree(111) = 224: RightTree(112) = 226: RightTree(113) = 228: _
   RightTree(114) = 230: RightTree(115) = 232: RightTree(116) = 234: _
   RightTree(117) = 236: RightTree(118) = 238: RightTree(119) = 240: _
   RightTree(120) = 242: RightTree(121) = 244: RightTree(122) = 246: _
   RightTree(123) = 248: RightTree(124) = 250: RightTree(125) = 252: _
   RightTree(126) = 254: RightTree(127) = 256: RightTree(128) = 258: _
   RightTree(129) = 260: RightTree(130) = 262: RightTree(131) = 264: _
   RightTree(132) = 266: RightTree(133) = 268: RightTree(134) = 270: _
   RightTree(135) = 272: RightTree(136) = 274: RightTree(137) = 276: _
   RightTree(138) = 278: RightTree(139) = 280: RightTree(140) = 282: _
   RightTree(141) = 284: RightTree(142) = 286: RightTree(143) = 288: _
   RightTree(144) = 290: RightTree(145) = 292: RightTree(146) = 294: _
   RightTree(147) = 296: RightTree(148) = 298: RightTree(149) = 300
   RightTree(150) = 302: RightTree(151) = 304: RightTree(152) = 306: _
   RightTree(153) = 308: RightTree(154) = 310: RightTree(155) = 312: _
   RightTree(156) = 314: RightTree(157) = 316: RightTree(158) = 318: _
   RightTree(159) = 320: RightTree(160) = 322: RightTree(161) = 324: _
   RightTree(162) = 326: RightTree(163) = 328: RightTree(164) = 330: _
   RightTree(165) = 332: RightTree(166) = 334: RightTree(167) = 336: _
   RightTree(168) = 338: RightTree(169) = 340: RightTree(170) = 342: _
   RightTree(171) = 344: RightTree(172) = 346: RightTree(173) = 348: _
   RightTree(174) = 350: RightTree(175) = 352: RightTree(176) = 354: _
   RightTree(177) = 356: RightTree(178) = 358: RightTree(179) = 360: _
   RightTree(180) = 362: RightTree(181) = 364: RightTree(182) = 366: _
   RightTree(183) = 368: RightTree(184) = 370: RightTree(185) = 372: _
   RightTree(186) = 374: RightTree(187) = 376: RightTree(188) = 378: _
   RightTree(189) = 380: RightTree(190) = 382: RightTree(191) = 384: _
   RightTree(192) = 386: RightTree(193) = 388: RightTree(194) = 390: _
   RightTree(195) = 392: RightTree(196) = 394: RightTree(197) = 396: _
   RightTree(198) = 398: RightTree(199) = 400: RightTree(200) = 402: _
   RightTree(201) = 404: RightTree(202) = 406: RightTree(203) = 408: _
   RightTree(204) = 410: RightTree(205) = 412: RightTree(206) = 414: _
   RightTree(207) = 416: RightTree(208) = 418: RightTree(209) = 420: _
   RightTree(210) = 422: RightTree(211) = 424: RightTree(212) = 426: _
   RightTree(213) = 428: RightTree(214) = 430: RightTree(215) = 432: _
   RightTree(216) = 434: RightTree(217) = 436: RightTree(218) = 438: _
   RightTree(219) = 440: RightTree(220) = 442: RightTree(221) = 444: _
   RightTree(222) = 446: RightTree(223) = 448: RightTree(224) = 450
   RightTree(225) = 452: RightTree(226) = 454: RightTree(227) = 456: _
   RightTree(228) = 458: RightTree(229) = 460: RightTree(230) = 462: _
   RightTree(231) = 464: RightTree(232) = 466: RightTree(233) = 468: _
   RightTree(234) = 470: RightTree(235) = 472: RightTree(236) = 474: _
   RightTree(237) = 476: RightTree(238) = 478: RightTree(239) = 480: _
   RightTree(240) = 482: RightTree(241) = 484: RightTree(242) = 486: _
   RightTree(243) = 488: RightTree(244) = 490: RightTree(245) = 492: _
   RightTree(246) = 494: RightTree(247) = 496: RightTree(248) = 498: _
   RightTree(249) = 500: RightTree(250) = 502: RightTree(251) = 504: _
   RightTree(252) = 506: RightTree(253) = 508: RightTree(254) = 510: _
   RightTree(255) = 512
End Sub

Private Sub ReadHeader()
   '-Read a compressed file header
   Get InFile, , FileHeader
   If (FileHeader(1) <> 65) And (FileHeader(2) <> 66) And (FileHeader(3) <> 68) _
      And (FileHeader(4) <> 90) And (FileHeader(5) <> 65) And (FileHeader(6) <> 80) Then
      MsgBox "Unrecognized file format!", vbCritical, "Error"
      End
   End If
End Sub

Private Sub Splay(Plain As Integer)
   'Rearrange the splay tree for each succeeding character
   Dim A As Integer
   Dim b As Integer
   Dim C As Integer
   Dim D As Integer
   A = Plain + 256
   Do
      'Walk up the tree semi-rotating pairs
      C = UpTree(A)
      If C <> 0 Then
         'A pair remains
         D = UpTree(C)
         'Exchange children of pair
         b = LeftTree(D)
         If C = b Then
            b = RightTree(D)
            RightTree(D) = A
         Else
            LeftTree(D) = A
         End If
         If A = LeftTree(C) Then
            LeftTree(C) = b
         Else
            RightTree(C) = b
         End If
        
         UpTree(A) = D
         UpTree(b) = C
         A = D
         
      Else
         'Handle odd node at end
         A = C
      End If
   Loop Until A = 0
 
End Sub

Public Sub SplayCompress(FName As String)
   'Splay Compress Procedure for compressing a file
   Dim OutFileName As String
   Dim InFileName As String
   Call InitSplayTree
   'Path = "C:\"
   InFileName = FName 'Path + FName
   OutFileName = FName + ".ZAP" 'Path + FName + ".ZAP"
   InFile = FreeFile
   Close InFile
   FileLength = FileLen(InFileName)
   ReDim InBuffer(1 To FileLength)
   ReDim RLEBuffer(1 To FileLength)
   Open InFileName For Binary Access Read As InFile
   Get InFile, , InBuffer
   Close InFile
   OutFile = FreeFile
   Close OutFile
   Call RunLengthEncode
   ReDim Preserve RLEBuffer(1 To RLELength)
   Open OutFileName For Binary Access Write As OutFile
   Call CompressFile
   Close OutFile
End Sub

Public Sub SplayExpand(FName As String)
   Dim OutFileName As String
   Dim InFileName As String
   Dim cnt As Integer
   Dim TempStr As String
   Dim ExtStr As String
   Dim ExStr As String
   Dim Ex1Str As String
   ExStr = ".zap"
   Ex1Str = ".ZAP"
   Call InitSplayTree
   'Path = "C:\"
   InFileName = FName 'Path + FName
   ExtStr = Right$(InFileName, 4)
   'MsgBox ExtStr
   If ExtStr = ExStr Or ExtStr = Ex1Str Then
      OutFileName = Left$(InFileName, (Len(InFileName) - 4))
      FileLength = FileLen(InFileName)
      ReDim InBuffer(1 To FileLength)
      InFile = FreeFile
      Close InFile
      Open InFileName For Binary Access Read As InFile
      Call ReadHeader
      Get InFile, , ExpFileLen
      Get InFile, , InBuffer
      Close InFile
      For cnt = 1 To 4
         TempStr = TempStr & Hex$(ExpFileLen(cnt))
      Next cnt
      ExpandedFileLength = CLng(Val("&H" & TempStr))
      OutFile = FreeFile
      Close OutFile
      ReDim OutBuffer(1 To ExpandedFileLength)
      Open OutFileName For Binary Access Write As OutFile
      Call ExpandFile
      Call RunLengthDecode
      Close OutFile
   Else
      MsgBox "File is not compress by ABD SPLAY TREE.", vbCritical, "ERROR"
      Exit Sub
   End If
End Sub

Private Sub RunLengthDecode()
   Dim PrevByte As Byte
   Dim ThisByte As Byte
   Dim x As Long
    
   RLEOutIndex = 1
   For x = 1 To ExpandedFileLength
      ThisByte = OutBuffer(x)
      If ThisByte = 126 Then '"~" Then
         Put OutFile, , PrevByte 'RLEOutBuffer(RLEOutIndex) = PrevByte
                                 'RLEOutIndex = RLEOutIndex + 1
         x = x + 1
      Else
         Put OutFile, , ThisByte  'RLEOutBuffer(RLEOutIndex) = ThisByte
                                  'RLEOutIndex = RLEOutIndex + 1
      End If
      PrevByte = ThisByte
   Next x
    
   FileLength = RLEOutIndex

End Sub

Private Sub RunLengthEncode()
   Dim LastByte As Byte
   Dim ThisByte As Byte
   Dim x As Long
   Dim RepeatCount As Integer

   RLEIndex = 1
   RepeatCount = 0
   For x = 1 To FileLength
      ThisByte = InBuffer(x)
      If LastByte = ThisByte Then
         'If there is only 1 repeating byte (like the e in Speech)
         'then don't encode because it will take 1 extra byte
         If InBuffer(x + 1) <> ThisByte And RepeatCount = 0 Then
            RLEBuffer(RLEIndex) = ThisByte
            RLEIndex = RLEIndex + 1
            LastByte = ThisByte
         Else
            RepeatCount = RepeatCount + 1
        
            'We can only encode up to 254 repeats after that
            'we have to start the new sequence again
            If RepeatCount = 254 Then
               RLEBuffer(RLEIndex) = 126
               RLEIndex = RLEIndex + 1
               RLEBuffer(RLEIndex) = CByte(RepeatCount)
               RLEIndex = RLEIndex + 1
               RepeatCount = 0
               LastByte = AscB("")
            End If
         End If
      Else
         If RepeatCount > 0 Then
            RLEBuffer(RLEIndex) = 126
            RLEIndex = RLEIndex + 1
            RLEBuffer(RLEIndex) = CByte(RepeatCount)
            RLEIndex = RLEIndex + 1
            RepeatCount = 0
         End If
    
            RLEBuffer(RLEIndex) = ThisByte
            RLEIndex = RLEIndex + 1
            LastByte = ThisByte
      End If
   Next x
    
   'If the last byte in InBuffer are repeats
   If RepeatCount > 0 Then
      RLEBuffer(RLEIndex) = 126
      RLEIndex = RLEIndex + 1
      RLEBuffer(RLEIndex) = CByte(RepeatCount)
      RepeatCount = 0
   End If
   RLELength = RLEIndex
End Sub

Sub Main()
   'SplayCompress ("H.txt")
   SplayExpand ("C:\Windows\Profiles\Alfredo Delgado\My Documents\VBProjects\ABDComp\Readme.txt.zap")
   MsgBox "END"
End Sub
