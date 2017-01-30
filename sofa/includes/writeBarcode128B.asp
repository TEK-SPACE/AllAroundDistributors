<%

' The following subroutine can be used to generate Code 128 B barcodes.
' Author: Remus Oproiu  
Function WriteBarcode128B(ByVal InputString)

	Const MinValidAscii = 32
	Const MaxValidAscii = 126

	Const Ascii_DEL   = 195
	Const Ascii_FNC3  = 196
	Const Ascii_FNC2  = 197
	Const Ascii_SHIFT = 198
	Const Ascii_CODEC = 199
	Const Ascii_FNC4  = 200
	Const Ascii_CODEA = 201
	Const Ascii_FNC1  = 202
	Const Ascii_STARTA = 203
	Const Ascii_STARTB = 204
	Const Ascii_STARTC = 205
	Const Ascii_STOP = 206

	Dim CharValue(255)
	CharValue(32)  = 0		' <SPACE>
	CharValue(33)  = 1		' !
	CharValue(34)  = 2		' "
	CharValue(35)  = 3		' #
	CharValue(36)  = 4		' $
	CharValue(37)  = 5		' %
	CharValue(38)  = 6		' &
	CharValue(39)  = 7		' '
	CharValue(40)  = 8		' (
	CharValue(41)  = 9		' )
	CharValue(42)  = 10		' *
	CharValue(43)  = 11		' +
	CharValue(44)  = 12		' ,
	CharValue(45)  = 13		' -
	CharValue(46)  = 14		' .
	CharValue(47)  = 15		' /
	CharValue(48)  = 16		' 0
	CharValue(49)  = 17		' 1
	CharValue(50)  = 18		' 2
	CharValue(51)  = 19		' 3
	CharValue(52)  = 20		' 4
	CharValue(53)  = 21		' 5
	CharValue(54)  = 22		' 6
	CharValue(55)  = 23		' 7
	CharValue(56)  = 24		' 8
	CharValue(57)  = 25		' 9
	CharValue(58)  = 26		' :
	CharValue(59)  = 27		' ;
	CharValue(60)  = 28		' <
	CharValue(61)  = 29		'  =
	CharValue(62)  = 30		' >
	CharValue(63)  = 31		' ?
	CharValue(64)  = 32		' @
	CharValue(65)  = 33		' A
	CharValue(66)  = 34		' B
	CharValue(67)  = 35		' C
	CharValue(68)  = 36		' D
	CharValue(69)  = 37		' E
	CharValue(70)  = 38		' F
	CharValue(71)  = 39		' G
	CharValue(72)  = 40		' H
	CharValue(73)  = 41		' I
	CharValue(74)  = 42		' J
	CharValue(75)  = 43		' K
	CharValue(76)  = 44		' L
	CharValue(77)  = 45		' M
	CharValue(78)  = 46		' N
	CharValue(79)  = 47		' O
	CharValue(80)  = 48		' P
	CharValue(81)  = 49		' Q
	CharValue(82)  = 50		' R
	CharValue(83)  = 51		' S
	CharValue(84)  = 52		' T
	CharValue(85)  = 53		' U
	CharValue(86)  = 54		' V
	CharValue(87)  = 55		' W
	CharValue(88)  = 56		' X
	CharValue(89)  = 57		' Y
	CharValue(90)  = 58		' Z
	CharValue(91)  = 59		' [
	CharValue(92)  = 60		' /
	CharValue(93)  = 61		' ]
	CharValue(94)  = 62		' ^
	CharValue(95)  = 63		' _
	CharValue(96)  = 64		' `
	CharValue(97)  = 65		' a
	CharValue(98)  = 66		' b
	CharValue(99)  = 67		' c
	CharValue(100) = 68		' d
	CharValue(101) = 69		' e
	CharValue(102) = 70		' f
	CharValue(103) = 71		' g
	CharValue(104) = 72		' h
	CharValue(105) = 73		' i
	CharValue(106) = 74		' j
	CharValue(107) = 75		' k
	CharValue(108) = 76		' l
	CharValue(109) = 77		' m
	CharValue(110) = 78		' n
	CharValue(111) = 79		' o
	CharValue(112) = 80		' p
	CharValue(113) = 81		' q
	CharValue(114) = 82		' r
	CharValue(115) = 83		' s
	CharValue(116) = 84		' t
	CharValue(117) = 85		' u
	CharValue(118) = 86		' v
	CharValue(119) = 87		' w
	CharValue(120) = 88		' x
	CharValue(121) = 89		' y
	CharValue(122) = 90		' z
	CharValue(123) = 91		' {
	CharValue(124) = 92		' |
	CharValue(125) = 93		' }
	CharValue(126) = 94		' ~
	CharValue(Ascii_DEL)   = 95
	CharValue(Ascii_FNC3)  = 96
	CharValue(Ascii_FNC2)  = 97
	CharValue(Ascii_SHIFT) = 98
	CharValue(Ascii_CODEC) = 99
	CharValue(Ascii_FNC4)  = 100
	CharValue(Ascii_CODEA) = 101
	CharValue(Ascii_FNC1)  = 102
	CharValue(Ascii_STARTA) = 103
	CharValue(Ascii_STARTB) = 104
	CharValue(Ascii_STARTC) = 105

	' Encode the input string

	InputString = Trim(InputString)

	Dim CheckDigitValue, CharPos, CharAscii, InvalidCharsFound

	InvalidCharsFound = false
	CheckDigitValue = CharValue(Ascii_STARTB)
	For CharPos = 1 To Len(InputString)
	    CharAscii = Asc(Mid(InputString, CharPos, 1))
	    if (CharAscii < MinValidAscii) OR (CharAscii > MaxValidAscii) then
			CharAscii = Asc("?")
			InvalidCharsFound = true
		end if
	    CheckDigitValue = CheckDigitValue + (CharValue(CharAscii) * CharPos)
	Next
	CheckDigitValue = (CheckDigitValue Mod 103)

	Dim CheckDigitAscii
	if CheckDigitValue < 95 then
		CheckDigitAscii = CheckDigitValue + 32
	else
		CheckDigitAscii = CheckDigitValue + 100
	end if

	Dim OutputString
	OutputString = Chr(Ascii_STARTB) & InputString & Chr(CheckDigitAscii) & Chr(Ascii_STOP)

	' Write the bars

	' Each Code 128 character consists of three black and three white bars "BWBWBW"
	' (except for the stop character which has an extra black bar).
	' The bars have four different widths (1, 2, 3 and 4).

	Dim BarcodePattern(255)
	BarcodePattern(32)  = "212222"		' <SPACE>
	BarcodePattern(33)  = "222122"		' !
	BarcodePattern(34)  = "222221"		' "
	BarcodePattern(35)  = "121223"		' #
	BarcodePattern(36)  = "121322"		' $
	BarcodePattern(37)  = "131222"		' %
	BarcodePattern(38)  = "122213"		' &
	BarcodePattern(39)  = "122312"		' '
	BarcodePattern(40)  = "132212"		' (
	BarcodePattern(41)  = "221213"		' )
	BarcodePattern(42)  = "221312"		' *
	BarcodePattern(43)  = "231212"		' +
	BarcodePattern(44)  = "112232"		' ,
	BarcodePattern(45)  = "122132"		' -
	BarcodePattern(46)  = "122231"		' .
	BarcodePattern(47)  = "113222"		' /
	BarcodePattern(48)  = "123122"		' 0
	BarcodePattern(49)  = "123221"		' 1
	BarcodePattern(50)  = "223211"		' 2
	BarcodePattern(51)  = "221132"		' 3
	BarcodePattern(52)  = "221231"		' 4
	BarcodePattern(53)  = "213212"		' 5
	BarcodePattern(54)  = "223112"		' 6
	BarcodePattern(55)  = "312131"		' 7
	BarcodePattern(56)  = "311222"		' 8
	BarcodePattern(57)  = "321122"		' 9
	BarcodePattern(58)  = "321221"		' :
	BarcodePattern(59)  = "312212"		' ;
	BarcodePattern(60)  = "322112"		' <
	BarcodePattern(61)  = "322211"		'  =
	BarcodePattern(62)  = "212123"		' >
	BarcodePattern(63)  = "212321"		' ?
	BarcodePattern(64)  = "232121"		' @
	BarcodePattern(65)  = "111323"		' A
	BarcodePattern(66)  = "131123"		' B
	BarcodePattern(67)  = "131321"		' C
	BarcodePattern(68)  = "112313"		' D
	BarcodePattern(69)  = "132113"		' E
	BarcodePattern(70)  = "132311"		' F
	BarcodePattern(71)  = "211313"		' G
	BarcodePattern(72)  = "231113"		' H
	BarcodePattern(73)  = "231311"		' I
	BarcodePattern(74)  = "112133"		' J
	BarcodePattern(75)  = "112331"		' K
	BarcodePattern(76)  = "132131"		' L
	BarcodePattern(77)  = "113123"		' M
	BarcodePattern(78)  = "113321"		' N
	BarcodePattern(79)  = "133121"		' O
	BarcodePattern(80)  = "313121"		' P
	BarcodePattern(81)  = "211331"		' Q
	BarcodePattern(82)  = "231131"		' R
	BarcodePattern(83)  = "213113"		' S
	BarcodePattern(84)  = "213311"		' T
	BarcodePattern(85)  = "213131"		' U
	BarcodePattern(86)  = "311123"		' V
	BarcodePattern(87)  = "311321"		' W
	BarcodePattern(88)  = "331121"		' X
	BarcodePattern(89)  = "312113"		' Y
	BarcodePattern(90)  = "312311"		' Z
	BarcodePattern(91)  = "332111"		' [
	BarcodePattern(92)  = "314111"		' /
	BarcodePattern(93)  = "221411"		' ]
	BarcodePattern(94)  = "431111"		' ^
	BarcodePattern(95)  = "111224"		' _
	BarcodePattern(96)  = "111422"		' `
	BarcodePattern(97)  = "121124"		' a
	BarcodePattern(98)  = "121421"		' b
	BarcodePattern(99)  = "141122"		' c
	BarcodePattern(100) = "141221"		' d
	BarcodePattern(101) = "112214"		' e
	BarcodePattern(102) = "112412"		' f
	BarcodePattern(103) = "122114"		' g
	BarcodePattern(104) = "122411"		' h
	BarcodePattern(105) = "142112"		' i
	BarcodePattern(106) = "142211"		' j
	BarcodePattern(107) = "241211"		' k
	BarcodePattern(108) = "221114"		' l
	BarcodePattern(109) = "413111"		' m
	BarcodePattern(110) = "241112"		' n
	BarcodePattern(111) = "134111"		' o
	BarcodePattern(112) = "111242"		' p
	BarcodePattern(113) = "121142"		' q
	BarcodePattern(114) = "121241"		' r
	BarcodePattern(115) = "114212"		' s
	BarcodePattern(116) = "124112"		' t
	BarcodePattern(117) = "124211"		' u
	BarcodePattern(118) = "411212"		' v
	BarcodePattern(119) = "421112"		' w
	BarcodePattern(120) = "421211"		' x
	BarcodePattern(121) = "212141"		' y
	BarcodePattern(122) = "214121"		' z
	BarcodePattern(123) = "412121"		' {
	BarcodePattern(124) = "111143"		' |
	BarcodePattern(125) = "111341"		' }
	BarcodePattern(126) = "131141"		' ~
	BarcodePattern(Ascii_DEL)    = "114113"
	BarcodePattern(Ascii_FNC3)   = "114311"
	BarcodePattern(Ascii_FNC2)   = "411113"
	BarcodePattern(Ascii_SHIFT)  = "411311"
	BarcodePattern(Ascii_CODEC)  = "113141"
	BarcodePattern(Ascii_FNC4)   = "114131"
	BarcodePattern(Ascii_CODEA)  = "311141"
	BarcodePattern(Ascii_FNC1)   = "411131"
	BarcodePattern(Ascii_STARTA) = "211412"
	BarcodePattern(Ascii_STARTB) = "211214"
	BarcodePattern(Ascii_STARTC) = "211232"
	BarcodePattern(Ascii_STOP)   = "2331112"

	Dim OutputPattern

	OutputPattern = ""
	for CharPos = 1 to Len(OutputString)
		OutputPattern = OutputPattern & BarcodePattern(Asc(Mid(OutputString, CharPos, 1)))
	next

	Const BarWidth = 1
	Const BarHeight = 35

	for CharPos = 1 to Len(OutputPattern) - 1 step 2
		Response.Write("<img src='images/Black.gif' width=" & (BarWidth * CInt(Mid(OutputPattern, CharPos, 1))) & " height=" & BarHeight & " border=0>")
		Response.Write("<img src='images/Blank.gif' width=" & (BarWidth * CInt(Mid(OutputPattern, CharPos + 1, 1))) & " height=" & BarHeight & " border=0>")
	next
	if CharPos = Len(OutputPattern) then
		Response.Write("<img src='images/Black.gif' width=" & (BarWidth * CInt(Mid(OutputPattern, CharPos, 1))) & " height=" & BarHeight & " border=0>")
	end if

	' Exit function

	if InvalidCharsFound then
		WriteBarcode128B = false
	else
		WriteBarcode128B = true
	end if

End Function

%>
