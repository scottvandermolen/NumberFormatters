<%
	' ASP NumberFormatters Library
	' 
	' Copyright (c) 2023, Scott Vander Molen; some rights reserved.
	' 
	' This work is licensed under a Creative Commons Attribution 4.0 International License.
	' To view a copy of this license, visit https://creativecommons.org/licenses/by/4.0/
	' 
	' @author  Scott Vander Molen
	' @version 2.0
	' @since   2023-11-29

	' Express a number in scientific notation.
	'
	' @param someFloat A floating point value.
	' @return string The floating point value in scientific notation, using E as a separate between the number and exponent.
	function FormatScientific(someFloat)
		if someFloat < 1 then
			' Shift decimal to the right.
			dim newFloat
			newFloat = someFloat
			
			dim moveCounter
			moveCounter = 0
			
			do until newFloat > 1
				newFloat = newFloat * 10
				moveCounter = moveCounter - 1
			loop
			
			FormatScientific = newFloat & "E" & cStr(moveCounter)
		elseif someFloat >= 10 then
			' Shift decimal to the left.
			dim wholeNumber
			wholeNumber = someFloat \ 1
			dim moveDistance
			moveDistance = Len(cStr(wholeNumber)) - 1
			FormatScientific = someFloat / 10 ^ moveDistance & "E" & cstr(moveDistance)
		else
			' Leave decimal alone.
			FormatScientific = cStr(someFloat) & "E0"
		end if
	end function
	
	' Express a cardinal number as an ordinal.
	'
	' @param cardinal A cardinal number.
	' @return string The ordinal equivalent of the number.
	function FormatOrdinal(ByVal cardinal)
		cardinal = CStr(cardinal)
		
		dim suffix
		
		if Right(cardinal, 1) = "1" and Right(cardinal, 2) <> "11" then
			suffix = "st"
		elseif Right(cardinal, 1) = "2" and Right(cardinal, 2) <> "12" then
			suffix = "nd"
		elseif Right(cardinal, 1) = "3" and Right(cardinal, 2) <> "13" then
			suffix = "rd"
		else
			suffix = "th"
		end if
		
		FormatOrdinal = cardinal & suffix
	end function
	
	' Allows modular division of large numbers.
	' Built-in modular division only supports integers.
	'
	' @param dividend The number to be divided.
	' @param divisor The number to divide by.
	' @return float The remainder of the division.
	function ModLarge(dividend, divisor)
		dim quotient
		quotient = Int(dividend / divisor)
		ModLarge = dividend - (quotient * divisor)
	end function
	
	' Spell out a cardinal number.
	' The largest numeric data type, double, supports the range -1.79769313486232E308 to -4.94065645841247E-324 for negative values and 4.94065645841247E-324 to 1.79769313486232E308 for positive values.
	' The Conway-Wechsler system is used for the names of numbers not established by dictionaries.
	' https://googology.fandom.com/wiki/-illion
	'
	' @param cardinal A cardinal number.
	' @return string The spelled-out equivalent of the number.
	function FormatSpellout(cardinal)
		dim result
		dim remainder
		remainder = CDbl(cardinal)
		
		dim NumberNames
		NumberNames = Array (_
			"zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", _
			"ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen" _
		)
		
		dim TenNames
		TenNames = Array(null, "ten", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety")
		
		do
			if remainder <= UBound(NumberNames) then
				result = result & NumberNames(remainder)
				remainder = 0
			elseif remainder < 100 then
				result = result & TenNames(remainder \ 10)
				remainder = remainder mod 10
				
				if remainder > 0 then
					result = result & "-"
				end if
			elseif remainder < 1000 then
				result = result & NumberNames(remainder \ 100) & " hundred "
				remainder = remainder mod 100
			elseif remainder < 10^6 then
				result = result & FormatSpellout(remainder \ 1000) & " thousand "
				remainder = remainder mod 1000
			elseif remainder < 10^9 then
				result = result & FormatSpellout(remainder \ 10^6) & " million "
				remainder = remainder mod 10^6
			' From this point forward, we've exceeded the limits of integers, so we need to use different math.
			elseif remainder < 10^12 then
				result = result & FormatSpellout(Fix(remainder / 10^9)) & " billion "
				remainder = ModLarge(remainder, 10^9)
			elseif remainder < 10^15 then
				result = result & FormatSpellout(Fix(remainder / 10^12)) & " trillion "
				remainder = ModLarge(remainder,  10^12)
			elseif remainder < 10^18 then
				result = result & FormatSpellout(Fix(remainder / 10^15)) & " quadrillion "
				remainder = ModLarge(remainder,  10^15)
			elseif remainder < 10^21 then
				result = result & FormatSpellout(Fix(remainder / 10^18)) & " quintillion "
				remainder = ModLarge(remainder,  10^18)
			elseif remainder < 10^24 then
				result = result & FormatSpellout(Fix(remainder / 10^21)) & " sextillion "
				remainder = ModLarge(remainder,  10^21)
			elseif remainder < 10^27 then
				result = result & FormatSpellout(Fix(remainder / 10^24)) & " septillion "
				remainder = ModLarge(remainder,  10^24)
			elseif remainder < 10^30 then
				result = result & FormatSpellout(Fix(remainder / 10^27)) & " octillion "
				remainder = ModLarge(remainder,  10^27)
			elseif remainder < 10^33 then
				result = result & FormatSpellout(Fix(remainder / 10^30)) & " nonillion "
				remainder = ModLarge(remainder,  10^30)
			elseif remainder < 10^36 then
				result = result & FormatSpellout(Fix(remainder / 10^33)) & " decillion "
				remainder = ModLarge(remainder,  10^33)
			elseif remainder < 10^39 then
				result = result & FormatSpellout(Fix(remainder / 10^36)) & " undecillion "
				remainder = ModLarge(remainder,  10^36)
			elseif remainder < 10^42 then
				result = result & FormatSpellout(Fix(remainder / 10^39)) & " duodecillion "
				remainder = ModLarge(remainder,  10^39)
			elseif remainder < 10^45 then
				result = result & FormatSpellout(Fix(remainder / 10^42)) & " tredecillion "
				remainder = ModLarge(remainder,  10^42)
			elseif remainder < 10^48 then
				result = result & FormatSpellout(Fix(remainder / 10^45)) & " quattuordecillion "
				remainder = ModLarge(remainder,  10^45)
			elseif remainder < 10^51 then
				result = result & FormatSpellout(Fix(remainder / 10^48)) & " quindecillion "
				remainder = ModLarge(remainder,  10^48)
			elseif remainder < 10^54 then
				result = result & FormatSpellout(Fix(remainder / 10^51)) & " sexdecillion "
				remainder = ModLarge(remainder,  10^51)
			elseif remainder < 10^57 then
				result = result & FormatSpellout(Fix(remainder / 10^54)) & " septendecillion "
				remainder = ModLarge(remainder,  10^54)
			elseif remainder < 10^60 then
				result = result & FormatSpellout(Fix(remainder / 10^57)) & " octodecillion "
				remainder = ModLarge(remainder,  10^57)
			elseif remainder < 10^63 then
				result = result & FormatSpellout(Fix(remainder / 10^60)) & " novemdecillion "
				remainder = ModLarge(remainder,  10^60)
			elseif remainder < 10^66 then
				result = result & FormatSpellout(Fix(remainder / 10^63)) & " vigintillion "
				remainder = ModLarge(remainder,  10^63)
			elseif remainder < 10^69 then
				result = result & FormatSpellout(Fix(remainder / 10^66)) & " unvigintillion "
				remainder = ModLarge(remainder,  10^66)
			elseif remainder < 10^72 then
				result = result & FormatSpellout(Fix(remainder / 10^69)) & " duovigintillion "
				remainder = ModLarge(remainder,  10^69)
			elseif remainder < 10^75 then
				result = result & FormatSpellout(Fix(remainder / 10^72)) & " tresvigintillion "
				remainder = ModLarge(remainder,  10^72)
			elseif remainder < 10^78 then
				result = result & FormatSpellout(Fix(remainder / 10^75)) & " quattuorvigintillion "
				remainder = ModLarge(remainder,  10^75)
			elseif remainder < 10^81 then
				result = result & FormatSpellout(Fix(remainder / 10^78)) & " quinvigintillion "
				remainder = ModLarge(remainder,  10^78)
			elseif remainder < 10^84 then
				result = result & FormatSpellout(Fix(remainder / 10^81)) & " sesvigintillion "
				remainder = ModLarge(remainder,  10^81)
			elseif remainder < 10^87 then
				result = result & FormatSpellout(Fix(remainder / 10^84)) & " septemvigintillion "
				remainder = ModLarge(remainder,  10^84)
			elseif remainder < 10^90 then
				result = result & FormatSpellout(Fix(remainder / 10^87)) & " octovigintillion "
				remainder = ModLarge(remainder,  10^87)
			elseif remainder < 10^93 then
				result = result & FormatSpellout(Fix(remainder / 10^90)) & " novemvigintillion "
				remainder = ModLarge(remainder,  10^90)
			elseif remainder < 10^96 then
				result = result & FormatSpellout(Fix(remainder / 10^93)) & " trigintillion "
				remainder = ModLarge(remainder,  10^93)
			elseif remainder < 10^99 then
				result = result & FormatSpellout(Fix(remainder / 10^96)) & " untrigintillion "
				remainder = ModLarge(remainder,  10^96)
			elseif remainder < 10^102 then
				result = result & FormatSpellout(Fix(remainder / 10^99)) & " duotrigintillion "
				remainder = ModLarge(remainder,  10^99)
			elseif remainder < 10^105 then
				result = result & FormatSpellout(Fix(remainder / 10^102)) & " trestrigintillion "
				remainder = ModLarge(remainder,  10^102)
			elseif remainder < 10^108 then
				result = result & FormatSpellout(Fix(remainder / 10^105)) & " quattuortrigintillion "
				remainder = ModLarge(remainder,  10^105)
			elseif remainder < 10^111 then
				result = result & FormatSpellout(Fix(remainder / 10^108)) & " quintrigintillion "
				remainder = ModLarge(remainder,  10^108)
			elseif remainder < 10^114 then
				result = result & FormatSpellout(Fix(remainder / 10^111)) & " sestrigintillion "
				remainder = ModLarge(remainder,  10^111)
			elseif remainder < 10^117 then
				result = result & FormatSpellout(Fix(remainder / 10^114)) & " septentrigintillion "
				remainder = ModLarge(remainder,  10^114)
			elseif remainder < 10^120 then
				result = result & FormatSpellout(Fix(remainder / 10^117)) & " octotrigintillion "
				remainder = ModLarge(remainder,  10^117)
			elseif remainder < 10^123 then
				result = result & FormatSpellout(Fix(remainder / 10^120)) & " noventrigintillion "
				remainder = ModLarge(remainder,  10^120)
			elseif remainder < 10^126 then
				result = result & FormatSpellout(Fix(remainder / 10^123)) & " quadragintillion "
				remainder = ModLarge(remainder,  10^123)
			elseif remainder < 10^129 then
				result = result & FormatSpellout(Fix(remainder / 10^126)) & " unquadragintillion "
				remainder = ModLarge(remainder,  10^126)
			elseif remainder < 10^132 then
				result = result & FormatSpellout(Fix(remainder / 10^129)) & " duoquadragintillion "
				remainder = ModLarge(remainder,  10^129)
			elseif remainder < 10^135 then
				result = result & FormatSpellout(Fix(remainder / 10^132)) & " tresquadragintillion "
				remainder = ModLarge(remainder,  10^132)
			elseif remainder < 10^138 then
				result = result & FormatSpellout(Fix(remainder / 10^135)) & " quattuorquadragintillion "
				remainder = ModLarge(remainder,  10^135)
			elseif remainder < 10^141 then
				result = result & FormatSpellout(Fix(remainder / 10^138)) & " quinquadragintillion "
				remainder = ModLarge(remainder,  10^138)
			elseif remainder < 10^144 then
				result = result & FormatSpellout(Fix(remainder / 10^141)) & " sesquadragintillion "
				remainder = ModLarge(remainder,  10^141)
			elseif remainder < 10^147 then
				result = result & FormatSpellout(Fix(remainder / 10^144)) & " septenquadragintillion "
				remainder = ModLarge(remainder,  10^144)
			elseif remainder < 10^150 then
				result = result & FormatSpellout(Fix(remainder / 10^147)) & " octoquadragintillion "
				remainder = ModLarge(remainder,  10^147)
			elseif remainder < 10^153 then
				result = result & FormatSpellout(Fix(remainder / 10^150)) & " novenquadragintillion "
				remainder = ModLarge(remainder,  10^150)
			elseif remainder < 10^156 then
				result = result & FormatSpellout(Fix(remainder / 10^153)) & " quinquagintillion "
				remainder = ModLarge(remainder,  10^153)
			elseif remainder < 10^159 then
				result = result & FormatSpellout(Fix(remainder / 10^156)) & " unquinquagintillion "
				remainder = ModLarge(remainder,  10^156)
			elseif remainder < 10^162 then
				result = result & FormatSpellout(Fix(remainder / 10^159)) & " duoquinquagintillion "
				remainder = ModLarge(remainder,  10^159)
			elseif remainder < 10^165 then
				result = result & FormatSpellout(Fix(remainder / 10^162)) & " tresquinquagintillion "
				remainder = ModLarge(remainder,  10^162)
			elseif remainder < 10^168 then
				result = result & FormatSpellout(Fix(remainder / 10^165)) & " quattuorquinquagintillion "
				remainder = ModLarge(remainder,  10^165)
			elseif remainder < 10^171 then
				result = result & FormatSpellout(Fix(remainder / 10^168)) & " quinquinquagintillion "
				remainder = ModLarge(remainder,  10^168)
			elseif remainder < 10^174 then
				result = result & FormatSpellout(Fix(remainder / 10^171)) & " sesquinquagintillion "
				remainder = ModLarge(remainder,  10^171)
			elseif remainder < 10^177 then
				result = result & FormatSpellout(Fix(remainder / 10^174)) & " septenquinquagintillion "
				remainder = ModLarge(remainder,  10^174)
			elseif remainder < 10^180 then
				result = result & FormatSpellout(Fix(remainder / 10^177)) & " octoquinquagintillion "
				remainder = ModLarge(remainder,  10^177)
			elseif remainder < 10^183 then
				result = result & FormatSpellout(Fix(remainder / 10^180)) & " novenquinquagintillion "
				remainder = ModLarge(remainder,  10^180)
			elseif remainder < 10^186 then
				result = result & FormatSpellout(Fix(remainder / 10^183)) & " sexagintillion "
				remainder = ModLarge(remainder,  10^183)
			elseif remainder < 10^189 then
				result = result & FormatSpellout(Fix(remainder / 10^186)) & " unsexagintillion "
				remainder = ModLarge(remainder,  10^186)
			elseif remainder < 10^192 then
				result = result & FormatSpellout(Fix(remainder / 10^189)) & " duosexagintillion "
				remainder = ModLarge(remainder,  10^189)
			elseif remainder < 10^195 then
				result = result & FormatSpellout(Fix(remainder / 10^192)) & " tresexagintillion "
				remainder = ModLarge(remainder,  10^192)
			elseif remainder < 10^198 then
				result = result & FormatSpellout(Fix(remainder / 10^195)) & " quattuorsexagintillion "
				remainder = ModLarge(remainder,  10^195)
			elseif remainder < 10^201 then
				result = result & FormatSpellout(Fix(remainder / 10^198)) & " quinsexagintillion "
				remainder = ModLarge(remainder,  10^198)
			elseif remainder < 10^204 then
				result = result & FormatSpellout(Fix(remainder / 10^201)) & " sesexagintillion "
				remainder = ModLarge(remainder,  10^201)
			elseif remainder < 10^207 then
				result = result & FormatSpellout(Fix(remainder / 10^204)) & " septensexagintillion "
				remainder = ModLarge(remainder,  10^204)
			elseif remainder < 10^210 then
				result = result & FormatSpellout(Fix(remainder / 10^207)) & " octosexagintillion "
				remainder = ModLarge(remainder,  10^207)
			elseif remainder < 10^213 then
				result = result & FormatSpellout(Fix(remainder / 10^210)) & " novensexagintillion "
				remainder = ModLarge(remainder,  10^210)
			elseif remainder < 10^216 then
				result = result & FormatSpellout(Fix(remainder / 10^213)) & " septuagintillion "
				remainder = ModLarge(remainder,  10^213)
			elseif remainder < 10^219 then
				result = result & FormatSpellout(Fix(remainder / 10^216)) & " unseptuagintillion "
				remainder = ModLarge(remainder,  10^216)
			elseif remainder < 10^222 then
				result = result & FormatSpellout(Fix(remainder / 10^219)) & " duoseptuagintillion "
				remainder = ModLarge(remainder,  10^219)
			elseif remainder < 10^225 then
				result = result & FormatSpellout(Fix(remainder / 10^222)) & " treseptuagintillion "
				remainder = ModLarge(remainder,  10^222)
			elseif remainder < 10^228 then
				result = result & FormatSpellout(Fix(remainder / 10^225)) & " quattuorseptuagintillion "
				remainder = ModLarge(remainder,  10^225)
			elseif remainder < 10^231 then
				result = result & FormatSpellout(Fix(remainder / 10^228)) & " quinseptuagintillion "
				remainder = ModLarge(remainder,  10^228)
			elseif remainder < 10^234 then
				result = result & FormatSpellout(Fix(remainder / 10^231)) & " seseptuagintillion "
				remainder = ModLarge(remainder,  10^231)
			elseif remainder < 10^237 then
				result = result & FormatSpellout(Fix(remainder / 10^234)) & " septenseptuagintillion "
				remainder = ModLarge(remainder,  10^234)
			elseif remainder < 10^240 then
				result = result & FormatSpellout(Fix(remainder / 10^237)) & " octoseptuagintillion "
				remainder = ModLarge(remainder,  10^237)
			elseif remainder < 10^243 then
				result = result & FormatSpellout(Fix(remainder / 10^240)) & " novenseptuagintillion "
				remainder = ModLarge(remainder,  10^240)
			elseif remainder < 10^246 then
				result = result & FormatSpellout(Fix(remainder / 10^243)) & " octogintillion "
				remainder = ModLarge(remainder,  10^243)
			elseif remainder < 10^249 then
				result = result & FormatSpellout(Fix(remainder / 10^246)) & " unoctogintillion "
				remainder = ModLarge(remainder,  10^246)
			elseif remainder < 10^252 then
				result = result & FormatSpellout(Fix(remainder / 10^249)) & " duooctogintillion "
				remainder = ModLarge(remainder,  10^249)
			elseif remainder < 10^255 then
				result = result & FormatSpellout(Fix(remainder / 10^252)) & " tresoctogintillion "
				remainder = ModLarge(remainder,  10^252)
			elseif remainder < 10^258 then
				result = result & FormatSpellout(Fix(remainder / 10^255)) & " quattuoroctogintillion "
				remainder = ModLarge(remainder,  10^255)
			elseif remainder < 10^261 then
				result = result & FormatSpellout(Fix(remainder / 10^258)) & " quinoctogintillion "
				remainder = ModLarge(remainder,  10^258)
			elseif remainder < 10^264 then
				result = result & FormatSpellout(Fix(remainder / 10^261)) & " sexoctogintillion "
				remainder = ModLarge(remainder,  10^261)
			elseif remainder < 10^267 then
				result = result & FormatSpellout(Fix(remainder / 10^264)) & " septemoctogintillion "
				remainder = ModLarge(remainder,  10^264)
			elseif remainder < 10^270 then
				result = result & FormatSpellout(Fix(remainder / 10^267)) & " octooctogintillion "
				remainder = ModLarge(remainder,  10^267)
			elseif remainder < 10^273 then
				result = result & FormatSpellout(Fix(remainder / 10^270)) & " novemoctogintillion "
				remainder = ModLarge(remainder,  10^270)
			elseif remainder < 10^276 then
				result = result & FormatSpellout(Fix(remainder / 10^273)) & " nonagintillion "
				remainder = ModLarge(remainder,  10^273)
			elseif remainder < 10^279 then
				result = result & FormatSpellout(Fix(remainder / 10^276)) & " unnonagintillion "
				remainder = ModLarge(remainder,  10^276)
			elseif remainder < 10^282 then
				result = result & FormatSpellout(Fix(remainder / 10^279)) & " duononagintillion "
				remainder = ModLarge(remainder,  10^279)
			elseif remainder < 10^285 then
				result = result & FormatSpellout(Fix(remainder / 10^282)) & " trenonagintillion "
				remainder = ModLarge(remainder,  10^282)
			elseif remainder < 10^288 then
				result = result & FormatSpellout(Fix(remainder / 10^285)) & " quattuornonagintillion "
				remainder = ModLarge(remainder,  10^285)
			elseif remainder < 10^291 then
				result = result & FormatSpellout(Fix(remainder / 10^288)) & " quinnonagintillion "
				remainder = ModLarge(remainder,  10^288)
			elseif remainder < 10^294 then
				result = result & FormatSpellout(Fix(remainder / 10^291)) & " senonagintillion "
				remainder = ModLarge(remainder,  10^291)
			elseif remainder < 10^297 then
				result = result & FormatSpellout(Fix(remainder / 10^294)) & " septenonagintillion "
				remainder = ModLarge(remainder,  10^294)
			elseif remainder < 10^300 then
				result = result & FormatSpellout(Fix(remainder / 10^297)) & " octononagintillion "
				remainder = ModLarge(remainder,  10^297)
			elseif remainder < 10^303 then
				result = result & FormatSpellout(Fix(remainder / 10^300)) & " novenonagintillion "
				remainder = ModLarge(remainder,  10^300)
			elseif remainder < 10^306 then
				result = result & FormatSpellout(Fix(remainder / 10^303)) & " centillion "
				remainder = ModLarge(remainder,  10^303)
			elseif remainder < 10^309 then
				result = result & FormatSpellout(Fix(remainder / 10^306)) & " uncentillion "
				remainder = ModLarge(remainder,  10^306)
			elseif remainder < 10^312 then
				result = result & FormatSpellout(Fix(remainder / 10^309)) & " duocentillion "
				remainder = ModLarge(remainder,  10^309)
			elseif remainder < 10^315 then
				result = result & FormatSpellout(Fix(remainder / 10^312)) & " trescentillion "
				remainder = ModLarge(remainder,  10^312)
			elseif remainder < 10^318 then
				result = result & FormatSpellout(Fix(remainder / 10^315)) & " quattuorcentillion "
				remainder = ModLarge(remainder,  10^315)
			elseif remainder < 10^321 then
				result = result & FormatSpellout(Fix(remainder / 10^318)) & " quincentillion "
				remainder = ModLarge(remainder,  10^318)
			elseif remainder < 10^324 then
				result = result & FormatSpellout(Fix(remainder / 10^321)) & " sexcentillion "
				remainder = ModLarge(remainder,  10^321)
			' Stopping here because larger numbers are outside the range of the largest datatype.
			else
				result = result & FormatSpellout(Fix(remainder / 10^324)) & " septencentillion "
				remainder = ModLarge(remainder,  10^324)
			end if
		loop until remainder = 0
		
		FormatSpellout = result
	end function
	
	' Spell out an ordinal number.
	'
	' @param cardinal A cardinal number.
	' @return string The spelled-out ordinal equivalent of the number.
	function FormatSpelloutOrdinal(cardinal)
		dim result
		
		dim NumberNames
		NumberNames = Array (_
			"zeroth", "first", "second", "third", "fourth", "fifth", "sixth", "seventh", "eighth", "ninth", _
			"tenth", "eleventh", "twelfth", "thirteenth", "fourteenth", "fifteenth", "sixteenth", "seventeenth", "eighteenth", "nineteenth" _
		)
		
		if cardinal <= UBound(NumberNames) then
			result = result & NumberNames(cardinal)
		else
			' Instead of reinventing the wheel, we're leveraging the FormatSpellout() function to do most of the work for us, and then we'll handle the differences afterwards.
			result = FormatSpellout(cardinal)
			
			if cardinal mod 100 < 20 then
				' Convert one, two, three, etc. to first, second, third, etc.
				
				' Remove last word from string.
				dim wordArray
				wordArray = Split(result, " ")
				redim preserve wordArray(ubound(wordArray) - 1)
				result = Join(wordArray, " ")
				
				' Add new last word to string.
				result = result & " " & NumberNames(cardinal mod 100)
			elseif cardinal mod 10 = 0 then
				' Convert twenty, thirty, forty, etc. to twentieth, thirtieth, fortieth, etc.
				result = Left(result, Len(result) - 1) & "ieth"
			else
				' 
				result = Left(result, Instrrev(result, "-")) & NumberNames(cardinal mod 10)
			end if
		end if
		
		FormatSpelloutOrdinal = result
	end function
%>