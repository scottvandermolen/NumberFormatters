<%@ CodePage=65001 Language="VBScript"%>
<% Option Explicit %>
<!--#include file="NumberFormatters.lib.asp"-->
<%
	' ASP Number Formatters Unit Tests
	' 
	' Copyright (c) 2023, Scott Vander Molen; some rights reserved.
	' 
	' This work is licensed under a Creative Commons Attribution 4.0 International License.
	' To view a copy of this license, visit https://creativecommons.org/licenses/by/4.0/
	' 
	' @author  Scott Vander Molen
	' @version 2.0
	' @since   2023-11-29

	' Ensure that UTF-8 encoding is used instead of Windows-1252
	Session.CodePage = 65001
	Response.CodePage = 65001
	Response.CharSet = "UTF-8"
	Response.ContentType = "text/html"
	
	' Framework for running tests and building result strings.
	' 
	' @param input The specified value.
	' @param expected The expected value.
	' @param actual The actual value.
	' @return string The results of the test.
	function testFramework(input, expected, actual)
		dim result
		dim resultText
		
		if actual = expected then
			result = true
			resultText = "successful"
		else
			result = false
			resultText = "failed"
		end if
		
		dim returnString
		returnString = "Input:     " & input & vbCrLf & _
			"Expected:  " & expected & vbCrLf & _
			"Actual:    " & actual & vbCrLf & _
			"Result:    Test " & resultText &  "!" & vbCrLf & vbCrLf
		
		testFramework = returnString
	end function
	
	' Create an HTML container for our output.
	Response.Write "<!DOCTYPE html>" & vbCrLf
	Response.Write "<html lang=""en"">" & vbCrLf
	Response.Write "<meta http-equiv=""Content-Type"" content=""text/html;charset=UTF-8"" />" & vbCrLf
	Response.Write "<body>" & vbCrLf
	
	' Display code header
	Response.Write "<pre>"
	Response.Write "/***************************************************************************************\" & vbCrLf
	Response.Write "| ASP Number Formatters Unit Tests                                                      |" & vbCrLf
	Response.Write "|                                                                                       |" & vbCrLf
	Response.Write "| Copyright (c) 2023, Scott Vander Molen; some rights reserved.                         |" & vbCrLf
	Response.Write "|                                                                                       |" & vbCrLf
	Response.Write "| This work is licensed under a Creative Commons Attribution 4.0 International License. |" & vbCrLf
	Response.Write "| To view a copy of this license, visit https://creativecommons.org/licenses/by/4.0/    |" & vbCrLf
	Response.Write "|                                                                                       |" & vbCrLf
	Response.Write "\***************************************************************************************/" & vbCrLf
	Response.Write "</pre>"
	
	' Run unit tests
	Response.Write "<pre>"
	
	Response.Write "Unit Test: FormatScientific()" & vbCrLf & testFramework(1234567891, "1.234567891E9", FormatScientific(1234567891))
	Response.Write "Unit Test: FormatOrdinal()" & vbCrLf & testFramework(1234567891, "1234567891st", FormatOrdinal(1234567891))
	Response.Write "Unit Test: FormatSpellout()" & vbCrLf & testFramework(1234567891, "one billion two hundred thirty-four million five hundred sixty-seven thousand eight hundred ninety-one", FormatSpellout(1234567891))
	Response.Write "Unit Test: FormatSpelloutOrdinal()" & vbCrLf & testFramework(1234567891, "one billion two hundred thirty-four million five hundred sixty-seven thousand eight hundred ninety-first", FormatSpelloutOrdinal(1234567891))
	
	Response.Write "</pre>" & vbCrLf
	
	' Close the HTML container.
	Response.Write "</body>" & vbCrLf
	Response.Write "</html>"
%>