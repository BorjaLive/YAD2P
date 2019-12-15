#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Batch PDF converter

#ce ----------------------------------------------------------------------------

#include <File.au3>
#include <Word.au3>

__DOC2PDF("C:\Users\borja\Desktop\1º Trimestre\")

#Region INTERFACE
Func __DOC2PDF($source, $destination = 0)
	If NOT IsArray($source) Then
		If StringInStr($source, ".doc") == 0 Then
			$files = _FileListToArray($source, "*.doc*", $FLTA_FILES , True)
		Else
			Local $files[2]
			$files[0] = 1
			$files[1] = $source
		EndIf
	Else
		If UBound($source)+1 <> $source[0] OR $source[0] == 0 Then Return SetError(1, "Invalid array formating")
		$files = $source
	EndIf
	If $destination == 0 Then
		$destination = StringMid($files[1], 1, StringInStr(StringReplace($files[1], "\", "/"), "/", 0, -1)-1)
	Else
		If Not FileExists($destination) Then Return SetError(2, "Destination folder does not exist.")
	EndIf

	DirCreate($destination&"\pdf")

	$oWord = _Word_Create()
	If @error Then Return SetError(3, "Exception invoking MSOFFICE instance. Returned code: "&@error&". "&@CRLF&@extended)

	For $i = 1 To $files[0]
		$oDoc = _Word_DocOpen($oWord, $files[$i], Default, Default, True)
		If @error Then Return SetError(4, "Exception opening doc file. Returned code: "&@error&". "&@CRLF&@extended)

		_Word_DocExport($oDoc, $destination&"\pdf\"&StringReplace(StringReplace(StringMid($files[$i], StringInStr(StringReplace($files[$i], "\", "/"), "/", 0, -1)+1), ".docx", ".pdf"), ".doc", ".pdf"))
		If @error Then Return SetError(5, "Exception exporting PDF. Returned code: "&@error&". "&@CRLF&@extended)

		_Word_DocClose($oDoc)
		If @error Then Return SetError(6, "Exception closing DOC. Returned code: "&@error&". "&@CRLF&@extended)
	Next

	_Word_Quit($oWord)
	If @error Then Return SetError(7, "Exception quitting Word. Returned code: "&@error&". "&@CRLF&@extended)
EndFunc
#EndRegion
