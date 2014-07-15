#include-once

#include "GDIPlusConstants.au3"
#include "StructureConstants.au3"
#include "WinAPI.au3"
#include "WinAPIGdi.au3"
#include "GDIPConstants.au3"

; #INDEX# =======================================================================================================================
; Title .........: GDIPlus
; AutoIt Version : 3.3.12.0
; Language ......: English
; Description ...: Functions that assist with Microsoft Windows GDI+ management.
;                  It enables applications to use graphics and formatted text on both the video display and the printer.
;                  Applications based on the Microsoft Win32 API do not access graphics hardware directly.
;                  Instead, GDI+ interacts with device drivers on behalf of applications.
;                  GDI+ can be used in all Windows-based applications.
;                  GDI+ is new technology that is included in Windows XP and the Windows Server 2003.
; Author ........: Paul Campbell (PaulIA), rover, smashly, monoceres, Malkey, Authenticity
; Modified ......: Gary Frost, UEZ, Eukalyptus, jpm
; Dll ...........: GDIPlus.dll
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
Global $ghGDIPMatrix = 0
Global $GDIP_STATUS = 0
Global $GDIP_ERROR = 0
; ===============================================================================================================================


#Region Image Functions

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageClone
; Description ...: Clones an Image object
; Syntax.........: _GDIPlus_ImageClone($hImage)
; Parameters ....: $hImage - Pointer to an Image object
; Return values .: Success      - Pointer to a new cloned Image object
;                  Failure      - 0 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: After you are done with the object, call _GDIPlus_ImageDispose to release the object resources
; Related .......: _GDIPlus_ImageDispose
; Link ..........; @@MsdnLink@@ GdipCloneImage
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageClone($hImage)
	Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipCloneImage", "hwnd", $hImage, "int*", 0)

	If @error Then Return SetError(@error, @extended, 0)
	$GDIP_STATUS = $aResult[0]
	Return $aResult[2]
EndFunc   ;==>_GDIPlus_ImageClone

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageForceValidation
; Description ...: Forces validation of an image
; Syntax.........: _GDIPlus_ImageForceValidation($hImage)
; Parameters ....: $hImage - Pointer to an Image object
; Return values .: Success      - True if image is correct, False otherwise
;                  Failure      - -1 and sets @error and @extended if DllCall failed
; Remarks .......: This function forces GDI+ to check for image correctness, usually this will be done by GDI+ when drawing the
;                  +image. Even though the function may return False, if the object exists, it's resources should be released
; Related .......: _GDIPlus_ImageDispose
; Link ..........; @@MsdnLink@@ GdipImageForceValidation
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageForceValidation($hImage)
	Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipImageForceValidation", "hwnd", $hImage)

	If @error Then Return SetError(@error, @extended, -1)
	$GDIP_STATUS = $aResult[0]
	Return $aResult[0] = 0
EndFunc   ;==>_GDIPlus_ImageForceValidation

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageGetAllPropertyItems
; Description ...: Gets all the property items (metadata) stored in an Image object
; Syntax.........: _GDIPlus_ImageGetAllPropertyItems($hImage)
; Parameters ....: $hImage - Pointer to an Image object
; Return values .: Success      - Array containing the image property items:
;                  |[0][0] - Number of property items
;                  |[1][0] - Property item 1 identifier (see remarks)
;                  |[1][1] - Property item 1 value size, in bytes
;                  |[1][2] - Property item 1 value type
;                  |[1][1] - Property item 1 value pointer
;                  |[1][0] - Property item n identifier
;                  |[1][1] - Property item n value size, in bytes
;                  |[1][2] - Property item n value type
;                  |[1][1] - Property item n value pointer
;                  Possible property value types are:
;                  |1 - The value pointer points to an array of bytes
;                  |2 - The value pointer points to a null terminated character stringASCII string
;                  |3 - The value pointer points to an array of unsigned shorts
;                  |4 - The value pointer points to an array of unsigned integers
;                  |5 - The value pointer points to an array of unsigned two longs (numerator, denomintor)
;                  |7 - The value pointer points to an array of bytes of any type
;                  |9 - The value pointer points to an array of signed integers
;                  |10- The value pointer points to an array of signed two longs (numerator, denomintor)
;                  Failure      - -1 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
;                  |$GDIP_ERROR:
;                  |	1 - The _GDIPlus_ImageGetPropertySize function failed, $GDIP_STATUS contains the error code
;                  |	2 - The image contains no property items metadata
;                  |	3 - The _GDIPlus_ImageGetAllPropertyItems function failed, $GDIP_STATUS contains the error code
; Remarks .......: The properties item tag identifiers are declared in GDIPConstants.au3, those that start with $GDIP_PROPERTYTAG
;                  +The value size is given in bytes, divide this by the size of the data (4 for integers, 2 for shorts, etc..)
; Related .......: _GDIPlus_ImageGetPropertySize
; Link ..........; @@MsdnLink@@ GdipGetAllPropertyItems
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageGetAllPropertyItems($hImage)
	Local $iI, $iCount, $tBuffer, $pBuffer, $iBuffer, $tPropertyItem, $aSize, $aPropertyItems[1][1], $aResult

	$aSize = _GDIPlus_ImageGetPropertySize($hImage)
	If @error Then Return SetError(@error, @extended, -1)

	If $GDIP_STATUS Then
		$GDIP_ERROR = 1
		Return -1
	ElseIf $aSize[1] = 0 Then
		$GDIP_ERROR = 2
		Return -1
	EndIf

	$iBuffer = $aSize[0]
	$tBuffer = DllStructCreate("byte[" & $iBuffer & "]")
	$pBuffer = DllStructGetPtr($tBuffer)
	$iCount = $aSize[1]

	$aResult = DllCall($__g_hGDIPDll, "uint", "GdipGetAllPropertyItems", "hwnd", $hImage, "uint", $iBuffer, "uint", $iCount, "ptr", $pBuffer)
	If @error Then Return SetError(@error, @extended, -1)

	$GDIP_STATUS = $aResult[0]
	If $GDIP_STATUS Then
		$GDIP_ERROR = 3
		Return -1
	EndIf

	ReDim $aPropertyItems[$iCount + 1][4]
	$aPropertyItems[0][0] = $iCount

	For $iI = 1 To $iCount
		$tPropertyItem = DllStructCreate($tagGDIPPROPERTYITEM, $pBuffer)
		$aPropertyItems[$iI][0] = DllStructGetData($tPropertyItem, "id")
		$aPropertyItems[$iI][1] = DllStructGetData($tPropertyItem, "length")
		$aPropertyItems[$iI][2] = DllStructGetData($tPropertyItem, "type")
		$aPropertyItems[$iI][3] = DllStructGetData($tPropertyItem, "value")
		$pBuffer += DllStructGetSize($tPropertyItem)
	Next

	Return $aPropertyItems
EndFunc   ;==>_GDIPlus_ImageGetAllPropertyItems

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageGetBounds
; Description ...: Gets the bounding rectangle for an image
; Syntax.........: _GDIPlus_ImageGetBounds($hImage)
; Parameters ....: $hImage - Pointer to an Image object
; Return values .: Success      - Array that contains the rectangle coordinates and dimensions:
;                  |[0] - X coordinate of the upper-left corner of the rectangle
;                  |[1] - Y coordinate of the upper-left corner of the rectangle
;                  |[2] - Width of the rectangle
;                  |[3] - Height of the rectangle
;                  Failure      - -1 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: None
; Related .......: None
; Link ..........; @@MsdnLink@@ GdipGetImageBounds
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageGetBounds($hImage)
	Local $tRectF, $pRectF, $iI, $aRectF[4], $aResult

	$tRectF = DllStructCreate($tagGDIPRECTF)
	$pRectF = DllStructGetPtr($tRectF)
	$aResult = DllCall($__g_hGDIPDll, "uint", "GdipGetImageBounds", "hwnd", $hImage, "ptr", $pRectF, "int*", 0)
	If @error Then Return SetError(@error, @extended, -1)

	$GDIP_STATUS = $aResult[0]
	If $GDIP_STATUS Then Return -1

	For $iI = 1 To 4
		$aRectF[$iI - 1] = DllStructGetData($tRectF, $iI)
	Next
	Return $aRectF
EndFunc   ;==>_GDIPlus_ImageGetBounds

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageGetDimension
; Description ...: Gets the width and height of an image
; Syntax.........: _GDIPlus_ImageGetDimension($hImage)
; Parameters ....: $hImage - Pointer to an Image object
; Return values .: Success      - Array that contains the rectangle coordinates and dimensions:
;                  |[0] - Width of the image
;                  |[1] - Height of the image
;                  Failure      - -1 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: None
; Related .......: None
; Link ..........; @@MsdnLink@@ GdipGetImageDimension
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageGetDimension($hImage)
	Local $aSize[2], $aResult

	$aResult = DllCall($__g_hGDIPDll, "uint", "GdipGetImageDimension", "hwnd", $hImage, "float*", 0, "float*", 0)
	If @error Then Return SetError(@error, @extended, -1)

	$GDIP_STATUS = $aResult[0]
	If $GDIP_STATUS Then Return -1

	$aSize[0] = $aResult[2]
	$aSize[1] = $aResult[3]
	Return $aSize
EndFunc   ;==>_GDIPlus_ImageGetDimension

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageGetFrameCount
; Description ...: Gets the number of frames in a specified dimension of an Image object
; Syntax.........: _GDIPlus_ImageGetFrameCount($hImage, $sDimensionID)
; Parameters ....: $hImage   	 - Pointer to an Image object
;                  $sDimensionID - GUID string of the dimension ID
; Return values .: Success      - The number of frames in the specified dimension of the Image object
;                  Failure      - -1 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: GUID constants are declared in GDIPConstants.au3, those that start with $GDIP_FRAMEDIMENSION_*.
; Related .......: _GDIPlus_ImageGetFrameDimensionsList
; Link ..........; @@MsdnLink@@ GdipImageGetFrameCount
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageGetFrameCount($hImage, $sDimensionID)
	Local $tGUID, $pGUID, $aResult

	$tGUID = _WinAPI_GUIDFromString($sDimensionID)
	$pGUID = DllStructGetPtr($tGUID)
	$aResult = DllCall($__g_hGDIPDll, "uint", "GdipImageGetFrameCount", "hwnd", $hImage, "ptr", $pGUID, "uint*", 0)

	If @error Then Return SetError(@error, @extended, -1)

	$GDIP_STATUS = $aResult[0]
	If $GDIP_STATUS Then Return -1
	Return $aResult[3]
EndFunc   ;==>_GDIPlus_ImageGetFrameCount

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageGetFrameDimensionsCount
; Description ...: Gets the number of frame dimensions in an Image object
; Syntax.........: _GDIPlus_ImageGetFrameDimensionsCount($hImage)
; Parameters ....: $hImage - Pointer to an Image object
; Return values .: Success      - The number of frames dimensions in the Image object
;                  Failure      - -1 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: None
; Related .......: _GDIPlus_ImageGetFrameCount, _GDIPlus_ImageGetFrameDimensionsList
; Link ..........; @@MsdnLink@@ GdipImageGetFrameDimensionsCount
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageGetFrameDimensionsCount($hImage)
	Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipImageGetFrameDimensionsCount", "hwnd", $hImage, "uint*", 0)

	If @error Then Return SetError(@error, @extended, -1)

	$GDIP_STATUS = $aResult[0]
	If $GDIP_STATUS Then Return -1
	Return $aResult[2]
EndFunc   ;==>_GDIPlus_ImageGetFrameDimensionsCount

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageGetFrameDimensionsList
; Description ...: Gets the identifiers for the frame dimensions of an Image object
; Syntax.........: _GDIPlus_ImageGetFrameDimensionsList($hImage)
; Parameters ....: $hImage - Pointer to an Image object
; Return values .: Success      - Array of GUID strings that define the frame dimensions identifier:
;                  |[0] - Number of GUID strings
;                  |[1] - GUID string 1
;                  |[2] - GUID string 2
;                  |[n] - GUID string n
;                  Failure      - -1 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
;                  |$GDIP_ERROR:
;                  |	1 - The _GDIPlus_ImageGetFrameDimensionsCount function failed, $GDIP_STATUS contains the error code
;                  |	2 - The image does not contain any frame dimension identifiers
;                  |	3 - The _GDIPlus_ImageGetFrameDimensionsList function failed, $GDIP_STATUS contains the error code
; Remarks .......: None
; Related .......: _GDIPlus_ImageGetFrameCount, _GDIPlus_ImageGetFrameDimensionsCount
; Link ..........; @@MsdnLink@@ GdipImageGetFrameDimensionsList
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageGetFrameDimensionsList($hImage)
	Local $iI, $iCount, $tBuffer, $pBuffer, $aPropertyIDs[1], $aResult

	$iCount = _GDIPlus_ImageGetFrameDimensionsCount($hImage)
	If @error Then Return SetError(@error, @extended, -1)

	If $GDIP_STATUS Then
		$GDIP_ERROR = 1
		Return -1
	ElseIf $iCount = 0 Then
		$GDIP_ERROR = 2
		Return -1
	EndIf

	$tBuffer = DllStructCreate("byte[" & $iCount * 16 & "]")
	$pBuffer = DllStructGetPtr($tBuffer)
	$aResult = DllCall($__g_hGDIPDll, "uint", "GdipImageGetFrameDimensionsList", "hwnd", $hImage, "ptr", $pBuffer, "int", $iCount)
	If @error Then Return SetError(@error, @extended, -1)

	$GDIP_STATUS = $aResult[0]
	If $GDIP_STATUS Then
		$GDIP_ERROR = 3
		Return -1
	EndIf

	ReDim $aPropertyIDs[$iCount + 1]
	$aPropertyIDs[0] = $iCount

	For $iI = 1 To $iCount
		$aPropertyIDs[$iI] = _WinAPI_StringFromGUID($pBuffer)
		$pBuffer += 16
	Next
	Return $aPropertyIDs
EndFunc   ;==>_GDIPlus_ImageGetFrameDimensionsList

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageGetPalette
; Description ...: Gets the color palette of an Image object
; Syntax.........: _GDIPlus_ImageGetPalette($hImage)
; Parameters ....: $hImage - Pointer to an Image object
; Return values .: Success      - $tagGDIPCOLORPALETTE structure.
;                  Failure      - -1 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
;                  |$GDIP_ERROR:
;                  |	1 - The _GDIPlus_ImageGetPaletteSize function failed, $GDIP_STATUS contains the error code
;                  |	2 - The image does not contain a palette
;                  |	3 - The _GDIPlus_ImageGetPalette function failed, $GDIP_STATUS contains the error code
; Remarks .......: None
; Related .......: _GDIPlus_ImageGetPaletteSize, $tagGDIPCOLORPALETTE
; Link ..........; @@MsdnLink@@ GdipGetImagePalette
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageGetPalette($hImage)
	Local $iCount, $iColorPalette, $tColorPalette, $pColorPalette, $aResult

	$iColorPalette = _GDIPlus_ImageGetPaletteSize($hImage)
	If @error Then Return SetError(@error, @extended, -1)

	If $GDIP_STATUS Then
		$GDIP_ERROR = 1
		Return -1
	ElseIf $iColorPalette = 0 Then
		$GDIP_ERROR = 2
		Return -1
	EndIf

	$iCount = ($iColorPalette - 8) / 4
	$tColorPalette = DllStructCreate("uint Flags;uint Count;uint Entries[" & $iCount & "];")
	$pColorPalette = DllStructGetPtr($tColorPalette)
	$aResult = DllCall($__g_hGDIPDll, "uint", "GdipGetImagePalette", "hwnd", $hImage, "ptr", $pColorPalette, "int", $iColorPalette)
	If @error Then Return SetError(@error, @extended, -1)

	$GDIP_STATUS = $aResult[0]
	If $GDIP_STATUS Then Return -1
	Return $tColorPalette
EndFunc   ;==>_GDIPlus_ImageGetPalette

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageGetPaletteSize
; Description ...: Gets the size, in bytes, of the color palette of an Image object
; Syntax.........: _GDIPlus_ImageGetPaletteSize($hImage)
; Parameters ....: $hImage - Pointer to an Image object
; Return values .: Success      - Size, in bytes, of the color palette
;                  Failure      - -1 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: None
; Related .......: _GDIPlus_ImageGetPalette, $tagGDIPCOLORPALETTE
; Link ..........; @@MsdnLink@@ GdipGetImagePaletteSize
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageGetPaletteSize($hImage)
	Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipGetImagePaletteSize", "hwnd", $hImage, "int*", 0)

	If @error Then Return SetError(@error, @extended, -1)

	$GDIP_STATUS = $aResult[0]
	If $GDIP_STATUS Then Return -1
	Return $aResult[2]
EndFunc   ;==>_GDIPlus_ImageGetPaletteSize

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageGetPropertyCount
; Description ...: Gets the number of properties (pieces of metadata) stored in an Image object
; Syntax.........: _GDIPlus_ImageGetPropertyCount($hImage)
; Parameters ....: $hImage - Pointer to an Image object
; Return values .: Success      - Number of property items store in the Image object
;                  Failure      - -1 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: None
; Related .......: _GDIPlus_ImageGetAllPropertyItems, _GDIPlus_ImageGetPropertyIdList
; Link ..........; @@MsdnLink@@ GdipGetPropertyCount
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageGetPropertyCount($hImage)
	Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipGetPropertyCount", "hwnd", $hImage, "uint*", 0)

	If @error Then Return SetError(@error, @extended, -1)

	$GDIP_STATUS = $aResult[0]
	If $GDIP_STATUS Then Return -1
	Return $aResult[2]
EndFunc   ;==>_GDIPlus_ImageGetPropertyCount

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageGetPropertyIdList
; Description ...: Gets a list of the property identifiers used in the metadata of an Image object
; Syntax.........: _GDIPlus_ImageGetPropertyIdList($hImage)
; Parameters ....: $hImage - Pointer to an Image object
; Return values .: Success      - Array of property identifiers:
;                  |[0] - Number of property identifiers
;                  |[1] - Property identifier 1
;                  |[2] - Property identifier 2
;                  |[n] - Property identifier n
;                  Failure      - -1 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
;                  |$GDIP_ERROR:
;                  |	1 - The _GDIPlus_ImageGetPropertyCount function failed, $GDIP_STATUS contains the error code
;                  |	2 - The image does not contain any property items
;                  |	3 - The _GDIPlus_ImageGetPropertyIdList function failed, $GDIP_STATUS contains the error code
; Remarks .......: The property item identifiers are declared in GDIPConstants.au3, those that start with $GDIP_PROPERTYTAGN*
; Related .......: _GDIPlus_ImageGetAllPropertyItems, _GDIPlus_ImageGetPropertyCount, _GDIPlus_ImageGetPropertyItem
; Link ..........; @@MsdnLink@@ GdipGetPropertyIdList
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageGetPropertyIdList($hImage)
	Local $iI, $iCount, $tProperties, $pProperties, $aProperties[1], $aResult

	$iCount = _GDIPlus_ImageGetPropertyCount($hImage)
	If @error Then Return SetError(@error, @extended, -1)

	If $GDIP_STATUS Then
		$GDIP_ERROR = 1
		Return -1
	ElseIf $iCount = 0 Then
		$GDIP_ERROR = 2
		Return -1
	EndIf

	$tProperties = DllStructCreate("uint[" & $iCount & "]")
	$pProperties = DllStructGetPtr($tProperties)
	$aResult = DllCall($__g_hGDIPDll, "uint", "GdipGetPropertyIdList", "hwnd", $hImage, "int", $iCount, "ptr", $pProperties)
	If @error Then Return SetError(@error, @extended, -1)

	$GDIP_STATUS = $aResult[0]
	If $GDIP_STATUS Then
		$GDIP_ERROR = 3
		Return -1
	EndIf

	ReDim $aProperties[$iCount + 1]
	$aProperties[0] = $iCount

	For $iI = 1 To $iCount
		$aProperties[$iI] = DllStructGetData($tProperties, 1, $iI)
	Next
	Return $aProperties
EndFunc   ;==>_GDIPlus_ImageGetPropertyIdList

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageGetPropertyItem
; Description ...: Gets a specified property item (piece of metadata) from an Image object
; Syntax.........: _GDIPlus_ImageGetPropertyItem($hImage, $iPropID)
; Parameters ....: $hImage  - Pointer to an Image object
;                  $iPropID - Identifier of the property item to be retrieved
; Return values .: Success      - $tagGDIPPROPERTYITEM structure containing the property size, type and value pointer
;                  Failure      - -1 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
;                  |$GDIP_ERROR:
;                  |	1 - The _GDIPlus_ImageGetPropertyItemSize function failed, $GDIP_STATUS contains the error code
;                  |	2 - The specified property identifier does not exist in the image
;                  |	3 - The _GDIPlus_ImageGetPropertyItem function failed, $GDIP_STATUS contains the error code
; Remarks .......: None
; Related .......: _GDIPlus_ImageGetPropertyIdList, _GDIPlus_ImageGetPropertyItemSize, $tagGDIPPROPERTYITEM
; Link ..........; @@MsdnLink@@ GdipGetPropertyItem
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageGetPropertyItem($hImage, $iPropID)
	Local $iBuffer,$pBuffer, $tPropertyItem, $aResult
	Global $tBuffer

	$iBuffer = _GDIPlus_ImageGetPropertyItemSize($hImage, $iPropID)
	If @error Then Return SetError(@error, @extended, -1)

	If $GDIP_STATUS Then
		$GDIP_ERROR = 1
		Return -1
	ElseIf $iBuffer = 0 Then
		$GDIP_ERROR = 2
		Return -1
	EndIf

	;ConsoleWrite("====iPropID is " & $iPropID & "====" & "Prop Lengti is " & $iBuffer & "====" & @CRLF )
	$tBuffer = DllStructCreate("byte[" & $iBuffer & "]")
	$pBuffer = DllStructGetPtr($tBuffer)
	$aResult = DllCall($__g_hGDIPDll, "uint", "GdipGetPropertyItem", "hwnd", $hImage, "int", $iPropID, "uint", $iBuffer, "ptr", $pBuffer)
	If @error Then Return SetError(@error, @extended, -1)


	$GDIP_STATUS = $aResult[0]
	If $GDIP_STATUS Then
		$GDIP_ERROR = 3
		Return -1
	EndIf

	$tPropertyItem = DllStructCreate($tagGDIPPROPERTYITEM, $pBuffer)

	$Struct_String = DllStructCreate("CHAR[" & DllStructGetData($tPropertyItem, "length") & "];", DllStructGetData($tPropertyItem, "value") )
    $vPhotoDate = DllStructGetData($Struct_String, 1)
	;ConsoleWrite( @CRLF & "ErrorCODE is " & $GDIP_ERROR & " vPhotoDateStruct id is " & DllStructGetData($tPropertyItem, "id") & @CRLF  & " type is " & DllStructGetData($tPropertyItem, "type") &@CRLF & " length is " & DllStructGetData($tPropertyItem, "length") & @CRLF & "value is " & DllStructGetData($tPropertyItem, "value") & @CRLF & " date is " & $vPhotoDate &@CRLF )
	Return $tPropertyItem
	;Return $_retvalue
EndFunc   ;==>_GDIPlus_ImageGetPropertyItem

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageGetPropertyItemSize
; Description ...: Gets the size, in bytes, of a specified property item of an Image object
; Syntax.........: _GDIPlus_ImageGetPropertyItemSize($hImage, $iPropID)
; Parameters ....: $hImage  - Pointer to an Image object
;                  $iPropID - Identifier of the property item to be retrieved
; Return values .: Success      - $tagGDIPPROPERTYITEM structure containing the property size, type and value pointer
;                  Failure      - -1 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: None
; Related .......: _GDIPlus_ImageGetPropertyIdList, _GDIPlus_ImageGetPropertyItem
; Link ..........; @@MsdnLink@@ GdipGetPropertyItemSize
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageGetPropertyItemSize($hImage, $iPropID)
	Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipGetPropertyItemSize", "hwnd", $hImage, "uint", $iPropID, "uint*", 0)

	If @error Then Return SetError(@error, @extended, -1)

	$GDIP_STATUS = $aResult[0]
	If $GDIP_STATUS Then Return -1
	Return $aResult[3]
EndFunc   ;==>_GDIPlus_ImageGetPropertyItemSize

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageGetPropertySize
; Description ...: Gets the total size, in bytes, and the number of all the property items stored in an Image object
; Syntax.........: _GDIPlus_ImageGetPropertySize($hImage)
; Parameters ....: $hImage  - Pointer to an Image object
; Return values .: Success      - Array containing the total size and the number of property items:
;                  |[0] - Total size, in bytes, of the property items
;                  |[1] - Number of the property items
;                  Failure      - -1 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: None
; Related .......: _GDIPlus_ImageGetPropertyIdList, _GDIPlus_ImageGetPropertyItem
; Link ..........; @@MsdnLink@@ GdipGetPropertyItemSize
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageGetPropertySize($hImage)
	Local $aSize[2], $aResult

	$aResult = DllCall($__g_hGDIPDll, "uint", "GdipGetPropertySize", "hwnd", $hImage, "uint*", 0, "uint*", 0)
	If @error Then Return SetError(@error, @extended, -1)

	$GDIP_STATUS = $aResult[0]
	If $GDIP_STATUS Then Return -1

	$aSize[0] = $aResult[2]
	$aSize[1] = $aResult[3]
	Return $aSize
EndFunc   ;==>_GDIPlus_ImageGetPropertySize

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageGetThumbnail
; Description ...: Gets a thumbnail image from an Image object
; Syntax.........: _GDIPlus_ImageGetThumbnail($hImage[, $iTNWidth = 32[, $iTNHeight = 32]])
; Parameters ....: $hImage    - Pointer to an Image object
;                  $iTNWidth  - Width, in pixels, of the requested thumbnail image
;                  $iTNHeight - Height, in pixels, of the requested thumbnail image
; Return values .: Success      - Pointer to a new thumbnailed Image object
;                  Failure      - 0 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: After you are done with the object, call _GDIPlus_ImageDispose to release the object resources
; Related .......: _GDIPlus_ImageDispose
; Link ..........; @@MsdnLink@@ GdipGetImageThumbnail
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageGetThumbnail($hImage, $iTNWidth = 32, $iTNHeight = 32)
	Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipGetImageThumbnail", "hwnd", $hImage, "uint", $iTNWidth, "uint", $iTNHeight, "int*", 0, "ptr", 0, "ptr", 0)

	If @error Then Return SetError(@error, @extended, 0)

	$GDIP_STATUS = $aResult[0]
	Return $aResult[4]
EndFunc   ;==>_GDIPlus_ImageGetThumbnail

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageLoadFromFileICM
; Description ...: Creates an Image object based on a file. This function uses ICM
; Syntax.........: _GDIPlus_ImageLoadFromFileICM($sFileName)
; Parameters ....: $sFileName - Fully qualified image file name
; Return values .: Success      - Pointer to a new Image object
;                  Failure      - 0 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: After you are done with the object, call _GDIPlus_ImageDispose to release the object resources
; Related .......: _GDIPlus_ImageDispose
; Link ..........; @@MsdnLink@@ GdipLoadImageFromFileICM
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageLoadFromFileICM($sFileName)
	Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipLoadImageFromFileICM", "wstr", $sFileName, "int*", 0)

	If @error Then Return SetError(@error, @extended, 0)
	$GDIP_STATUS = $aResult[0]
	Return $aResult[2]
EndFunc   ;==>_GDIPlus_ImageLoadFromFileICM


; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageLoadFromStreamICM
; Description ...: Creates an Image object based on a stream. This function uses ICM
; Syntax.........: _GDIPlus_ImageLoadFromStreamICM($pStream)
; Parameters ....: $pStream - Pointer to an IStream interface
; Return values .: Success      - Pointer to a new Image object
;                  Failure      - 0 and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: After you are done with the object, call _GDIPlus_ImageDispose to release the object resources
; Related .......: _GDIPlus_ImageDispose
; Link ..........; @@MsdnLink@@ GdipLoadImageFromStreamICM
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageLoadFromStreamICM($pStream)
	Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipLoadImageFromStreamICM", "ptr", $pStream, "int*", 0)

	If @error Then Return SetError(@error, @extended, 0)
	$GDIP_STATUS = $aResult[0]
	Return $aResult[2]
EndFunc   ;==>_GDIPlus_ImageLoadFromStreamICM

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageRemovePropertyItem
; Description ...: Removes a property item (piece of metadata) from an Image object
; Syntax.........: _GDIPlus_ImageRemovePropertyItem($hImage, $iPropID)
; Parameters ....: $hImage  - Pointer to an Image object
;                  $iPropID - Identifier of the property item to be removed
; Return values .: Success      - True
;                  Failure      - False and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: None
; Related .......: _GDIPlus_ImageGetPropertyIdList
; Link ..........; @@MsdnLink@@ GdipRemovePropertyItem
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageRemovePropertyItem($hImage, $iPropID)
	Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipRemovePropertyItem", "hwnd", $hImage, "uint", $iPropID)

	If @error Then Return SetError(@error, @extended, False)
	$GDIP_STATUS = $aResult[0]
	Return $aResult[0] = 0
EndFunc   ;==>_GDIPlus_ImageRemovePropertyItem


; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageSaveAdd
; Description ...: Adds a frame to a file or stream
; Syntax.........: _GDIPlus_ImageSaveAdd($hImage, $pParams)
; Parameters ....: $hImage  - Pointer to an Image object
;                  $pParams - Pointer to a $tagGDIPPENCODERPARAMS structure
; Return values .: Success      - True
;                  Failure      - False and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: Use this Function to save selected frames from a multiple-frame image to another multiple-frame image
; Related .......: _GDIPlus_ImageSaveToFile, _GDIPlus_ImageSaveToStream, _GDIPlus_ImageSelectActiveFrame, $tagGDIPPENCODERPARAMS
; Link ..........; @@MsdnLink@@ GdipSaveAdd
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageSaveAdd($hImage, $pParams)
	Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipSaveAdd", "hwnd", $hImage, "ptr", $pParams)

	If @error Then Return SetError(@error, @extended, False)
	$GDIP_STATUS = $aResult[0]
	Return $aResult[0] = 0
EndFunc   ;==>_GDIPlus_ImageSaveAdd

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageSaveAddImage
; Description ...: Adds a frame to a file or stream
; Syntax.........: _GDIPlus_ImageSaveAddImage($hImage, $hImageNew, $pParams)
; Parameters ....: $hImage    - Pointer to an Image object
;                  $hImageNew - Pointer to an Image object that holds the frame to be added
;                  $pParams   - Pointer to a $tagGDIPPENCODERPARAMS structure
; Return values .: Success      - True
;                  Failure      - False and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: None
; Related .......: _GDIPlus_ImageSaveToFile, _GDIPlus_ImageSaveToStream, _GDIPlus_ImageSelectActiveFrame, $tagGDIPPENCODERPARAMS
; Link ..........; @@MsdnLink@@ GdipSaveAddImage
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageSaveAddImage($hImage, $hImageNew, $pParams)
	Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipSaveAddImage", "hwnd", $hImage, "hwnd", $hImageNew, "ptr", $pParams)

	If @error Then Return SetError(@error, @extended, False)
	$GDIP_STATUS = $aResult[0]
	Return $aResult[0] = 0
EndFunc   ;==>_GDIPlus_ImageSaveAddImage


; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageSelectActiveFrame
; Description ...: Selects a frame in an Image object specified by a dimension and an index
; Syntax.........: _GDIPlus_ImageSelectActiveFrame($hImage, $sDimensionID, $iFrameIndex)
; Parameters ....: $hImage     	  - Pointer to an Image object
;                  $sDimensionID  - GUID string specifies the frame dimension (see remarks):
;                  |$GDIP_FRAMEDIMENSION_TIME - GIF image
;                  |$GDIP_FRAMEDIMENSION_PAGE - TIFF image
;                  $iFrameIndex   - Zero-based index of the frame within the specified frame dimension
; Return values .: Success      - True
;                  Failure      - False and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: Among all the image formats currently supported by GDI+, the only formats that support multiple-frame images
;                  +are GIF and TIFF
; Related .......: _GDIPlus_ImageLoadFromFile, _GDIPlus_ImageLoadFromStream, _GDIPlus_ImageSaveAdd, _GDIPlus_ImageSaveAddImage
; Link ..........; @@MsdnLink@@ GdipImageSelectActiveFrame
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageSelectActiveFrame($hImage, $sDimensionID, $iFrameIndex)
	Local $pGUID, $tGUID, $aResult

	$tGUID = DllStructCreate($tagGUID)
	$pGUID = DllStructGetPtr($tGUID)
	_WinAPI_GUIDFromStringEx($sDimensionID, $pGUID)

	$aResult = DllCall($__g_hGDIPDll, "uint", "GdipImageSelectActiveFrame", "hwnd", $hImage, "ptr", $pGUID, "uint", $iFrameIndex)

	If @error Then Return SetError(@error, @extended, False)
	$GDIP_STATUS = $aResult[0]
	Return $aResult[0] = 0
EndFunc   ;==>_GDIPlus_ImageSelectActiveFrame

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageSetPalette
; Description ...: Sets the color palette of an Image object
; Syntax.........: _GDIPlus_ImageSetPalette($hImage, $pColorPalette)
; Parameters ....: $hImage     	  - Pointer to an Image object
;                  $pColorPalette - Pointer to a $tagGDIPCOLORPALETTE structure that specifies the palette
; Return values .: Success      - True
;                  Failure      - False and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: None
; Related .......: _GDIPlus_ImageGetPalette, $tagGDIPCOLORPALETTE
; Link ..........; @@MsdnLink@@ GdipSetImagePalette
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageSetPalette($hImage, $pColorPalette)
	Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipSetImagePalette", "hwnd", $hImage, "ptr", $pColorPalette)

	If @error Then Return SetError(@error, @extended, False)
	$GDIP_STATUS = $aResult[0]
	Return $aResult[0] = 0
EndFunc   ;==>_GDIPlus_ImageSetPalette

; #FUNCTION# ====================================================================================================================
; Name...........: _GDIPlus_ImageSetPropertyItem
; Description ...: Sets the color palette of an Image object
; Syntax.........: _GDIPlus_ImageSetPropertyItem($hImage, $pPropertyItem)
; Parameters ....: $hImage     	  - Pointer to an Image object
;                  $pPropertyItem - Pointer to a $tagGDIPPROPERTYITEM structure that specifies the property item to be set
; Return values .: Success      - True
;                  Failure      - False and either:
;                  |@error and @extended are set if DllCall failed
;                  |$GDIP_STATUS contains a non zero value specifying the error code
; Remarks .......: If the item already exists, then its contents are updated; otherwise, a new item is added
; Related .......: _GDIPlus_ImageGetPropertyItem, $tagGDIPPROPERTYITEM
; Link ..........; @@MsdnLink@@ GdipSetPropertyItem
; Example .......; No
; ===============================================================================================================================
Func _GDIPlus_ImageSetPropertyItem($hImage, $pPropertyItem)
	Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipSetPropertyItem", "hwnd", $hImage, "ptr", $pPropertyItem)

	If @error Then Return SetError(@error, @extended, False)
	$GDIP_STATUS = $aResult[0]
	Return $aResult[0] = 0
EndFunc   ;==>_GDIPlus_ImageSetPropertyItem

#EndRegion Image Functions
