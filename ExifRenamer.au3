#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiStatusBar.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>
#include <ProgressConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <File.au3>
#include <GDIPlus.au3>
#include <array.au3>

#include "Libs/GDIPlus_Image.au3"
#include "FileSelectFolder.au3"


;#include "Libs\GDIPConstants.au3"

#Region Global Variables
Global $gFileLocation = ""
Global $gFileCount = 0
Global $gRenamingData[1][1]
Global $gColumHeader =  "ID|原始文件名|拍摄日期|目标文件名"
Global $gNamingRule = "[N]"
#EndRegion

#Region Global Constants
Global Const $FILERENAME_SUCCESS = 0
Global Const $FILERENAME_ORG_NOT_FOUND = -1
Global Const $FILERENAME_ORG_DST_SAME = -2
Global Const $FILERENAME_DST_EXISTS = -3
Global Const $FILERENAME_ORG_NOT_ACCESSABLE = -4
#EndRegion

#Region ### START Koda GUI section ### Form=d:\1--devspace\autoit-codes\exif renamer\forms\main.kxf
Global $ERMain = GUICreate("EXIF Renamer", 528, 532, 192, 124)
Global $Label1 = GUICtrlCreateLabel("源文件位置：", 24, 24, 76, 20)
GUICtrlSetFont(-1, 11, 400, 0, "MS Sans Serif")
Global $idInputFolder = GUICtrlCreateInput("", 112, 24, 293, 21)
Global $idButtonBrowseForFolder = GUICtrlCreateButton("浏览...", 416, 22, 86, 25)
Global $idProgressBar = GUICtrlCreateProgress(22, 477, 372, 17, $WS_BORDER)
Global $Group1 = GUICtrlCreateGroup("重命名规则：文件名", 23, 54, 381, 92)
Global $idInputNamingReg = GUICtrlCreateInput($gNamingRule, 39, 78, 338, 21)
Global $idButtonFileName = GUICtrlCreateButton("[N]  原文件名", 39, 110, 99, 25)
Global $idButtonPhotoDate = GUICtrlCreateButton("[YMD]  拍摄日期", 157, 111, 99, 25)
Global $idButtonCount = GUICtrlCreateButton("[C]  计数器", 279, 111, 99, 25)
GUICtrlCreateGroup("", -99, -99, 1, 1)
Global $idButtonExecute = GUICtrlCreateButton("执行", 415, 474, 86, 25)
Global $idListViewPreview = GUICtrlCreateListView($gColumHeader, 22, 161, 479, 294)

Global $idStatusBar1 = _GUICtrlStatusBar_Create($ERMain)
Global $idStatusBar1_PartsWidth[2] = [150, -1]
_GUICtrlStatusBar_SetParts($idStatusBar1, $idStatusBar1_PartsWidth)
_GUICtrlStatusBar_SetText($idStatusBar1, "FileCount in DIR", 0)
_GUICtrlStatusBar_SetText($idStatusBar1, "Directory", 1)
Global $idButtonPreview = GUICtrlCreateButton("预 览", 415, 61, 86, 86)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###


;_DebugSetup("debug exif renamer", True)



While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $idButtonBrowseForFolder
			$gFileLocation = _FileSelectFolder("浏览目标目录", "请选择照片所在的目录位置", "", $gFileLocation)
			ConsoleWrite("选定目录位置:" & $gFileLocation & @CRLF)
			If $gFileLocation == 0 Then
				ConsoleWrite("Continue Case with no directory selected" & @CRLF)
				ContinueLoop
			EndIf
			GUICtrlSetData($idInputFolder, $gFileLocation)

			;初始化预览数据
			ER_genData( )
			ER_updateListView( )

			_GUICtrlStatusBar_SetText($idStatusBar1, "发现文件：" & $gFileCount , 0)
			_GUICtrlStatusBar_SetText($idStatusBar1, "文件位置：" & $gFileLocation , 1)

		Case $idButtonFileName
			$vDestFileNameRule = GUICtrlRead( $idInputNamingReg )
			$vDestFileNameRule &= "[N]"
			GUICtrlSetData( $idInputNamingReg, $vDestFileNameRule )
		Case $idButtonPhotoDate
			$vDestFileNameRule = GUICtrlRead( $idInputNamingReg )
			$vDestFileNameRule &= "[YMD]"
			GUICtrlSetData( $idInputNamingReg, $vDestFileNameRule )
		Case $idButtonCount
			$vDestFileNameRule = GUICtrlRead( $idInputNamingReg )
			$vDestFileNameRule &= "[C]"
			GUICtrlSetData( $idInputNamingReg, $vDestFileNameRule )
		Case $idButtonPreview
			$gNamingRule = GUICtrlRead( $idInputNamingReg )

			ER_updateDestFilename()
			ER_updateListView()
		Case $idButtonExecute

			If $gFileCount == 0 Then
				ConsoleWrite("Nothing to be processed" & @CRLF)
				ContinueLoop
			EndIf
			$gNamingRule = GUICtrlRead( $idInputNamingReg )
			ER_updateDestFilename()
			ER_Execute()

			;显示处理结果
			MsgBox($MB_SYSTEMMODAL, "", "处理结束")

			;重新初始化各内存变量
			ER_init()
			GUICtrlSetData( $idProgressBar, 0 )
			_GUICtrlStatusBar_SetText($idStatusBar1, "" , 0)
			_GUICtrlStatusBar_SetText($idStatusBar1, "" , 1)

	EndSwitch
WEnd


Func getShootDT( $filename )
	Local $myImgObject = _GDIPlus_ImageLoadFromFile( $filename )
	If @error Then
		MsgBox (0, "Error", "Can't open file.")
		Return -1
	Endif

	Local $v_ImgProperty = _GDIPlus_ImageGetPropertyItem( $myImgObject, $GDIP_PROPERTYTAGEXIFDTORIG )
	;_DebugReportVar( "$v_imageproperty" , DllStructGetData($v_ImgProperty, "id"))
	;ConsoleWrite( @CRLF & "PP - ErrorCODE is " & $GDIP_STATUS & " filename is " & $filename & @CRLF & " vPhotoDateStruct id is " & DllStructGetData($v_ImgProperty, "id") & @CRLF  & " type is " & DllStructGetData($v_ImgProperty, "type") &@CRLF & " length is " & DllStructGetData($v_ImgProperty, "length") & @CRLF & "value is " & DllStructGetData($v_ImgProperty, "value") & @CRLF )

	Local $myStructString= DllStructCreate("CHAR[" & DllStructGetData($v_ImgProperty, "length") & "];", DllStructGetData($v_ImgProperty, "value") )
	Local $vstring = DllStructGetData($myStructString, 1 )
	;ConsoleWrite( "photo date is " & $vstring  & "===" & @CRLF)
	_GDIPlus_ImageDispose ( $myImgObject )
	Return StringRegExpReplace( $vstring, "(\d{4}):(\d{2}):(\d{2})\s(\d{2}):(\d{2}):(\d{2})", "$1-$2-$3_$4$5$6" )

	;Return $rString
EndFunc ;==> end of readImgDatetime

func genFilename( $filename, $photoDT, $fileCount, $counter )

	Local $rule = $gNamingRule

	Local $_dstFN = StringSplit( $filename, "." ) ;文件名
	Local $_dstFNExt = $_dstFN[2]					;现有文件的扩展名

	$_dstFilename = $_dstFN[1]

	;_DebugReportVar( "_dstFN[0] is ", $_dstfn[0])
	;_DebugReportVar( "_dstFN[1] is ", $_dstfn[1])

	$_dstFilename = StringRegExpReplace( $rule, "\[[N|n]\]", $_dstFilename )
	$_dstFilename = StringRegExpReplace( $_dstFilename, "\[[Y|y][M|m][D|d]\]", $photoDT )
	$_dstFilename = StringRegExpReplace( $_dstFilename, "\[[C|c]\]", StringFormat( "%04s", $fileCount) )

	if $counter > 0 then
		$_dstFilename = $_dstFilename & "_" & $counter & "." & $_dstFNExt	;形成新的文件名
	Else
		$_dstFilename = $_dstFilename & "." & $_dstFNExt	;形成新的文件名
	EndIf

	return $_dstFilename
EndFunc

Func ER_Execute()
	;read data from gui control listView
	Local $retCode = 0

	if $gFileCount == 0 Then
		Return
	EndIf

	for $i=1 to $gFileCount Step 1

		$_filenameOrg = $gFileLocation & "\" & $gRenamingData[$i-1][1]
		$_filenameDst = $gFileLocation & "\" & $gRenamingData[$i-1][3]

		;_DebugReportVar( "Origenal file name is : ", $_filenameOrg)
		;_DebugReportVar( "Destination file name is : ", $_filenameDst)
		$retCode = _fileRename( $_filenameOrg, $_filenameDst )
		Switch $retCode
			Case $FILERENAME_SUCCESS
				ContinueLoop
			Case $FILERENAME_ORG_NOT_FOUND
				ConsoleWrite( "File " & $_filenameOrg & " not found. Moved? Go to next file(s)." )
			Case $FILERENAME_ORG_DST_SAME

			Case $FILERENAME_ORG_NOT_ACCESSABLE

			Case $FILERENAME_DST_EXISTS
				;目标文件存在，需要改名
				ConsoleWrite( "File " & $_filenameDst & " Exists. Please report the bug." )
		EndSwitch
		GUICtrlSetData( $idProgressBar, $i/$gFileCount*100 )
	Next
EndFunc

Func _fileRename( $org, $dst )
	Local $retCode

	if not FileExists( $org ) Then
		;ConsoleWrite( "File not found: " & $org & @CRLF )
		Return $FILERENAME_ORG_NOT_FOUND
	EndIf

	If StringCompare( $org, $dst ) == 0 Then
		;ConsoleWrite( "org and dst is same" & @CRLF )
		return $FILERENAME_ORG_DST_SAME
	EndIf

	if FileExists( $dst ) Then
		;ConsoleWrite( "File exists in destination: " & $dst & @CRLF )
		Return $FILERENAME_DST_EXISTS
	EndIf

	$retCode = FileMove( $org, $dst )
	If $retCode == 0 Then
		;ConsoleWrite( "Can not rename original file. File in use?" & @CRLF )
		Return $FILERENAME_ORG_NOT_ACCESSABLE
	EndIf
	Return $FILERENAME_SUCCESS

EndFunc

Func ER_init()
	;初始化全局变量
	$gFileLocation=""
	ReDim $gRenamingData[1][1]
	$gFileCount=0

	;初始化控件内容
	GUICtrlSetData( $idInputFolder, "" )
	_GUICtrlListView_DeleteAllItems( $idListViewPreview )

	Return
EndFunc

Func ER_genData( )
	Local $lFileList, $lFilePath, $lShootDT
	;罗列目录中的照片文件，要求扩展名为jpg
	$lFileList = _FileListToArrayRec($gFileLocation, "*.jpg;*.jpeg", $FLTAR_FILES)
	If @error Then
		MsgBox(BitOR($MB_SYSTEMMODAL,$MB_ICONERROR), "错误", "没有找到照片文件（JPG/JPEG）")
		Return -1
	EndIf

	ConsoleWrite("Found files in Directory:" & $lFileList[0] & @CRLF)

	$gFileCount = $lFileList[0]

	ReDim $gRenamingData[ $lFileList[0] ][4]

	_GDIPlus_Startup( )

	For $i = 1 To $lFileList[0] Step 1

		$lFilePath = $gFileLocation & "\" & $lFileList[$i]
		;ConsoleWrite( "Processing file " & $lFilePath &@CRLF )
		;读取照片拍摄时间
		$lShootDT = getShootDT( $lFilePath )

		;记录预览数据
		$gRenamingData[$i-1][0] = $i
		$gRenamingData[$i-1][1] = $lFileList[$i]
		$gRenamingData[$i-1][2] = $lShootDT

		;设置ProgressBar数据
		_GUICtrlStatusBar_SetText($idStatusBar1, "读取文件：" & $i & "/" & $gFileCount , 0)
		GUICtrlSetData( $idProgressBar, $i/$gFileCount*100 )
	Next

	GUICtrlSetData( $idProgressBar, 0 )

	ER_updateDestFilename()

	_GDIPlus_Shutdown()
EndFunc

Func ER_updateDestFilename()
	Local $lDestFileName, $lCounter=0

	if $gFileCount == 0 Then
		Return
	EndIf

	For $i = 1 To $gFileCount Step 1
		$gRenamingData[$i-1][3] = ""
	Next

	For $i = 1 To $gFileCount Step 1
		$lDestFileName = genFilename( $gRenamingData[$i-1][1], $gRenamingData[$i-1][2], $gRenamingData[$i-1][0], $lCounter )

		$lDestFileExist = _ArraySearch( $gRenamingData, $lDestFileName, 0, 0, 0, 1, 1, 3 )
		if @error <> 6 Then
			$lCounter += 1
			$i -=1
			ContinueLoop
		Else
			$lCounter = 0
		EndIf
		$gRenamingData[$i-1][3] = $lDestFileName
		;设置ProgressBar数据
		_GUICtrlStatusBar_SetText($idStatusBar1, "生成目标文件名：" & $i & "/" & $gFileCount , 0)
		GUICtrlSetData( $idProgressBar, $i/$gFileCount*100 )

	Next
EndFunc

Func ER_updateListView()
	If _GUICtrlListView_GetItemCount( $idListViewPreview ) > 0 Then
		_GUICtrlListView_DeleteAllItems( $idListViewPreview )
	EndIf
	_GUICtrlListView_AddArray( $idListViewPreview, $gRenamingData )

	_GUICtrlListView_SetColumnWidth( $idListViewPreview, 0, $LVSCW_AUTOSIZE  )
	_GUICtrlListView_SetColumnWidth( $idListViewPreview, 1, $LVSCW_AUTOSIZE  )
	_GUICtrlListView_SetColumnWidth( $idListViewPreview, 2, $LVSCW_AUTOSIZE  )
	_GUICtrlListView_SetColumnWidth( $idListViewPreview, 3, $LVSCW_AUTOSIZE  )
EndFunc