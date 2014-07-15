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

#include <debug.au3>

#include "Libs/GDIPlus_Image.au3"
#include "FileSelectFolder.au3"
#include "image_get_info.au3"

;#include "Libs\GDIPConstants.au3"



#Region ### START Koda GUI section ### Form=D:\1--DevSpace\Autoit-Codes\Exif Renamer\Forms\Main.kxf

$ERMain = GUICreate("EXIF Renamer", 504, 538, 192, 124)
$Label1 = GUICtrlCreateLabel("源文件位置：", 24, 24, 76, 20)
GUICtrlSetFont(-1, 11, 400, 0, "MS Sans Serif")
Global $idInputFolder = GUICtrlCreateInput("", 112, 24, 273, 21)
$idButtonBrowseForFolder = GUICtrlCreateButton("浏览...", 400, 24, 75, 25)


$Group1 = GUICtrlCreateGroup("重命名规则：文件名", 24, 56, 449, 92)
Global $idInputNamingReg = GUICtrlCreateInput("[N]", 40, 80, 409, 21)
$idButtonFileName = GUICtrlCreateButton("[N]  原文件名", 74, 112, 99, 25)
$idButtonPhotoDate = GUICtrlCreateButton("[YMD]  拍摄日期", 192, 113, 99, 25)
$idButtonCount = GUICtrlCreateButton("[C]  计数器", 314, 113, 99, 25)
GUICtrlCreateGroup("", -99, -99, 1, 1);close Group

$idProgressBar = GUICtrlCreateProgress(23, 478, 364, 17, $WS_BORDER)
$idButtonExecute = GUICtrlCreateButton("执行", 397, 474, 75, 25)

$idListViewPreview = GUICtrlCreateListView("原文件名|拍摄日期|目标文件名", 23, 161, 450, 294)
;Local $idItem1 = GUICtrlCreateListViewItem("MSC-1220.jpg|2014-04-05 15:30:20|2014-04-05_153020.jpg", $idListViewPreview)

$idStatusBar1 = _GUICtrlStatusBar_Create($ERMain)
Dim $idStatusBar1_PartsWidth[2] = [150, -1]
_GUICtrlStatusBar_SetParts($idStatusBar1, $idStatusBar1_PartsWidth)

GUISetState(@SW_SHOW)

#EndRegion ### END Koda GUI section ###

Local $aFileList
Local $file
Local $_FileSelectFolder = ""
;_DebugSetup("debug exif renamer", True)

Global $gFileLocation=""
Global $gRenamingData=0

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $idButtonBrowseForFolder
			$_FileSelectFolder = _FileSelectFolder("浏览目标目录", "请选择照片所在的目录位置", "", $_FileSelectFolder)
			ConsoleWrite("选定目录位置:" & $_FileSelectFolder & @CRLF)
			If $_FileSelectFolder == 0 Then
				ConsoleWrite("Continue Case with no directory selected" & @CRLF)
				ContinueLoop
			EndIf
			GUICtrlSetData($idInputFolder, $_FileSelectFolder)

			;罗列目录中的照片文件，要求扩展名为jpg
			$aFileList = _FileListToArrayRec($_FileSelectFolder, "*.jpg;*.jpeg", $FLTAR_FILES)
			If @error Then
				MsgBox(BitOR($MB_SYSTEMMODAL,$MB_ICONERROR), "错误", "没有找到照片文件（JPG/JPEG）")
				ContinueLoop
			EndIf

			ConsoleWrite("Found files in Directory:" & $aFileList[0] & @CRLF)
			If _GUICtrlListView_GetItemCount( $idListViewPreview ) > 0 Then
				_GUICtrlListView_DeleteAllItems( $idListViewPreview )
			EndIf

			_GUICtrlStatusBar_SetText($idStatusBar1, "发现文件：" & $aFileList[0] , 0)
			_GUICtrlStatusBar_SetText($idStatusBar1, "文件位置：" & $_FileSelectFolder , 1)
			$vDestFileNameRule = GUICtrlRead( $idInputNamingReg )
			updateListView( $aFileList, $vDestFileNameRule )

		Case $idButtonFileName
			$vDestFileNameRule = GUICtrlRead( $idInputNamingReg )
			$vDestFileNameRule &= "[N]"
			GUICtrlSetData( $idInputNamingReg, $vDestFileNameRule )
			updateListView($aFileList, $vDestFileNameRule )
		Case $idButtonPhotoDate
			$vDestFileNameRule = GUICtrlRead( $idInputNamingReg )
			$vDestFileNameRule &= "[YMD]"
			GUICtrlSetData( $idInputNamingReg, $vDestFileNameRule )
			updateListView($aFileList, $vDestFileNameRule )
		Case $idButtonCount
			$vDestFileNameRule = GUICtrlRead( $idInputNamingReg )
			$vDestFileNameRule &= "[C]"
			GUICtrlSetData( $idInputNamingReg, $vDestFileNameRule )
			updateListView($aFileList, $vDestFileNameRule )

		Case $idButtonExecute
			ConsoleWrite("$aFileList is " & VarGetType($aFileList) & " data type" & @CRLF)
			;判断$aFileList的有效性：未初始化
			If Not IsArray($aFileList) Then
				ContinueLoop
			EndIf
			;判断$aFileList的有效性：已处理过，相关数据已失效
			If $aFileList[0] == 0 Then
				ConsoleWrite("Nothing to be processed" & @CRLF)
				ContinueLoop
			EndIf

			_Execute()

			;显示处理结果
			MsgBox($MB_SYSTEMMODAL, "", "处理结束" & VarGetType($aFileList))

			;重新初始化各内存变量
			$aFileList = 0
			_GUICtrlListView_DeleteAllItems( $idListViewPreview )
			_GUICtrlStatusBar_SetText($idStatusBar1, "" , 0)
			_GUICtrlStatusBar_SetText($idStatusBar1, "" , 1)

		Case $idInputNamingReg
			If Not IsArray($aFileList) Then
				ContinueLoop
			EndIf
			;判断$aFileList的有效性：已处理过，相关数据已失效
			If $aFileList[0] == 0 Then
				ConsoleWrite("Nothing to be processed" & @CRLF)
				ContinueLoop
			EndIf

			_GUICtrlListView_DeleteAllItems( $idListViewPreview )
			_GUICtrlStatusBar_SetText($idStatusBar1, "" , 0)
			_GUICtrlStatusBar_SetText($idStatusBar1, "" , 1)

			$vDestFileNameRule = GUICtrlRead( $idInputNamingReg )
			updateListView( $aFileList, $vDestFileNameRule )
	EndSwitch
WEnd

#cs
Func readImgDatetime( $filename )
	Local $myImgObject = _GDIPlus_ImageLoadFromFile( $file )
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

func genFilename( $rule, $filename, $photoDT, $fileCount, $counter )

	Local $_dstFN = StringSplit( $filename, "." ) ;文件名
	Local $_dstFNExt = $_dstFN[2]					;现有文件的扩展名

	$_dstFilename = $_dstFN[1]

	;_DebugReportVar( "_dstFN[0] is ", $_dstfn[0])
	;_DebugReportVar( "_dstFN[1] is ", $_dstfn[1])

	$_dstFilename = StringRegExpReplace( $rule, "\[[N|n]\]", $_dstFilename )
	$_dstFilename = StringRegExpReplace( $_dstFilename, "\[[Y|y][M|m][D|d]\]", $photoDT )
	$_dstFilename = StringRegExpReplace( $_dstFilename, "\[[C|c]\]", $fileCount )
	if $counter > 0 then
		$_dstFilename = $_dstFilename & "_" & $counter & "." & $_dstFNExt	;形成新的文件名
	Else
		$_dstFilename = $_dstFilename & "." & $_dstFNExt	;形成新的文件名
	EndIf

	return $_dstFilename
EndFunc

Func updateListView( $fileList, $rule )

	Local $namingCounter=0, $tInfo

	if not IsArray($fileList ) Then
		Return
	EndIf

	_GDIPlus_Startup( )

	_GUICtrlListView_DeleteAllItems( $idListViewPreview )
	_GUICtrlListView_BeginUpdate( $idListViewPreview )
	For $i=1 To $fileList[0] Step 1
		$file = $_FileSelectFolder & "\" & $fileList[$i]
		$rString = readImgDatetime($file)
		$vDestFileName = genFilename( $rule, $fileList[$i], $rString, $i, $namingCounter )
		;测试目标文件名是否可用
		$tInfo = DllStructCreate($tagLVFINDINFO)
		DllStructSetData($tInfo, "Flags", $LVFI_STRING )
		DllStructSetData($tInfo, "Text", $vDestFileName)

		If _GUICtrlListView_FindInText( $idListViewPreview, $vDestFileName ) >= 0 Then
			;目标文件名已存在，需要解决目标文件名冲突问题
			$namingCounter +=1
			$i -= 1
			ContinueLoop
		Else
			$namingCounter = 0
		EndIf

		GUICtrlCreateListViewItem( $fileList[$i] & "|" & $rString & "|" & $vDestFileName, $idListViewPreview )

		_GUICtrlStatusBar_SetText($idStatusBar1, "读取文件：" & $i & "/" & $fileList[0] , 0)
	Next
	_GUICtrlListView_EndUpdate( $idListViewPreview )
	_GDIPlus_Shutdown()

EndFunc

Func _Execute()
	;read data from gui control listView
	Local $retCode = 0, $count

	$count = _GUICtrlListView_GetItemCount( $idListViewPreview )

	if $count == 0 Then
		Return
	EndIf

	for $i=0 to $count-1 Step 1
		$fileArray = _GUICtrlListView_GetItemTextArray( $idListViewPreview, $i )

		$_filenameOrg = $_FileSelectFolder & "\" & $fileArray[1]
		$_filenameDst = $_FileSelectFolder & "\" & $fileArray[3]

		;_DebugReportVar( "Origenal file name is : ", $_filenameOrg)
		;_DebugReportVar( "Destination file name is : ", $_filenameDst)
		$retCode = _fileRename( $_filenameOrg, $_filenameDst )
		Switch $retCode
			Case 0
				ContinueLoop
			Case -1
				ConsoleWrite( "File " & $_filenameOrg & " not found. Moved? Go to next file(s)." )
			Case -2
				;目标文件存在，需要改名
				ConsoleWrite( "File " & $_filenameDst & " Exists. Please report the bug." )
		EndSwitch
	Next
EndFunc

Func _fileRename( $org, $dst )

	if not FileExists( $org ) Then
		ConsoleWrite( "File not found: " & $org & @CRLF )
		Return -1
	EndIf

	if FileExists( $dst ) Then
		ConsoleWrite( "File exists in destination: " & $dst & @CRLF )
		Return -2
	EndIf

	FileMove( $org, $dst )
	Return 0

EndFunc
#ce

Func ER_init()
	;初始化全局变量
	$gFileLocation=""
	$gRenamingData=0

	;初始化控件内容
	GUICtrlSetData( $idInputFolder, "" )
	;_GUICtrlListView_DeleteAllItems( $idListViewPreview )
	for $i = 1 to _GUICtrlListView_GetColumnCount( $idListViewPreview ) Step 1
		_GUICtrlListView_DeleteColumn( $idListViewPreview, 0 )
	Next
	Return
EndFunc

Func ER_genData( )

EndFunc

