#Include <GUIConstantsEx.au3>
#Include <GUIImageList.au3>
#Include <GUITreeView.au3>
#Include <TreeViewConstants.au3>
#Include <WindowsConstants.au3>
#Include <WinAPIEx.au3>
#Include <Array.au3>

#AutoIt3Wrapper_Run_AU3Check=n

; Code from:
; http://www.autoitscript.com/forum/topic/161098-can-you-please-test-my-custom-fileselectfolder-on-your-system/?hl=select+folder

;~ Opt('MustDeclareVars', 1)

;$_FileSelectFolder = _FileSelectFolder("Browse for folder","Choose where you want to place the Script","Script Name")
;ConsoleWrite("Selected path:"&@CRLF&$_FileSelectFolder&@CRLF)


; #FUNCTION# ====================================================================================================================
; Name ..........: _FileSelectFolder
; Description ...: Custom FileSelectFolder
; Syntax ........: _FileSelectFolder([$Title = ""[, $Text = ""[, $InstallDir = ""]]])
; Parameters ....: $Title               - [optional] The title of the GUI. Default is "".
;                  $Text                - [optional] The text to show in the GUI. Default is "".
;                  $InstallDir          - [optional] If $InstallDir <> "" then when you select a folder, the path will automatically update with the new folder... Default is "".
; Return values .: 0 if the user exit from the GUI
;                   Path string if the user cliked OK
; Author ........: gil900. Based on:
;                   http://www.autoitscript.com/forum/topic/124430-display-on-the-fly-a-directory-tree-in-a-treeview/
; Modified ......: NONE
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; Thanks to .....: Yashied
; ===============================================================================================================================
Func _FileSelectFolder($Title = "",$Text = "",$InstallDir = "",$DefaultDir = "")
    Local $OK_Button , $Cancel_Button , $aDrives
    Local $hForm , $hImageList , $hItem , $hNext , $sPath , $Output = 0
    Global $FSF_hTreeView , $FSF_hSelect = 0, $FSF_Path_Input, $FSF_Dummy1, $FSF_Dummy2 , $FSF_GInstallDir = $InstallDir , $FSF_GDefaultDir = $DefaultDir;, $sRoot = "\" , $Input, $FSF_GDefaultDir

    $hForm = GUICreate($Title, 328, 314)
	If StringLen($DefaultDir) > 0 Then
		$FSF_Path_Input = GUICtrlCreateInput($DefaultDir, 12, 36, 304, 21)
	Else
		$FSF_Path_Input = GUICtrlCreateInput($InstallDir, 12, 36, 304, 21)
	EndIf

    $OK_Button = GUICtrlCreateButton("OK", 160, 280, 73, 25)
    $Cancel_Button = GUICtrlCreateButton("Cancel", 239, 280, 73, 25)
    GUICtrlCreateLabel($Text, 16, 11, 300, 17)
    GUICtrlSetFont(-1, 9, 800, 0, "Tahoma")

    GUICtrlCreateTreeView(12, 74, 304, 186, -1, $WS_EX_CLIENTEDGE)
    $FSF_hTreeView = GUICtrlGetHandle(-1)
    $FSF_Dummy1 = GUICtrlCreateDummy()
    $FSF_Dummy2 = GUICtrlCreateDummy()


    If _WinAPI_GetVersion() >= '6.0' Then
        _WinAPI_SetWindowTheme($FSF_hTreeView, 'Explorer')
    EndIf

    $hImageList = _GUIImageList_Create(16, 16, 5, 1)
    _GUIImageList_AddIcon($hImageList, @SystemDir & '\shell32.dll', 3)
    _GUIImageList_AddIcon($hImageList, @SystemDir & '\shell32.dll', 4)
    _GUICtrlTreeView_SetNormalImageList($FSF_hTreeView, $hImageList)

    $aDrives = DriveGetDrive("FIXED")
    If IsArray($aDrives) Then
        For $a = 1 To $aDrives[0]
            _TVUpdate($FSF_hTreeView,_GUICtrlTreeView_AddChild($FSF_hTreeView, 0,StringUpper($aDrives[$a]), 0, 0) )
        Next
    EndIf
    _TVUpdate($FSF_hTreeView,_GUICtrlTreeView_AddChild($FSF_hTreeView, 0,@DesktopDir, 0, 0) )

    GUIRegisterMsg($WM_NOTIFY, 'WM_NOTIFY')
    GUISetState()

    ;If $GInstallDir <> "" Then $GInstallDir = "\"&$GInstallDir

    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE , $Cancel_Button
                $Output = 0
                ExitLoop
            Case $FSF_Dummy1 ; Update
                ;_ArrayDisplay(GUIGetMsg(1))

                GUISetCursor(1, 1)
                $hItem = _GUICtrlTreeView_GetFirstChild($FSF_hTreeView, GUICtrlRead($FSF_Dummy1))
                If $hItem Then
                    While $hItem
                        $hNext = _GUICtrlTreeView_GetNextSibling($FSF_hTreeView, $hItem)
                        If Not _TVUpdate($FSF_hTreeView, $hItem) Then
                            _GUICtrlTreeView_Delete($FSF_hTreeView, $hItem)
                        EndIf
                        $hItem = $hNext
                    WEnd
                    _WinAPI_RedrawWindow($FSF_hTreeView)
                EndIf
                GUISetCursor(2, 0)
            Case $OK_Button
                $Output = StringStripWS(GUICtrlRead($FSF_Path_Input),3)
				;MsgBox(0, "Info", "Selected path is " & $Output )
                If Not PathIsValid($Output) Then
                    MsgBox(48,"Error","Please select a valid path",0,$hForm)
                Else
                    ExitLoop
                EndIf
        EndSwitch
    WEnd

    GUIDelete($hForm)
    GUIRegisterMsg($WM_NOTIFY, '')
    Return $Output
EndFunc



Func _TVGetPath($hTV, $hItem, $sRoot)
    Local $Path = StringRegExpReplace(_GUICtrlTreeView_GetTree($hTV, $hItem), '([|]+)|(\\[|])', '\\')
    If Not $Path Then
        Return ''
    EndIf
    If Not StringInStr($Path, ':') Then
        Return StringRegExpReplace($sRoot, '(\\[^\\]*(\\|)+)\Z', '\\') & $Path
    EndIf
    Return $Path
EndFunc   ;==>_TVGetPath

Func _TVSetPath($hTV, $hItem, $sRoot)
    ;GUICtrlSetData($Input, _WinAPI_PathCompactPath($hInput, _TVGetPath($hTV, $hItem, $sRoot), 554))
    Local $NewPath = _TVGetPath($hTV, $hItem, $sRoot)
    If $FSF_GInstallDir <> "" And StringRight($NewPath,1) <> "\" Then $NewPath = $NewPath&"\"
    GUICtrlSetData($FSF_Path_Input,$NewPath&$FSF_GInstallDir)
    $FSF_hSelect = $hItem
EndFunc   ;==>_TVSetPath

Func _TVUpdate($hTV, $hItem)

    Local $hImageList = _SendMessage($hTV, $TVM_GETIMAGELIST)

    Local $Path = StringRegExpReplace(_TVGetPath($hTV, $hItem, "\"), '\\+\Z', '')
    ;ConsoleWrite($Path&@CRLF)

    Local $var = StringSplit($Path,"\",1)
    Local $foldername = $var[$var[0]]
    Local $Excluded[8] = [7, "$Recycle.Bin", "Documents and Settings", 'MSOCache', "PerfLogs", "ProgramData", "Recovery" , "System Volume Information"]
    If _ArraySearch($Excluded,$foldername,1) > 0 Then Return 0
    ;If DirGetSize($Path) = 0 Then Return 1
    Local $hSearch, $hIcon, $Index, $File

    $hSearch = FileFindFirstFile($Path & '\*')
    If $hSearch = -1 Then
        If Not @error Then
            If FileExists($Path) Then
;               If _WinAPI_PathIsDirectory($Path) Then
;                   ; Access denied
;               EndIf
            Else
                Return 0
            EndIf
        EndIf
    Else
        While 1
            $File = FileFindNextFile($hSearch)
            If @error Then
                ExitLoop
            EndIf
            ;If DirGetSize($Path) = 0 Then Return 0

            If @extended Then
                _GUICtrlTreeView_AddChild($hTV, $hItem, $File, 0, 0)

            EndIf
        WEnd
        FileClose($hSearch)
        ;ConsoleWrite($File&@CRLF)
    EndIf
    Return 1
EndFunc   ;==>_TVUpdate


Func WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
    ;ConsoleWrite($hWnd&" , "&$iMsg&" , "&$wParam&" , "&$lParam&@CRLF)
    Local $tNMTREEVIEW = DllStructCreate($tagNMTREEVIEW, $lParam)
    Local $hItem = DllStructGetData($tNMTREEVIEW, 'NewhItem')
    Local $iState = DllStructGetData($tNMTREEVIEW, 'NewState')
    Local $hTV = DllStructGetData($tNMTREEVIEW, 'hWndFrom')
    Local $ID = DllStructGetData($tNMTREEVIEW, 'Code')
    Local $tTVHTI, $tPoint

    Switch $hTV
        Case $FSF_hTreeView
            Switch $ID
                Case $TVN_ITEMEXPANDEDW
                    If Not FileExists(_TVGetPath($hTV, $hItem, "\")) Then
                        _GUICtrlTreeView_Delete($hTV, $hItem)
                        If BitAND($iState, $TVIS_SELECTED) Then
                            _TVSetPath($hTV, _GUICtrlTreeView_GetSelection($hTV), "\")
                        EndIf
                    Else
                        If Not BitAND($iState, $TVIS_EXPANDED) Then
                            _GUICtrlTreeView_SetSelectedImageIndex($hTV, $hItem, 0)
                            _GUICtrlTreeView_SetImageIndex($hTV, $hItem, 0)
                        Else
                            _GUICtrlTreeView_SetSelectedImageIndex($hTV, $hItem, 1)
                            _GUICtrlTreeView_SetImageIndex($hTV, $hItem, 1)
                            If Not _GUICtrlTreeView_GetItemParam($hTV, $hItem) Then
                                _GUICtrlTreeView_SetItemParam($hTV, $hItem, 0x7FFFFFFF)
                                GUICtrlSendToDummy($FSF_Dummy1, $hItem)
                            EndIf
                        EndIf
                    EndIf
                Case $TVN_SELCHANGEDW
                    If BitAND($iState, $TVIS_SELECTED) Then
                        If Not FileExists(_TVGetPath($hTV, $hItem, "\")) Then
                            _GUICtrlTreeView_Delete($hTV, $hItem)
                            $hItem = _GUICtrlTreeView_GetSelection($hTV)
                        EndIf
                        If $hItem <> $FSF_hSelect Then
                            _TVSetPath($hTV, $hItem, "\")
                        EndIf
                    EndIf
                Case $NM_RCLICK
                        $tPoint = _WinAPI_GetMousePos(1, $hTV)
                        $tTVHTI = _GUICtrlTreeView_HitTestEx($hTV, DllStructGetData($tPoint, 1), DllStructGetData($tPoint, 2))
                        $hItem = DllStructGetData($tTVHTI, 'Item')
                        If BitAND(DllStructGetData($tTVHTI, 'Flags'), $TVHT_ONITEM) Then
                            _GUICtrlTreeView_SelectItem($FSF_hTreeView, $hItem)
                            If Not FileExists(_TVGetPath($hTV, $hItem, "\")) Then
                                _GUICtrlTreeView_Delete($hTV, $hItem)
                                $hItem = _GUICtrlTreeView_GetSelection($hTV)
                            Else
                                GUICtrlSendToDummy($FSF_Dummy2, $hItem)
                            EndIf
                            If $hItem <> $FSF_hSelect Then
                                _TVSetPath($hTV, $hItem, "\")
                            EndIf
                        EndIf
                EndSwitch
    EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY



Func PathIsValid($Path)
    Local $var = StringSplit($Path,"\",1)
    If StringLen($var[1]) = 2 And StringRight($var[1],1) = ":" And StringIsASCII($var[1]) And StringIsAlpha(StringLeft($var[1],1)) Then
        ;Return 1
        If $var[0] > 1 Then
            Local $Excluded[9] = [8, "?", "*", '"', "<", ">", "|" , "/" , ":"]
            For $a = 2 To $var[0]
                If $var[$a] = "" And $a > 2 Or StringIsSpace($var[$a]) Then Return 0
                If StringStripWS($var[$a],3) <> $var[$a] Then Return 0
                For $a2 = 1 To $Excluded[0]
                    If StringInStr($var[$a],$Excluded[$a2]) > 0 Then Return 0
                Next
            Next
            Return 1
        Else
            Return 1
        EndIf
    Else
        Return 0
    EndIf
EndFunc