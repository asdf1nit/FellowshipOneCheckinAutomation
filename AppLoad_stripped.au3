Global Const $MB_SYSTEMMODAL = 4096
Global Const $PS_SOLID = 0
Global Const $UBOUND_DIMENSIONS = 0
Global Const $UBOUND_COLUMNS = 2
Global Const $STR_STRIPLEADING = 1
Global Const $STR_STRIPTRAILING = 2
Global Const $STR_ENTIRESPLIT = 1
Global Const $STR_NOCOUNT = 2
Global Const $STR_REGEXPMATCH = 0
Global Const $tagPOINT = "struct;long X;long Y;endstruct"
Global Const $tagRECT = "struct;long Left;long Top;long Right;long Bottom;endstruct"
Global Const $tagSYSTEMTIME = "struct;word Year;word Month;word Dow;word Day;word Hour;word Minute;word Second;word MSeconds;endstruct"
Global Const $tagGDIPENCODERPARAM = "struct;byte GUID[16];ulong NumberOfValues;ulong Type;ptr Values;endstruct"
Global Const $tagGDIPENCODERPARAMS = "uint Count;" & $tagGDIPENCODERPARAM
Global Const $tagGDIPSTARTUPINPUT = "uint Version;ptr Callback;bool NoThread;bool NoCodecs"
Global Const $tagGDIPIMAGECODECINFO = "byte CLSID[16];byte FormatID[16];ptr CodecName;ptr DllName;ptr FormatDesc;ptr FileExt;" & "ptr MimeType;dword Flags;dword Version;dword SigCount;dword SigSize;ptr SigPattern;ptr SigMask"
Global Const $tagREBARBANDINFO = "uint cbSize;uint fMask;uint fStyle;dword clrFore;dword clrBack;ptr lpText;uint cch;" & "int iImage;hwnd hwndChild;uint cxMinChild;uint cyMinChild;uint cx;handle hbmBack;uint wID;uint cyChild;uint cyMaxChild;" & "uint cyIntegral;uint cxIdeal;lparam lParam;uint cxHeader" &((@OSVersion = "WIN_XP") ? "" : ";" & $tagRECT & ";uint uChevronState")
Global Const $tagGUID = "struct;ulong Data1;ushort Data2;ushort Data3;byte Data4[8];endstruct"
Global Const $HGDI_ERROR = Ptr(-1)
Global Const $INVALID_HANDLE_VALUE = Ptr(-1)
Global Const $KF_EXTENDED = 0x0100
Global Const $KF_ALTDOWN = 0x2000
Global Const $KF_UP = 0x8000
Global Const $LLKHF_EXTENDED = BitShift($KF_EXTENDED, 8)
Global Const $LLKHF_ALTDOWN = BitShift($KF_ALTDOWN, 8)
Global Const $LLKHF_UP = BitShift($KF_UP, 8)
Global Const $tagCURSORINFO = "dword Size;dword Flags;handle hCursor;" & $tagPOINT
Global Const $tagICONINFO = "bool Icon;dword XHotSpot;dword YHotSpot;handle hMask;handle hColor"
Func _WinAPI_BitBlt($hDestDC, $iXDest, $iYDest, $iWidth, $iHeight, $hSrcDC, $iXSrc, $iYSrc, $iROP)
Local $aResult = DllCall("gdi32.dll", "bool", "BitBlt", "handle", $hDestDC, "int", $iXDest, "int", $iYDest, "int", $iWidth, "int", $iHeight, "handle", $hSrcDC, "int", $iXSrc, "int", $iYSrc, "dword", $iROP)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_CopyIcon($hIcon)
Local $aResult = DllCall("user32.dll", "handle", "CopyIcon", "handle", $hIcon)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_CreateCompatibleBitmap($hDC, $iWidth, $iHeight)
Local $aResult = DllCall("gdi32.dll", "handle", "CreateCompatibleBitmap", "handle", $hDC, "int", $iWidth, "int", $iHeight)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_CreateCompatibleDC($hDC)
Local $aResult = DllCall("gdi32.dll", "handle", "CreateCompatibleDC", "handle", $hDC)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_CreatePen($iPenStyle, $iWidth, $iColor)
Local $aResult = DllCall("gdi32.dll", "handle", "CreatePen", "int", $iPenStyle, "int", $iWidth, "INT", $iColor)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_DeleteDC($hDC)
Local $aResult = DllCall("gdi32.dll", "bool", "DeleteDC", "handle", $hDC)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_DeleteObject($hObject)
Local $aResult = DllCall("gdi32.dll", "bool", "DeleteObject", "handle", $hObject)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_DestroyIcon($hIcon)
Local $aResult = DllCall("user32.dll", "bool", "DestroyIcon", "handle", $hIcon)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_DrawIcon($hDC, $iX, $iY, $hIcon)
Local $aResult = DllCall("user32.dll", "bool", "DrawIcon", "handle", $hDC, "int", $iX, "int", $iY, "handle", $hIcon)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_DrawLine($hDC, $iX1, $iY1, $iX2, $iY2)
_WinAPI_MoveTo($hDC, $iX1, $iY1)
If @error Then Return SetError(@error, @extended, False)
_WinAPI_LineTo($hDC, $iX2, $iY2)
If @error Then Return SetError(@error + 10, @extended, False)
Return True
EndFunc
Func _WinAPI_GetCursorInfo()
Local $tCursor = DllStructCreate($tagCURSORINFO)
Local $iCursor = DllStructGetSize($tCursor)
DllStructSetData($tCursor, "Size", $iCursor)
Local $aRet = DllCall("user32.dll", "bool", "GetCursorInfo", "struct*", $tCursor)
If @error Or Not $aRet[0] Then Return SetError(@error + 10, @extended, 0)
Local $aCursor[5]
$aCursor[0] = True
$aCursor[1] = DllStructGetData($tCursor, "Flags") <> 0
$aCursor[2] = DllStructGetData($tCursor, "hCursor")
$aCursor[3] = DllStructGetData($tCursor, "X")
$aCursor[4] = DllStructGetData($tCursor, "Y")
Return $aCursor
EndFunc
Func _WinAPI_GetDC($hWnd)
Local $aResult = DllCall("user32.dll", "handle", "GetDC", "hwnd", $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetDesktopWindow()
Local $aResult = DllCall("user32.dll", "hwnd", "GetDesktopWindow")
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetIconInfo($hIcon)
Local $tInfo = DllStructCreate($tagICONINFO)
Local $aRet = DllCall("user32.dll", "bool", "GetIconInfo", "handle", $hIcon, "struct*", $tInfo)
If @error Or Not $aRet[0] Then Return SetError(@error + 10, @extended, 0)
Local $aIcon[6]
$aIcon[0] = True
$aIcon[1] = DllStructGetData($tInfo, "Icon") <> 0
$aIcon[2] = DllStructGetData($tInfo, "XHotSpot")
$aIcon[3] = DllStructGetData($tInfo, "YHotSpot")
$aIcon[4] = DllStructGetData($tInfo, "hMask")
$aIcon[5] = DllStructGetData($tInfo, "hColor")
Return $aIcon
EndFunc
Func _WinAPI_GetSystemMetrics($iIndex)
Local $aResult = DllCall("user32.dll", "int", "GetSystemMetrics", "int", $iIndex)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetWindowDC($hWnd)
Local $aResult = DllCall("user32.dll", "handle", "GetWindowDC", "hwnd", $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GUIDFromString($sGUID)
Local $tGUID = DllStructCreate($tagGUID)
_WinAPI_GUIDFromStringEx($sGUID, $tGUID)
If @error Then Return SetError(@error + 10, @extended, 0)
Return $tGUID
EndFunc
Func _WinAPI_GUIDFromStringEx($sGUID, $tGUID)
Local $aResult = DllCall("ole32.dll", "long", "CLSIDFromString", "wstr", $sGUID, "struct*", $tGUID)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_LineTo($hDC, $iX, $iY)
Local $aResult = DllCall("gdi32.dll", "bool", "LineTo", "handle", $hDC, "int", $iX, "int", $iY)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_MoveTo($hDC, $iX, $iY)
Local $aResult = DllCall("gdi32.dll", "bool", "MoveToEx", "handle", $hDC, "int", $iX, "int", $iY, "ptr", 0)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_ReleaseDC($hWnd, $hDC)
Local $aResult = DllCall("user32.dll", "int", "ReleaseDC", "hwnd", $hWnd, "handle", $hDC)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_SelectObject($hDC, $hGDIObj)
Local $aResult = DllCall("gdi32.dll", "handle", "SelectObject", "handle", $hDC, "handle", $hGDIObj)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_StringFromGUID($tGUID)
Local $aResult = DllCall("ole32.dll", "int", "StringFromGUID2", "struct*", $tGUID, "wstr", "", "int", 40)
If @error Or Not $aResult[0] Then Return SetError(@error, @extended, "")
Return SetExtended($aResult[0], $aResult[2])
EndFunc
Func _WinAPI_WideCharToMultiByte($vUnicode, $iCodePage = 0, $bRetString = True)
Local $sUnicodeType = "wstr"
If Not IsString($vUnicode) Then $sUnicodeType = "struct*"
Local $aResult = DllCall("kernel32.dll", "int", "WideCharToMultiByte", "uint", $iCodePage, "dword", 0, $sUnicodeType, $vUnicode, "int", -1, "ptr", 0, "int", 0, "ptr", 0, "ptr", 0)
If @error Or Not $aResult[0] Then Return SetError(@error + 20, @extended, "")
Local $tMultiByte = DllStructCreate("char[" & $aResult[0] & "]")
$aResult = DllCall("kernel32.dll", "int", "WideCharToMultiByte", "uint", $iCodePage, "dword", 0, $sUnicodeType, $vUnicode, "int", -1, "struct*", $tMultiByte, "int", $aResult[0], "ptr", 0, "ptr", 0)
If @error Or Not $aResult[0] Then Return SetError(@error + 10, @extended, "")
If $bRetString Then Return DllStructGetData($tMultiByte, 1)
Return $tMultiByte
EndFunc
Func _ArraySearch(Const ByRef $aArray, $vValue, $iStart = 0, $iEnd = 0, $iCase = 0, $iCompare = 0, $iForward = 1, $iSubItem = -1, $bRow = False)
If $iStart = Default Then $iStart = 0
If $iEnd = Default Then $iEnd = 0
If $iCase = Default Then $iCase = 0
If $iCompare = Default Then $iCompare = 0
If $iForward = Default Then $iForward = 1
If $iSubItem = Default Then $iSubItem = -1
If $bRow = Default Then $bRow = False
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray) - 1
If $iDim_1 = -1 Then Return SetError(3, 0, -1)
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS) - 1
Local $bCompType = False
If $iCompare = 2 Then
$iCompare = 0
$bCompType = True
EndIf
If $bRow Then
If UBound($aArray, $UBOUND_DIMENSIONS) = 1 Then Return SetError(5, 0, -1)
If $iEnd < 1 Or $iEnd > $iDim_2 Then $iEnd = $iDim_2
If $iStart < 0 Then $iStart = 0
If $iStart > $iEnd Then Return SetError(4, 0, -1)
Else
If $iEnd < 1 Or $iEnd > $iDim_1 Then $iEnd = $iDim_1
If $iStart < 0 Then $iStart = 0
If $iStart > $iEnd Then Return SetError(4, 0, -1)
EndIf
Local $iStep = 1
If Not $iForward Then
Local $iTmp = $iStart
$iStart = $iEnd
$iEnd = $iTmp
$iStep = -1
EndIf
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
If Not $iCompare Then
If Not $iCase Then
For $i = $iStart To $iEnd Step $iStep
If $bCompType And VarGetType($aArray[$i]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$i] = $vValue Then Return $i
Next
Else
For $i = $iStart To $iEnd Step $iStep
If $bCompType And VarGetType($aArray[$i]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$i] == $vValue Then Return $i
Next
EndIf
Else
For $i = $iStart To $iEnd Step $iStep
If $iCompare = 3 Then
If StringRegExp($aArray[$i], $vValue) Then Return $i
Else
If StringInStr($aArray[$i], $vValue, $iCase) > 0 Then Return $i
EndIf
Next
EndIf
Case 2
Local $iDim_Sub
If $bRow Then
$iDim_Sub = $iDim_1
If $iSubItem > $iDim_Sub Then $iSubItem = $iDim_Sub
If $iSubItem < 0 Then
$iSubItem = 0
Else
$iDim_Sub = $iSubItem
EndIf
Else
$iDim_Sub = $iDim_2
If $iSubItem > $iDim_Sub Then $iSubItem = $iDim_Sub
If $iSubItem < 0 Then
$iSubItem = 0
Else
$iDim_Sub = $iSubItem
EndIf
EndIf
For $j = $iSubItem To $iDim_Sub
If Not $iCompare Then
If Not $iCase Then
For $i = $iStart To $iEnd Step $iStep
If $bRow Then
If $bCompType And VarGetType($aArray[$j][$j]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$j][$i] = $vValue Then Return $i
Else
If $bCompType And VarGetType($aArray[$i][$j]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$i][$j] = $vValue Then Return $i
EndIf
Next
Else
For $i = $iStart To $iEnd Step $iStep
If $bRow Then
If $bCompType And VarGetType($aArray[$j][$i]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$j][$i] == $vValue Then Return $i
Else
If $bCompType And VarGetType($aArray[$i][$j]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$i][$j] == $vValue Then Return $i
EndIf
Next
EndIf
Else
For $i = $iStart To $iEnd Step $iStep
If $iCompare = 3 Then
If $bRow Then
If StringRegExp($aArray[$j][$i], $vValue) Then Return $i
Else
If StringRegExp($aArray[$i][$j], $vValue) Then Return $i
EndIf
Else
If $bRow Then
If StringInStr($aArray[$j][$i], $vValue, $iCase) > 0 Then Return $i
Else
If StringInStr($aArray[$i][$j], $vValue, $iCase) > 0 Then Return $i
EndIf
EndIf
Next
EndIf
Next
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return SetError(6, 0, -1)
EndFunc
Global Const $GDIP_EPGCOLORDEPTH = '{66087055-AD66-4C7C-9A18-38A2310B8337}'
Global Const $GDIP_EPGCOMPRESSION = '{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}'
Global Const $GDIP_EPGQUALITY = '{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}'
Global Const $GDIP_EPTLONG = 4
Global Const $GDIP_EVTCOMPRESSIONLZW = 2
Global Const $GDIP_PXF24RGB = 0x00021808
Global Const $tagOSVERSIONINFO = 'struct;dword OSVersionInfoSize;dword MajorVersion;dword MinorVersion;dword BuildNumber;dword PlatformId;wchar CSDVersion[128];endstruct'
Global Const $__WINVER = __WINVER()
Func __WINVER()
Local $tOSVI = DllStructCreate($tagOSVERSIONINFO)
DllStructSetData($tOSVI, 1, DllStructGetSize($tOSVI))
Local $aRet = DllCall('kernel32.dll', 'bool', 'GetVersionExW', 'struct*', $tOSVI)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, 0)
Return BitOR(BitShift(DllStructGetData($tOSVI, 2), -8), DllStructGetData($tOSVI, 3))
EndFunc
Global $__g_hGDIPDll = 0
Global $__g_iGDIPRef = 0
Global $__g_iGDIPToken = 0
Global $__g_bGDIP_V1_0 = True
Func _GDIPlus_BitmapCloneArea($hBitmap, $nLeft, $nTop, $nWidth, $nHeight, $iFormat = 0x00021808)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipCloneBitmapArea", "float", $nLeft, "float", $nTop, "float", $nWidth, "float", $nHeight, "int", $iFormat, "handle", $hBitmap, "handle*", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Return $aResult[7]
EndFunc
Func _GDIPlus_BitmapCreateFromHBITMAP($hBitmap, $hPal = 0)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipCreateBitmapFromHBITMAP", "handle", $hBitmap, "handle", $hPal, "handle*", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Return $aResult[3]
EndFunc
Func _GDIPlus_Encoders()
Local $iCount = _GDIPlus_EncodersGetCount()
Local $iSize = _GDIPlus_EncodersGetSize()
Local $tBuffer = DllStructCreate("byte[" & $iSize & "]")
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipGetImageEncoders", "uint", $iCount, "uint", $iSize, "struct*", $tBuffer)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return SetError(10, $aResult[0], 0)
Local $pBuffer = DllStructGetPtr($tBuffer)
Local $tCodec, $aInfo[$iCount + 1][14]
$aInfo[0][0] = $iCount
For $iI = 1 To $iCount
$tCodec = DllStructCreate($tagGDIPIMAGECODECINFO, $pBuffer)
$aInfo[$iI][1] = _WinAPI_StringFromGUID(DllStructGetPtr($tCodec, "CLSID"))
$aInfo[$iI][2] = _WinAPI_StringFromGUID(DllStructGetPtr($tCodec, "FormatID"))
$aInfo[$iI][3] = _WinAPI_WideCharToMultiByte(DllStructGetData($tCodec, "CodecName"))
$aInfo[$iI][4] = _WinAPI_WideCharToMultiByte(DllStructGetData($tCodec, "DllName"))
$aInfo[$iI][5] = _WinAPI_WideCharToMultiByte(DllStructGetData($tCodec, "FormatDesc"))
$aInfo[$iI][6] = _WinAPI_WideCharToMultiByte(DllStructGetData($tCodec, "FileExt"))
$aInfo[$iI][7] = _WinAPI_WideCharToMultiByte(DllStructGetData($tCodec, "MimeType"))
$aInfo[$iI][8] = DllStructGetData($tCodec, "Flags")
$aInfo[$iI][9] = DllStructGetData($tCodec, "Version")
$aInfo[$iI][10] = DllStructGetData($tCodec, "SigCount")
$aInfo[$iI][11] = DllStructGetData($tCodec, "SigSize")
$aInfo[$iI][12] = DllStructGetData($tCodec, "SigPattern")
$aInfo[$iI][13] = DllStructGetData($tCodec, "SigMask")
$pBuffer += DllStructGetSize($tCodec)
Next
Return $aInfo
EndFunc
Func _GDIPlus_EncodersGetCLSID($sFileExtension)
Local $aEncoders = _GDIPlus_Encoders()
If @error Then Return SetError(@error, 0, "")
For $iI = 1 To $aEncoders[0][0]
If StringInStr($aEncoders[$iI][6], "*." & $sFileExtension) > 0 Then Return $aEncoders[$iI][1]
Next
Return SetError(-1, -1, "")
EndFunc
Func _GDIPlus_EncodersGetCount()
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipGetImageEncodersSize", "uint*", 0, "uint*", 0)
If @error Then Return SetError(@error, @extended, -1)
If $aResult[0] Then Return SetError(10, $aResult[0], -1)
Return $aResult[1]
EndFunc
Func _GDIPlus_EncodersGetSize()
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipGetImageEncodersSize", "uint*", 0, "uint*", 0)
If @error Then Return SetError(@error, @extended, -1)
If $aResult[0] Then Return SetError(10, $aResult[0], -1)
Return $aResult[2]
EndFunc
Func _GDIPlus_ImageDispose($hImage)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipDisposeImage", "handle", $hImage)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_ImageGetHeight($hImage)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipGetImageHeight", "handle", $hImage, "uint*", 0)
If @error Then Return SetError(@error, @extended, -1)
If $aResult[0] Then Return SetError(10, $aResult[0], -1)
Return $aResult[2]
EndFunc
Func _GDIPlus_ImageGetWidth($hImage)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipGetImageWidth", "handle", $hImage, "uint*", -1)
If @error Then Return SetError(@error, @extended, -1)
If $aResult[0] Then Return SetError(10, $aResult[0], -1)
Return $aResult[2]
EndFunc
Func _GDIPlus_ImageSaveToFileEx($hImage, $sFileName, $sEncoder, $tParams = 0)
Local $tGUID = _WinAPI_GUIDFromString($sEncoder)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipSaveImageToFile", "handle", $hImage, "wstr", $sFileName, "struct*", $tGUID, "struct*", $tParams)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
Return True
EndFunc
Func _GDIPlus_ParamAdd(ByRef $tParams, $sGUID, $iNbOfValues, $iType, $pValues)
Local $iCount = DllStructGetData($tParams, "Count")
Local $pGUID = DllStructGetPtr($tParams, "GUID") +($iCount * _GDIPlus_ParamSize())
Local $tParam = DllStructCreate($tagGDIPENCODERPARAM, $pGUID)
_WinAPI_GUIDFromStringEx($sGUID, $pGUID)
DllStructSetData($tParam, "Type", $iType)
DllStructSetData($tParam, "NumberOfValues", $iNbOfValues)
DllStructSetData($tParam, "Values", $pValues)
DllStructSetData($tParams, "Count", $iCount + 1)
EndFunc
Func _GDIPlus_ParamInit($iCount)
Local $sStruct = $tagGDIPENCODERPARAMS
For $i = 2 To $iCount
$sStruct &= ";struct;byte[16];ulong;ulong;ptr;endstruct"
Next
Return DllStructCreate($sStruct)
EndFunc
Func _GDIPlus_ParamSize()
Local $tParam = DllStructCreate($tagGDIPENCODERPARAM)
Return DllStructGetSize($tParam)
EndFunc
Func _GDIPlus_Shutdown()
If $__g_hGDIPDll = 0 Then Return SetError(-1, -1, False)
$__g_iGDIPRef -= 1
If $__g_iGDIPRef = 0 Then
DllCall($__g_hGDIPDll, "none", "GdiplusShutdown", "ulong_ptr", $__g_iGDIPToken)
DllClose($__g_hGDIPDll)
$__g_hGDIPDll = 0
EndIf
Return True
EndFunc
Func _GDIPlus_Startup($sGDIPDLL = Default, $bRetDllHandle = False)
$__g_iGDIPRef += 1
If $__g_iGDIPRef > 1 Then Return True
If $sGDIPDLL = Default Then $sGDIPDLL = "gdiplus.dll"
$__g_hGDIPDll = DllOpen($sGDIPDLL)
If $__g_hGDIPDll = -1 Then
$__g_iGDIPRef = 0
Return SetError(1, 2, False)
EndIf
Local $sVer = FileGetVersion($sGDIPDLL)
$sVer = StringSplit($sVer, ".")
If $sVer[1] > 5 Then $__g_bGDIP_V1_0 = False
Local $tInput = DllStructCreate($tagGDIPSTARTUPINPUT)
Local $tToken = DllStructCreate("ulong_ptr Data")
DllStructSetData($tInput, "Version", 1)
Local $aResult = DllCall($__g_hGDIPDll, "int", "GdiplusStartup", "struct*", $tToken, "struct*", $tInput, "ptr", 0)
If @error Then Return SetError(@error, @extended, False)
If $aResult[0] Then Return SetError(10, $aResult[0], False)
$__g_iGDIPToken = DllStructGetData($tToken, "Data")
If $bRetDllHandle Then Return $__g_hGDIPDll
Return SetExtended($sVer[1], True)
EndFunc
Func __GDIPlus_ExtractFileExt($sFileName, $bNoDot = True)
Local $iIndex = __GDIPlus_LastDelimiter(".\:", $sFileName)
If($iIndex > 0) And(StringMid($sFileName, $iIndex, 1) = '.') Then
If $bNoDot Then
Return StringMid($sFileName, $iIndex + 1)
Else
Return StringMid($sFileName, $iIndex)
EndIf
Else
Return ""
EndIf
EndFunc
Func __GDIPlus_LastDelimiter($sDelimiters, $sString)
Local $sDelimiter, $iN
For $iI = 1 To StringLen($sDelimiters)
$sDelimiter = StringMid($sDelimiters, $iI, 1)
$iN = StringInStr($sString, $sDelimiter, 0, -1)
If $iN > 0 Then Return $iN
Next
EndFunc
Global $__g_iBMPFormat = $GDIP_PXF24RGB
Global $__g_iJPGQuality = 100
Global $__g_iTIFColorDepth = 24
Global $__g_iTIFCompression = $GDIP_EVTCOMPRESSIONLZW
Global Const $__SCREENCAPTURECONSTANT_SM_CXSCREEN = 0
Global Const $__SCREENCAPTURECONSTANT_SM_CYSCREEN = 1
Global Const $__SCREENCAPTURECONSTANT_SRCCOPY = 0x00CC0020
Func _ScreenCapture_Capture($sFileName = "", $iLeft = 0, $iTop = 0, $iRight = -1, $iBottom = -1, $bCursor = True)
Local $bRet = False
If $iRight = -1 Then $iRight = _WinAPI_GetSystemMetrics($__SCREENCAPTURECONSTANT_SM_CXSCREEN) - 1
If $iBottom = -1 Then $iBottom = _WinAPI_GetSystemMetrics($__SCREENCAPTURECONSTANT_SM_CYSCREEN) - 1
If $iRight < $iLeft Then Return SetError(-1, 0, $bRet)
If $iBottom < $iTop Then Return SetError(-2, 0, $bRet)
Local $iW =($iRight - $iLeft) + 1
Local $iH =($iBottom - $iTop) + 1
Local $hWnd = _WinAPI_GetDesktopWindow()
Local $hDDC = _WinAPI_GetDC($hWnd)
Local $hCDC = _WinAPI_CreateCompatibleDC($hDDC)
Local $hBMP = _WinAPI_CreateCompatibleBitmap($hDDC, $iW, $iH)
_WinAPI_SelectObject($hCDC, $hBMP)
_WinAPI_BitBlt($hCDC, 0, 0, $iW, $iH, $hDDC, $iLeft, $iTop, $__SCREENCAPTURECONSTANT_SRCCOPY)
If $bCursor Then
Local $aCursor = _WinAPI_GetCursorInfo()
If Not @error And $aCursor[1] Then
$bCursor = True
Local $hIcon = _WinAPI_CopyIcon($aCursor[2])
Local $aIcon = _WinAPI_GetIconInfo($hIcon)
If Not @error Then
_WinAPI_DeleteObject($aIcon[4])
If $aIcon[5] <> 0 Then _WinAPI_DeleteObject($aIcon[5])
_WinAPI_DrawIcon($hCDC, $aCursor[3] - $aIcon[2] - $iLeft, $aCursor[4] - $aIcon[3] - $iTop, $hIcon)
EndIf
_WinAPI_DestroyIcon($hIcon)
EndIf
EndIf
_WinAPI_ReleaseDC($hWnd, $hDDC)
_WinAPI_DeleteDC($hCDC)
If $sFileName = "" Then Return $hBMP
$bRet = _ScreenCapture_SaveImage($sFileName, $hBMP, True)
Return SetError(@error, @extended, $bRet)
EndFunc
Func _ScreenCapture_SaveImage($sFileName, $hBitmap, $bFreeBmp = True)
_GDIPlus_Startup()
If @error Then Return SetError(-1, -1, False)
Local $sExt = StringUpper(__GDIPlus_ExtractFileExt($sFileName))
Local $sCLSID = _GDIPlus_EncodersGetCLSID($sExt)
If $sCLSID = "" Then Return SetError(-2, -2, False)
Local $hImage = _GDIPlus_BitmapCreateFromHBITMAP($hBitmap)
If @error Then Return SetError(-3, -3, False)
Local $tData, $tParams
Switch $sExt
Case "BMP"
Local $iX = _GDIPlus_ImageGetWidth($hImage)
Local $iY = _GDIPlus_ImageGetHeight($hImage)
Local $hClone = _GDIPlus_BitmapCloneArea($hImage, 0, 0, $iX, $iY, $__g_iBMPFormat)
_GDIPlus_ImageDispose($hImage)
$hImage = $hClone
Case "JPG", "JPEG"
$tParams = _GDIPlus_ParamInit(1)
$tData = DllStructCreate("int Quality")
DllStructSetData($tData, "Quality", $__g_iJPGQuality)
_GDIPlus_ParamAdd($tParams, $GDIP_EPGQUALITY, 1, $GDIP_EPTLONG, DllStructGetPtr($tData))
Case "TIF", "TIFF"
$tParams = _GDIPlus_ParamInit(2)
$tData = DllStructCreate("int ColorDepth;int Compression")
DllStructSetData($tData, "ColorDepth", $__g_iTIFColorDepth)
DllStructSetData($tData, "Compression", $__g_iTIFCompression)
_GDIPlus_ParamAdd($tParams, $GDIP_EPGCOLORDEPTH, 1, $GDIP_EPTLONG, DllStructGetPtr($tData, "ColorDepth"))
_GDIPlus_ParamAdd($tParams, $GDIP_EPGCOMPRESSION, 1, $GDIP_EPTLONG, DllStructGetPtr($tData, "Compression"))
EndSwitch
Local $pParams = 0
If IsDllStruct($tParams) Then $pParams = $tParams
Local $bRet = _GDIPlus_ImageSaveToFileEx($hImage, $sFileName, $sCLSID, $pParams)
_GDIPlus_ImageDispose($hImage)
If $bFreeBmp Then _WinAPI_DeleteObject($hBitmap)
_GDIPlus_Shutdown()
Return SetError($bRet = False, 0, $bRet)
EndFunc
$sCLSID_UIAutomationClient="{944DE083-8FB8-45CF-BCB7-C477ACB2F897}"
Global Const $sCLSID_CUIAutomation="{FF48DBA4-60EF-4201-AA87-54103EEF594E}"
Global Const $UIA_InvokePatternId=10000
Global Const $UIA_SelectionPatternId=10001
Global Const $UIA_ValuePatternId=10002
Global Const $UIA_RangeValuePatternId=10003
Global Const $UIA_ScrollPatternId=10004
Global Const $UIA_ExpandCollapsePatternId=10005
Global Const $UIA_GridPatternId=10006
Global Const $UIA_GridItemPatternId=10007
Global Const $UIA_MultipleViewPatternId=10008
Global Const $UIA_WindowPatternId=10009
Global Const $UIA_SelectionItemPatternId=10010
Global Const $UIA_DockPatternId=10011
Global Const $UIA_TablePatternId=10012
Global Const $UIA_TextPatternId=10014
Global Const $UIA_TogglePatternId=10015
Global Const $UIA_TransformPatternId=10016
Global Const $UIA_ScrollItemPatternId=10017
Global Const $UIA_LegacyIAccessiblePatternId=10018
Global Const $UIA_ItemContainerPatternId=10019
Global Const $UIA_VirtualizedItemPatternId=10020
Global Const $UIA_SynchronizedInputPatternId=10021
Global Const $UIA_RuntimeIdPropertyId=30000
Global Const $UIA_BoundingRectanglePropertyId=30001
Global Const $UIA_ProcessIdPropertyId=30002
Global Const $UIA_ControlTypePropertyId=30003
Global Const $UIA_LocalizedControlTypePropertyId=30004
Global Const $UIA_NamePropertyId=30005
Global Const $UIA_AcceleratorKeyPropertyId=30006
Global Const $UIA_AccessKeyPropertyId=30007
Global Const $UIA_HasKeyboardFocusPropertyId=30008
Global Const $UIA_IsKeyboardFocusablePropertyId=30009
Global Const $UIA_IsEnabledPropertyId=30010
Global Const $UIA_AutomationIdPropertyId=30011
Global Const $UIA_ClassNamePropertyId=30012
Global Const $UIA_HelpTextPropertyId=30013
Global Const $UIA_ClickablePointPropertyId=30014
Global Const $UIA_CulturePropertyId=30015
Global Const $UIA_IsControlElementPropertyId=30016
Global Const $UIA_IsContentElementPropertyId=30017
Global Const $UIA_LabeledByPropertyId=30018
Global Const $UIA_IsPasswordPropertyId=30019
Global Const $UIA_NativeWindowHandlePropertyId=30020
Global Const $UIA_ItemTypePropertyId=30021
Global Const $UIA_IsOffscreenPropertyId=30022
Global Const $UIA_OrientationPropertyId=30023
Global Const $UIA_FrameworkIdPropertyId=30024
Global Const $UIA_IsRequiredForFormPropertyId=30025
Global Const $UIA_ItemStatusPropertyId=30026
Global Const $UIA_IsDockPatternAvailablePropertyId=30027
Global Const $UIA_IsExpandCollapsePatternAvailablePropertyId=30028
Global Const $UIA_IsGridItemPatternAvailablePropertyId=30029
Global Const $UIA_IsGridPatternAvailablePropertyId=30030
Global Const $UIA_IsInvokePatternAvailablePropertyId=30031
Global Const $UIA_IsMultipleViewPatternAvailablePropertyId=30032
Global Const $UIA_IsRangeValuePatternAvailablePropertyId=30033
Global Const $UIA_IsScrollPatternAvailablePropertyId=30034
Global Const $UIA_IsScrollItemPatternAvailablePropertyId=30035
Global Const $UIA_IsSelectionItemPatternAvailablePropertyId=30036
Global Const $UIA_IsSelectionPatternAvailablePropertyId=30037
Global Const $UIA_IsTablePatternAvailablePropertyId=30038
Global Const $UIA_IsTableItemPatternAvailablePropertyId=30039
Global Const $UIA_IsTextPatternAvailablePropertyId=30040
Global Const $UIA_IsTogglePatternAvailablePropertyId=30041
Global Const $UIA_IsTransformPatternAvailablePropertyId=30042
Global Const $UIA_IsValuePatternAvailablePropertyId=30043
Global Const $UIA_IsWindowPatternAvailablePropertyId=30044
Global Const $UIA_ValueValuePropertyId=30045
Global Const $UIA_ValueIsReadOnlyPropertyId=30046
Global Const $UIA_RangeValueValuePropertyId=30047
Global Const $UIA_RangeValueIsReadOnlyPropertyId=30048
Global Const $UIA_RangeValueMinimumPropertyId=30049
Global Const $UIA_RangeValueMaximumPropertyId=30050
Global Const $UIA_RangeValueLargeChangePropertyId=30051
Global Const $UIA_RangeValueSmallChangePropertyId=30052
Global Const $UIA_ScrollHorizontalScrollPercentPropertyId=30053
Global Const $UIA_ScrollHorizontalViewSizePropertyId=30054
Global Const $UIA_ScrollVerticalScrollPercentPropertyId=30055
Global Const $UIA_ScrollVerticalViewSizePropertyId=30056
Global Const $UIA_ScrollHorizontallyScrollablePropertyId=30057
Global Const $UIA_ScrollVerticallyScrollablePropertyId=30058
Global Const $UIA_SelectionSelectionPropertyId=30059
Global Const $UIA_SelectionCanSelectMultiplePropertyId=30060
Global Const $UIA_SelectionIsSelectionRequiredPropertyId=30061
Global Const $UIA_GridRowCountPropertyId=30062
Global Const $UIA_GridColumnCountPropertyId=30063
Global Const $UIA_GridItemRowPropertyId=30064
Global Const $UIA_GridItemColumnPropertyId=30065
Global Const $UIA_GridItemRowSpanPropertyId=30066
Global Const $UIA_GridItemColumnSpanPropertyId=30067
Global Const $UIA_GridItemContainingGridPropertyId=30068
Global Const $UIA_DockDockPositionPropertyId=30069
Global Const $UIA_ExpandCollapseExpandCollapseStatePropertyId=30070
Global Const $UIA_MultipleViewCurrentViewPropertyId=30071
Global Const $UIA_MultipleViewSupportedViewsPropertyId=30072
Global Const $UIA_WindowCanMaximizePropertyId=30073
Global Const $UIA_WindowCanMinimizePropertyId=30074
Global Const $UIA_WindowWindowVisualStatePropertyId=30075
Global Const $UIA_WindowWindowInteractionStatePropertyId=30076
Global Const $UIA_WindowIsModalPropertyId=30077
Global Const $UIA_WindowIsTopmostPropertyId=30078
Global Const $UIA_SelectionItemIsSelectedPropertyId=30079
Global Const $UIA_SelectionItemSelectionContainerPropertyId=30080
Global Const $UIA_TableRowHeadersPropertyId=30081
Global Const $UIA_TableColumnHeadersPropertyId=30082
Global Const $UIA_TableRowOrColumnMajorPropertyId=30083
Global Const $UIA_TableItemRowHeaderItemsPropertyId=30084
Global Const $UIA_TableItemColumnHeaderItemsPropertyId=30085
Global Const $UIA_ToggleToggleStatePropertyId=30086
Global Const $UIA_TransformCanMovePropertyId=30087
Global Const $UIA_TransformCanResizePropertyId=30088
Global Const $UIA_TransformCanRotatePropertyId=30089
Global Const $UIA_IsLegacyIAccessiblePatternAvailablePropertyId=30090
Global Const $UIA_LegacyIAccessibleChildIdPropertyId=30091
Global Const $UIA_LegacyIAccessibleNamePropertyId=30092
Global Const $UIA_LegacyIAccessibleValuePropertyId=30093
Global Const $UIA_LegacyIAccessibleDescriptionPropertyId=30094
Global Const $UIA_LegacyIAccessibleRolePropertyId=30095
Global Const $UIA_LegacyIAccessibleStatePropertyId=30096
Global Const $UIA_LegacyIAccessibleHelpPropertyId=30097
Global Const $UIA_LegacyIAccessibleKeyboardShortcutPropertyId=30098
Global Const $UIA_LegacyIAccessibleSelectionPropertyId=30099
Global Const $UIA_LegacyIAccessibleDefaultActionPropertyId=30100
Global Const $UIA_AriaRolePropertyId=30101
Global Const $UIA_AriaPropertiesPropertyId=30102
Global Const $UIA_IsDataValidForFormPropertyId=30103
Global Const $UIA_ControllerForPropertyId=30104
Global Const $UIA_DescribedByPropertyId=30105
Global Const $UIA_FlowsToPropertyId=30106
Global Const $UIA_ProviderDescriptionPropertyId=30107
Global Const $UIA_IsItemContainerPatternAvailablePropertyId=30108
Global Const $UIA_IsVirtualizedItemPatternAvailablePropertyId=30109
Global Const $UIA_IsSynchronizedInputPatternAvailablePropertyId=30110
Global Const $UIA_WindowControlTypeId=50032
Global Const $TreeScope_Children=2
Global Const $TreeScope_Subtree=7
Global Const $WindowVisualState_Normal=0
Global Const $WindowVisualState_Maximized=1
Global Const $WindowVisualState_Minimized=2
Global Const $sIID_IUIAutomationElement="{D22108AA-8AC5-49A5-837B-37BBB3D7591E}"
Global $dtagIUIAutomationElement = "SetFocus hresult();" & "GetRuntimeId hresult(ptr*);" & "FindFirst hresult(long;ptr;ptr*);" & "FindAll hresult(long;ptr;ptr*);" & "FindFirstBuildCache hresult(long;ptr;ptr;ptr*);" & "FindAllBuildCache hresult(long;ptr;ptr;ptr*);" & "BuildUpdatedCache hresult(ptr;ptr*);" & "GetCurrentPropertyValue hresult(int;variant*);" & "GetCurrentPropertyValueEx hresult(int;long;variant*);" & "GetCachedPropertyValue hresult(int;variant*);" & "GetCachedPropertyValueEx hresult(int;long;variant*);" & "GetCurrentPatternAs hresult(int;none;none*);" & "GetCachedPatternAs hresult(int;none;none*);" & "GetCurrentPattern hresult(int;ptr*);" & "GetCachedPattern hresult(int;ptr*);" & "GetCachedParent hresult(ptr*);" & "GetCachedChildren hresult(ptr*);" & "CurrentProcessId hresult(int*);" & "CurrentControlType hresult(int*);" & "CurrentLocalizedControlType hresult(bstr*);" & "CurrentName hresult(bstr*);" & "CurrentAcceleratorKey hresult(bstr*);" & "CurrentAccessKey hresult(bstr*);" & "CurrentHasKeyboardFocus hresult(long*);" & "CurrentIsKeyboardFocusable hresult(long*);" & "CurrentIsEnabled hresult(long*);" & "CurrentAutomationId hresult(bstr*);" & "CurrentClassName hresult(bstr*);" & "CurrentHelpText hresult(bstr*);" & "CurrentCulture hresult(int*);" & "CurrentIsControlElement hresult(long*);" & "CurrentIsContentElement hresult(long*);" & "CurrentIsPassword hresult(long*);" & "CurrentNativeWindowHandle hresult(hwnd*);" & "CurrentItemType hresult(bstr*);" & "CurrentIsOffscreen hresult(long*);" & "CurrentOrientation hresult(long*);" & "CurrentFrameworkId hresult(bstr*);" & "CurrentIsRequiredForForm hresult(long*);" & "CurrentItemStatus hresult(bstr*);" & "CurrentBoundingRectangle hresult(struct*);" & "CurrentLabeledBy hresult(ptr*);" & "CurrentAriaRole hresult(bstr*);" & "CurrentAriaProperties hresult(bstr*);" & "CurrentIsDataValidForForm hresult(long*);" & "CurrentControllerFor hresult(ptr*);" & "CurrentDescribedBy hresult(ptr*);" & "CurrentFlowsTo hresult(ptr*);" & "CurrentProviderDescription hresult(bstr*);" & "CachedProcessId hresult(int*);" & "CachedControlType hresult(int*);" & "CachedLocalizedControlType hresult(bstr*);" & "CachedName hresult(bstr*);" & "CachedAcceleratorKey hresult(bstr*);" & "CachedAccessKey hresult(bstr*);" & "CachedHasKeyboardFocus hresult(long*);" & "CachedIsKeyboardFocusable hresult(long*);" & "CachedIsEnabled hresult(long*);" & "CachedAutomationId hresult(bstr*);" & "CachedClassName hresult(bstr*);" & "CachedHelpText hresult(bstr*);" & "CachedCulture hresult(int*);" & "CachedIsControlElement hresult(long*);" & "CachedIsContentElement hresult(long*);" & "CachedIsPassword hresult(long*);" & "CachedNativeWindowHandle hresult(hwnd*);" & "CachedItemType hresult(bstr*);" & "CachedIsOffscreen hresult(long*);" & "CachedOrientation hresult(long*);" & "CachedFrameworkId hresult(bstr*);" & "CachedIsRequiredForForm hresult(long*);" & "CachedItemStatus hresult(bstr*);" & "CachedBoundingRectangle hresult(struct*);" & "CachedLabeledBy hresult(ptr*);" & "CachedAriaRole hresult(bstr*);" & "CachedAriaProperties hresult(bstr*);" & "CachedIsDataValidForForm hresult(long*);" & "CachedControllerFor hresult(ptr*);" & "CachedDescribedBy hresult(ptr*);" & "CachedFlowsTo hresult(ptr*);" & "CachedProviderDescription hresult(bstr*);" & "GetClickablePoint hresult(struct*;long*);"
Global Const $sIID_IUIAutomationCondition="{352FFBA8-0973-437C-A61F-F64CAFD81DF9}"
Global $dtagIUIAutomationCondition = ""
Global Const $sIID_IUIAutomationElementArray="{14314595-B4BC-4055-95F2-58F2E42C9855}"
Global $dtagIUIAutomationElementArray = "Length hresult(int*);" & "GetElement hresult(int;ptr*);"
Global $dtagIUIAutomationCacheRequest = "AddProperty hresult(int);" & "AddPattern hresult(int);" & "Clone hresult(ptr*);" & "get_TreeScope hresult(long*);" & "put_TreeScope hresult(long);" & "get_TreeFilter hresult(ptr*);" & "put_TreeFilter hresult(ptr);" & "get_AutomationElementMode hresult(long*);" & "put_AutomationElementMode hresult(long);"
Global $dtagIUIAutomationBoolCondition = "BooleanValue hresult(long*);"
Global $dtagIUIAutomationPropertyCondition = "propertyId hresult(int*);" & "PropertyValue hresult(variant*);" & "PropertyConditionFlags hresult(long*);"
Global $dtagIUIAutomationAndCondition = "ChildCount hresult(int*);" & "GetChildrenAsNativeArray hresult(ptr*;int*);" & "GetChildren hresult(ptr*);"
Global $dtagIUIAutomationOrCondition = "ChildCount hresult(int*);" & "GetChildrenAsNativeArray hresult(ptr*;int*);" & "GetChildren hresult(ptr*);"
Global $dtagIUIAutomationNotCondition = "GetChild hresult(ptr*);"
Global Const $sIID_IUIAutomationTreeWalker="{4042C624-389C-4AFC-A630-9DF854A541FC}"
Global $dtagIUIAutomationTreeWalker = "GetParentElement hresult(ptr;ptr*);" & "GetFirstChildElement hresult(ptr;ptr*);" & "GetLastChildElement hresult(ptr;ptr*);" & "GetNextSiblingElement hresult(ptr;ptr*);" & "GetPreviousSiblingElement hresult(ptr;ptr*);" & "NormalizeElement hresult(ptr;ptr*);" & "GetParentElementBuildCache hresult(ptr;ptr;ptr*);" & "GetFirstChildElementBuildCache hresult(ptr;ptr;ptr*);" & "GetLastChildElementBuildCache hresult(ptr;ptr;ptr*);" & "GetNextSiblingElementBuildCache hresult(ptr;ptr;ptr*);" & "GetPreviousSiblingElementBuildCache hresult(ptr;ptr;ptr*);" & "NormalizeElementBuildCache hresult(ptr;ptr;ptr*);" & "condition hresult(ptr*);"
Global $dtagIUIAutomationEventHandler = "HandleAutomationEvent hresult(ptr;int);"
Global $dtagIUIAutomationPropertyChangedEventHandler = "HandlePropertyChangedEvent hresult(ptr;int;variant);"
Global $dtagIUIAutomationStructureChangedEventHandler = "HandleStructureChangedEvent hresult(ptr;long;ptr);"
Global $dtagIUIAutomationFocusChangedEventHandler = "HandleFocusChangedEvent hresult(ptr);"
Global Const $sIID_IUIAutomationInvokePattern="{FB377FBE-8EA6-46D5-9C73-6499642D3059}"
Global $dtagIUIAutomationInvokePattern = "Invoke hresult();"
Global Const $sIID_IUIAutomationDockPattern="{FDE5EF97-1464-48F6-90BF-43D0948E86EC}"
Global $dtagIUIAutomationDockPattern = "SetDockPosition hresult(long);" & "CurrentDockPosition hresult(long*);" & "CachedDockPosition hresult(long*);"
Global Const $sIID_IUIAutomationExpandCollapsePattern="{619BE086-1F4E-4EE4-BAFA-210128738730}"
Global $dtagIUIAutomationExpandCollapsePattern = "Expand hresult();" & "Collapse hresult();" & "CurrentExpandCollapseState hresult(long*);" & "CachedExpandCollapseState hresult(long*);"
Global Const $sIID_IUIAutomationGridPattern="{414C3CDC-856B-4F5B-8538-3131C6302550}"
Global $dtagIUIAutomationGridPattern = "GetItem hresult(int;int;ptr*);" & "CurrentRowCount hresult(int*);" & "CurrentColumnCount hresult(int*);" & "CachedRowCount hresult(int*);" & "CachedColumnCount hresult(int*);"
Global Const $sIID_IUIAutomationGridItemPattern="{78F8EF57-66C3-4E09-BD7C-E79B2004894D}"
Global $dtagIUIAutomationGridItemPattern = "CurrentContainingGrid hresult(ptr*);" & "CurrentRow hresult(int*);" & "CurrentColumn hresult(int*);" & "CurrentRowSpan hresult(int*);" & "CurrentColumnSpan hresult(int*);" & "CachedContainingGrid hresult(ptr*);" & "CachedRow hresult(int*);" & "CachedColumn hresult(int*);" & "CachedRowSpan hresult(int*);" & "CachedColumnSpan hresult(int*);"
Global Const $sIID_IUIAutomationMultipleViewPattern="{8D253C91-1DC5-4BB5-B18F-ADE16FA495E8}"
Global $dtagIUIAutomationMultipleViewPattern = "GetViewName hresult(int;bstr*);" & "SetCurrentView hresult(int);" & "CurrentCurrentView hresult(int*);" & "GetCurrentSupportedViews hresult(ptr*);" & "CachedCurrentView hresult(int*);" & "GetCachedSupportedViews hresult(ptr*);"
Global Const $sIID_IUIAutomationRangeValuePattern="{59213F4F-7346-49E5-B120-80555987A148}"
Global $dtagIUIAutomationRangeValuePattern = "SetValue hresult(ushort);" & "CurrentValue hresult(ushort*);" & "CurrentIsReadOnly hresult(long*);" & "CurrentMaximum hresult(ushort*);" & "CurrentMinimum hresult(ushort*);" & "CurrentLargeChange hresult(ushort*);" & "CurrentSmallChange hresult(ushort*);" & "CachedValue hresult(ushort*);" & "CachedIsReadOnly hresult(long*);" & "CachedMaximum hresult(ushort*);" & "CachedMinimum hresult(ushort*);" & "CachedLargeChange hresult(ushort*);" & "CachedSmallChange hresult(ushort*);"
Global Const $sIID_IUIAutomationScrollPattern="{88F4D42A-E881-459D-A77C-73BBBB7E02DC}"
Global $dtagIUIAutomationScrollPattern = "Scroll hresult(long;long);" & "SetScrollPercent hresult(ushort;ushort);" & "CurrentHorizontalScrollPercent hresult(ushort*);" & "CurrentVerticalScrollPercent hresult(ushort*);" & "CurrentHorizontalViewSize hresult(ushort*);" & "CurrentVerticalViewSize hresult(ushort*);" & "CurrentHorizontallyScrollable hresult(long*);" & "CurrentVerticallyScrollable hresult(long*);" & "CachedHorizontalScrollPercent hresult(ushort*);" & "CachedVerticalScrollPercent hresult(ushort*);" & "CachedHorizontalViewSize hresult(ushort*);" & "CachedVerticalViewSize hresult(ushort*);" & "CachedHorizontallyScrollable hresult(long*);" & "CachedVerticallyScrollable hresult(long*);"
Global Const $sIID_IUIAutomationScrollItemPattern="{B488300F-D015-4F19-9C29-BB595E3645EF}"
Global $dtagIUIAutomationScrollItemPattern = "ScrollIntoView hresult();"
Global Const $sIID_IUIAutomationSelectionPattern="{5ED5202E-B2AC-47A6-B638-4B0BF140D78E}"
Global $dtagIUIAutomationSelectionPattern = "GetCurrentSelection hresult(ptr*);" & "CurrentCanSelectMultiple hresult(long*);" & "CurrentIsSelectionRequired hresult(long*);" & "GetCachedSelection hresult(ptr*);" & "CachedCanSelectMultiple hresult(long*);" & "CachedIsSelectionRequired hresult(long*);"
Global Const $sIID_IUIAutomationSelectionItemPattern="{A8EFA66A-0FDA-421A-9194-38021F3578EA}"
Global $dtagIUIAutomationSelectionItemPattern = "Select hresult();" & "AddToSelection hresult();" & "RemoveFromSelection hresult();" & "CurrentIsSelected hresult(long*);" & "CurrentSelectionContainer hresult(ptr*);" & "CachedIsSelected hresult(long*);" & "CachedSelectionContainer hresult(ptr*);"
Global Const $sIID_IUIAutomationSynchronizedInputPattern="{2233BE0B-AFB7-448B-9FDA-3B378AA5EAE1}"
Global $dtagIUIAutomationSynchronizedInputPattern = "StartListening hresult(long);" & "Cancel hresult();"
Global Const $sIID_IUIAutomationTablePattern="{620E691C-EA96-4710-A850-754B24CE2417}"
Global $dtagIUIAutomationTablePattern = "GetCurrentRowHeaders hresult(ptr*);" & "GetCurrentColumnHeaders hresult(ptr*);" & "CurrentRowOrColumnMajor hresult(long*);" & "GetCachedRowHeaders hresult(ptr*);" & "GetCachedColumnHeaders hresult(ptr*);" & "CachedRowOrColumnMajor hresult(long*);"
Global $dtagIUIAutomationTableItemPattern = "GetCurrentRowHeaderItems hresult(ptr*);" & "GetCurrentColumnHeaderItems hresult(ptr*);" & "GetCachedRowHeaderItems hresult(ptr*);" & "GetCachedColumnHeaderItems hresult(ptr*);"
Global Const $sIID_IUIAutomationTogglePattern="{94CF8058-9B8D-4AB9-8BFD-4CD0A33C8C70}"
Global $dtagIUIAutomationTogglePattern = "Toggle hresult();" & "CurrentToggleState hresult(long*);" & "CachedToggleState hresult(long*);"
Global Const $sIID_IUIAutomationTransformPattern="{A9B55844-A55D-4EF0-926D-569C16FF89BB}"
Global $dtagIUIAutomationTransformPattern = "Move hresult(double;double);" & "Resize hresult(double;double);" & "Rotate hresult(ushort);" & "CurrentCanMove hresult(long*);" & "CurrentCanResize hresult(long*);" & "CurrentCanRotate hresult(long*);" & "CachedCanMove hresult(long*);" & "CachedCanResize hresult(long*);" & "CachedCanRotate hresult(long*);"
Global Const $sIID_IUIAutomationValuePattern="{A94CD8B1-0844-4CD6-9D2D-640537AB39E9}"
Global $dtagIUIAutomationValuePattern = "SetValue hresult(bstr);" & "CurrentValue hresult(bstr*);" & "CurrentIsReadOnly hresult(long*);" & "CachedValue hresult(bstr*);" & "CachedIsReadOnly hresult(long*);"
Global Const $sIID_IUIAutomationWindowPattern="{0FAEF453-9208-43EF-BBB2-3B485177864F}"
Global $dtagIUIAutomationWindowPattern = "Close hresult();" & "WaitForInputIdle hresult(int;long*);" & "SetWindowVisualState hresult(long);" & "CurrentCanMaximize hresult(long*);" & "CurrentCanMinimize hresult(long*);" & "CurrentIsModal hresult(long*);" & "CurrentIsTopmost hresult(long*);" & "CurrentWindowVisualState hresult(long*);" & "CurrentWindowInteractionState hresult(long*);" & "CachedCanMaximize hresult(long*);" & "CachedCanMinimize hresult(long*);" & "CachedIsModal hresult(long*);" & "CachedIsTopmost hresult(long*);" & "CachedWindowVisualState hresult(long*);" & "CachedWindowInteractionState hresult(long*);"
Global $dtagIUIAutomationTextRange = "Clone hresult(ptr*);" & "Compare hresult(ptr;long*);" & "CompareEndpoints hresult(long;ptr;long;int*);" & "ExpandToEnclosingUnit hresult(long);" & "FindAttribute hresult(int;variant;long;ptr*);" & "FindText hresult(bstr;long;long;ptr*);" & "GetAttributeValue hresult(int;variant*);" & "GetBoundingRectangles hresult(ptr*);" & "GetEnclosingElement hresult(ptr*);" & "GetText hresult(int;bstr*);" & "Move hresult(long;int;int*);" & "MoveEndpointByUnit hresult(long;long;int;int*);" & "MoveEndpointByRange hresult(long;ptr;long);" & "Select hresult();" & "AddToSelection hresult();" & "RemoveFromSelection hresult();" & "ScrollIntoView hresult(long);" & "GetChildren hresult(ptr*);"
Global $dtagIUIAutomationTextRangeArray = "Length hresult(int*);" & "GetElement hresult(int;ptr*);"
Global Const $sIID_IUIAutomationTextPattern="{32EBA289-3583-42C9-9C59-3B6D9A1E9B6A}"
Global $dtagIUIAutomationTextPattern = "RangeFromPoint hresult(struct;ptr*);" & "RangeFromChild hresult(ptr;ptr*);" & "GetSelection hresult(ptr*);" & "GetVisibleRanges hresult(ptr*);" & "DocumentRange hresult(ptr*);" & "SupportedTextSelection hresult(long*);"
Global Const $sIID_IUIAutomationLegacyIAccessiblePattern="{828055AD-355B-4435-86D5-3B51C14A9B1B}"
Global $dtagIUIAutomationLegacyIAccessiblePattern = "Select hresult(long);" & "DoDefaultAction hresult();" & "SetValue hresult(wstr);" & "CurrentChildId hresult(int*);" & "CurrentName hresult(bstr*);" & "CurrentValue hresult(bstr*);" & "CurrentDescription hresult(bstr*);" & "CurrentRole hresult(uint*);" & "CurrentState hresult(uint*);" & "CurrentHelp hresult(bstr*);" & "CurrentKeyboardShortcut hresult(bstr*);" & "GetCurrentSelection hresult(ptr*);" & "CurrentDefaultAction hresult(bstr*);" & "CachedChildId hresult(int*);" & "CachedName hresult(bstr*);" & "CachedValue hresult(bstr*);" & "CachedDescription hresult(bstr*);" & "CachedRole hresult(uint*);" & "CachedState hresult(uint*);" & "CachedHelp hresult(bstr*);" & "CachedKeyboardShortcut hresult(bstr*);" & "GetCachedSelection hresult(ptr*);" & "CachedDefaultAction hresult(bstr*);" & "GetIAccessible hresult(idispatch*);"
Global $dtagIAccessible = "GetTypeInfoCount hresult(uint*);" & "GetTypeInfo hresult(uint;int;ptr*);" & "GetIDsOfNames hresult(struct*;wstr;uint;int;int);" & "Invoke hresult(int;struct*;int;word;ptr*;ptr*;ptr*;uint*);" & "get_accParent hresult(ptr*);" & "get_accChildCount hresult(long*);" & "get_accChild hresult(variant;idispatch*);" & "get_accName hresult(variant;bstr*);" & "get_accValue hresult(variant;bstr*);" & "get_accDescription hresult(variant;bstr*);" & "get_accRole hresult(variant;variant*);" & "get_accState hresult(variant;variant*);" & "get_accHelp hresult(variant;bstr*);" & "get_accHelpTopic hresult(bstr*;variant;long*);" & "get_accKeyboardShortcut hresult(variant;bstr*);" & "get_accFocus hresult(struct*);" & "get_accSelection hresult(variant*);" & "get_accDefaultAction hresult(variant;bstr*);" & "accSelect hresult(long;variant);" & "accLocation hresult(long*;long*;long*;long*;variant);" & "accNavigate hresult(long;variant;variant*);" & "accHitTest hresult(long;long;variant*);" & "accDoDefaultAction hresult(variant);" & "put_accName hresult(variant;bstr);" & "put_accValue hresult(variant;bstr);"
Global Const $sIID_IUIAutomationItemContainerPattern="{C690FDB2-27A8-423C-812D-429773C9084E}"
Global $dtagIUIAutomationItemContainerPattern = "FindItemByProperty hresult(ptr;int;variant;ptr*);"
Global Const $sIID_IUIAutomationVirtualizedItemPattern="{6BA3D7A6-04CF-4F11-8793-A8D1CDE9969F}"
Global $dtagIUIAutomationVirtualizedItemPattern = "Realize hresult();"
Global $dtagIUIAutomationProxyFactory = "CreateProvider hresult(hwnd;long;long;ptr*);" & "ProxyFactoryId hresult(bstr*);"
Global $dtagIRawElementProviderSimple = "ProviderOptions hresult(long*);" & "GetPatternProvider hresult(int;ptr*);" & "GetPropertyValue hresult(int;variant*);" & "HostRawElementProvider hresult(ptr*);"
Global $dtagIUIAutomationProxyFactoryEntry = "ProxyFactory hresult(ptr*);" & "ClassName hresult(bstr*);" & "ImageName hresult(bstr*);" & "AllowSubstringMatch hresult(long*);" & "CanCheckBaseClass hresult(long*);" & "NeedsAdviseEvents hresult(long*);" & "ClassName hresult(wstr);" & "ImageName hresult(wstr);" & "AllowSubstringMatch hresult(long);" & "CanCheckBaseClass hresult(long);" & "NeedsAdviseEvents hresult(long);" & "SetWinEventsForAutomationEvent hresult(int;int;ptr);" & "GetWinEventsForAutomationEvent hresult(int;int;ptr*);"
Global $dtagIUIAutomationProxyFactoryMapping = "count hresult(uint*);" & "GetTable hresult(ptr*);" & "GetEntry hresult(uint;ptr*);" & "SetTable hresult(ptr);" & "InsertEntries hresult(uint;ptr);" & "InsertEntry hresult(uint;ptr);" & "RemoveEntry hresult(uint);" & "ClearTable hresult();" & "RestoreDefaultTable hresult();"
Global Const $sIID_IUIAutomation="{30CBE57D-D9D0-452A-AB13-7AC5AC4825EE}"
Global $dtagIUIAutomation = "CompareElements hresult(ptr;ptr;long*);" & "CompareRuntimeIds hresult(ptr;ptr;long*);" & "GetRootElement hresult(ptr*);" & "ElementFromHandle hresult(hwnd;ptr*);" & "ElementFromPoint hresult(struct;ptr*);" & "GetFocusedElement hresult(ptr*);" & "GetRootElementBuildCache hresult(ptr;ptr*);" & "ElementFromHandleBuildCache hresult(hwnd;ptr;ptr*);" & "ElementFromPointBuildCache hresult(struct;ptr;ptr*);" & "GetFocusedElementBuildCache hresult(ptr;ptr*);" & "CreateTreeWalker hresult(ptr;ptr*);" & "ControlViewWalker hresult(ptr*);" & "ContentViewWalker hresult(ptr*);" & "RawViewWalker hresult(ptr*);" & "RawViewCondition hresult(ptr*);" & "ControlViewCondition hresult(ptr*);" & "ContentViewCondition hresult(ptr*);" & "CreateCacheRequest hresult(ptr*);" & "CreateTrueCondition hresult(ptr*);" & "CreateFalseCondition hresult(ptr*);" & "CreatePropertyCondition hresult(int;variant;ptr*);" & "CreatePropertyConditionEx hresult(int;variant;long;ptr*);" & "CreateAndCondition hresult(ptr;ptr;ptr*);" & "CreateAndConditionFromArray hresult(ptr;ptr*);" & "CreateAndConditionFromNativeArray hresult(ptr;int;ptr*);" & "CreateOrCondition hresult(ptr;ptr;ptr*);" & "CreateOrConditionFromArray hresult(ptr;ptr*);" & "CreateOrConditionFromNativeArray hresult(ptr;int;ptr*);" & "CreateNotCondition hresult(ptr;ptr*);" & "AddAutomationEventHandler hresult(int;ptr;long;ptr;ptr);" & "RemoveAutomationEventHandler hresult(int;ptr;ptr);" & "AddPropertyChangedEventHandlerNativeArray hresult(ptr;long;ptr;ptr;struct*;int);" & "AddPropertyChangedEventHandler hresult(ptr;long;ptr;ptr;ptr);" & "RemovePropertyChangedEventHandler hresult(ptr;ptr);" & "AddStructureChangedEventHandler hresult(ptr;long;ptr;ptr);" & "RemoveStructureChangedEventHandler hresult(ptr;ptr);" & "AddFocusChangedEventHandler hresult(ptr;ptr);" & "RemoveFocusChangedEventHandler hresult(ptr);" & "RemoveAllEventHandlers hresult();" & "IntNativeArrayToSafeArray hresult(int;int;ptr*);" & "IntSafeArrayToNativeArray hresult(ptr;int*;int*);" & "RectToVariant hresult(struct;variant*);" & "VariantToRect hresult(variant;struct*);" & "SafeArrayToRectNativeArray hresult(ptr;struct*;int*);" & "CreateProxyFactoryEntry hresult(ptr;ptr*);" & "ProxyFactoryMapping hresult(ptr*);" & "GetPropertyProgrammaticName hresult(int;bstr*);" & "GetPatternProgrammaticName hresult(int;bstr*);" & "PollForPotentialSupportedPatterns hresult(ptr;ptr*;ptr*);" & "PollForPotentialSupportedProperties hresult(ptr;ptr*;ptr*);" & "CheckNotSupported hresult(variant;long*);" & "ReservedNotSupportedValue hresult(ptr*);" & "ReservedMixedAttributeValue hresult(ptr*);" & "ElementFromIAccessible hresult(idispatch;int;ptr*);" & "ElementFromIAccessibleBuildCache hresult(iaccessible;int;ptr;ptr*);"
Global $UIA_oUIAutomation
Global $UIA_oDesktop, $UIA_pDesktop
Global $UIA_oUIElement, $UIA_pUIElement
Global $UIA_oTW, $UIA_pTW
Global $UIA_oTRUECondition
Global $UIA_Vars
Global $UIA_DefaultWaitTime = 200
Global Const $__gaUIAAU3VersionInfo[6] = ["T", 0, 5, 0, "20140912", "T0.5-2"]
Global Const $_UIA_MAXDEPTH=25
Local Const $UIA_CFGFileName = "UIA.CFG"
Local Const $UIA_Log_Wrapper = 5, $UIA_Log_trace = 10, $UIA_Log_debug = 20, $UIA_Log_info = 30, $UIA_Log_warn = 40, $UIA_Log_error = 50, $UIA_Log_fatal = 60
local const $__UIA_debugCacheOn=1
local const $__UIA_debugCacheOff=2
local $__gl_XMLCache
local $__l_UIA_CacheState=false
Global Enum $_UIASTATUS_Success = 0, $_UIASTATUS_GeneralError, $_UIASTATUS_InvalidValue, $_UIASTATUS_NoMatch
_UIA_Init()
Func _UIA_Init()
Local $UIA_pTRUECondition
$UIA_oUIAutomation = ObjCreateInterface($sCLSID_CUIAutomation, $sIID_IUIAutomation, $dtagIUIAutomation)
If _UIA_IsElement($UIA_oUIAutomation) = 0 Then
Return SetError(1, 0, 0)
EndIf
$UIA_oUIAutomation.GetRootElement($UIA_pDesktop)
$UIA_oDesktop = ObjCreateInterface($UIA_pDesktop, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
If IsObj($UIA_oDesktop) = 0 Then
MsgBox(1, "UI automation desktop failed", "Fatal: UI Automation desktop failed")
Return SetError(2, 0, 0)
EndIf
$UIA_oUIAutomation.RawViewWalker($UIA_pTW)
$UIA_oTW = ObjCreateInterface($UIA_pTW, $sIID_IUIAutomationTreeWalker, $dtagIUIAutomationTreeWalker)
If _UIA_IsElement($UIA_oTW) = 0 Then
MsgBox(1, "UI automation treewalker failed", "UI Automation failed to setup treewalker", 10)
EndIf
$UIA_oUIAutomation.CreateTrueCondition($UIA_pTRUECondition)
$UIA_oTRUECondition = ObjCreateInterface($UIA_pTRUECondition, $sIID_IUIAutomationCondition, $dtagIUIAutomationCondition)
Return seterror($_UIASTATUS_Success,$_UIASTATUS_Success,true)
EndFunc
Local $UIA_propertiesSupportedArray[123][2] = [ ["indexrelative", -1], ["index", -1], ["instance", -1], ["title", $UIA_NamePropertyId], ["text", $UIA_NamePropertyId], ["regexptitle", $UIA_NamePropertyId], ["class", $UIA_ClassNamePropertyId], ["regexpclass", $UIA_ClassNamePropertyId], ["iaccessiblevalue", $UIA_LegacyIAccessibleValuePropertyId], ["iaccessiblechildId", $UIA_LegacyIAccessibleChildIdPropertyId], ["id", $UIA_AutomationIdPropertyId], ["handle", $UIA_NativeWindowHandlePropertyId], ["RuntimeId", $UIA_RuntimeIdPropertyId], ["BoundingRectangle", $UIA_BoundingRectanglePropertyId], ["ProcessId", $UIA_ProcessIdPropertyId], ["ControlType", $UIA_ControlTypePropertyId], ["LocalizedControlType", $UIA_LocalizedControlTypePropertyId], ["Name", $UIA_NamePropertyId], ["AcceleratorKey", $UIA_AcceleratorKeyPropertyId], ["AccessKey", $UIA_AccessKeyPropertyId], ["HasKeyboardFocus", $UIA_HasKeyboardFocusPropertyId], ["IsKeyboardFocusable", $UIA_IsKeyboardFocusablePropertyId], ["IsEnabled", $UIA_IsEnabledPropertyId], ["AutomationId", $UIA_AutomationIdPropertyId], ["ClassName", $UIA_ClassNamePropertyId], ["HelpText", $UIA_HelpTextPropertyId], ["ClickablePoint", $UIA_ClickablePointPropertyId], ["Culture", $UIA_CulturePropertyId], ["IsControlElement", $UIA_IsControlElementPropertyId], ["IsContentElement", $UIA_IsContentElementPropertyId], ["LabeledBy", $UIA_LabeledByPropertyId], ["IsPassword", $UIA_IsPasswordPropertyId], ["NativeWindowHandle", $UIA_NativeWindowHandlePropertyId], ["ItemType", $UIA_ItemTypePropertyId], ["IsOffscreen", $UIA_IsOffscreenPropertyId], ["Orientation", $UIA_OrientationPropertyId], ["FrameworkId", $UIA_FrameworkIdPropertyId], ["IsRequiredForForm", $UIA_IsRequiredForFormPropertyId], ["ItemStatus", $UIA_ItemStatusPropertyId], ["IsDockPatternAvailable", $UIA_IsDockPatternAvailablePropertyId], ["IsExpandCollapsePatternAvailable", $UIA_IsExpandCollapsePatternAvailablePropertyId], ["IsGridItemPatternAvailable", $UIA_IsGridItemPatternAvailablePropertyId], ["IsGridPatternAvailable", $UIA_IsGridPatternAvailablePropertyId], ["IsInvokePatternAvailable", $UIA_IsInvokePatternAvailablePropertyId], ["IsMultipleViewPatternAvailable", $UIA_IsMultipleViewPatternAvailablePropertyId], ["IsRangeValuePatternAvailable", $UIA_IsRangeValuePatternAvailablePropertyId], ["IsScrollPatternAvailable", $UIA_IsScrollPatternAvailablePropertyId], ["IsScrollItemPatternAvailable", $UIA_IsScrollItemPatternAvailablePropertyId], ["IsSelectionItemPatternAvailable", $UIA_IsSelectionItemPatternAvailablePropertyId], ["IsSelectionPatternAvailable", $UIA_IsSelectionPatternAvailablePropertyId], ["IsTablePatternAvailable", $UIA_IsTablePatternAvailablePropertyId], ["IsTableItemPatternAvailable", $UIA_IsTableItemPatternAvailablePropertyId], ["IsTextPatternAvailable", $UIA_IsTextPatternAvailablePropertyId], ["IsTogglePatternAvailable", $UIA_IsTogglePatternAvailablePropertyId], ["IsTransformPatternAvailable", $UIA_IsTransformPatternAvailablePropertyId], ["IsValuePatternAvailable", $UIA_IsValuePatternAvailablePropertyId], ["IsWindowPatternAvailable", $UIA_IsWindowPatternAvailablePropertyId], ["ValueValue", $UIA_ValueValuePropertyId], ["ValueIsReadOnly", $UIA_ValueIsReadOnlyPropertyId], ["RangeValueValue", $UIA_RangeValueValuePropertyId], ["RangeValueIsReadOnly", $UIA_RangeValueIsReadOnlyPropertyId], ["RangeValueMinimum", $UIA_RangeValueMinimumPropertyId], ["RangeValueMaximum", $UIA_RangeValueMaximumPropertyId], ["RangeValueLargeChange", $UIA_RangeValueLargeChangePropertyId], ["RangeValueSmallChange", $UIA_RangeValueSmallChangePropertyId], ["ScrollHorizontalScrollPercent", $UIA_ScrollHorizontalScrollPercentPropertyId], ["ScrollHorizontalViewSize", $UIA_ScrollHorizontalViewSizePropertyId], ["ScrollVerticalScrollPercent", $UIA_ScrollVerticalScrollPercentPropertyId], ["ScrollVerticalViewSize", $UIA_ScrollVerticalViewSizePropertyId], ["ScrollHorizontallyScrollable", $UIA_ScrollHorizontallyScrollablePropertyId], ["ScrollVerticallyScrollable", $UIA_ScrollVerticallyScrollablePropertyId], _
["SelectionSelection", $UIA_SelectionSelectionPropertyId], ["SelectionCanSelectMultiple", $UIA_SelectionCanSelectMultiplePropertyId], ["SelectionIsSelectionRequired", $UIA_SelectionIsSelectionRequiredPropertyId], ["GridRowCount", $UIA_GridRowCountPropertyId], ["GridColumnCount", $UIA_GridColumnCountPropertyId], ["GridItemRow", $UIA_GridItemRowPropertyId], ["GridItemColumn", $UIA_GridItemColumnPropertyId], ["GridItemRowSpan", $UIA_GridItemRowSpanPropertyId], ["GridItemColumnSpan", $UIA_GridItemColumnSpanPropertyId], ["GridItemContainingGrid", $UIA_GridItemContainingGridPropertyId], ["DockDockPosition", $UIA_DockDockPositionPropertyId], ["ExpandCollapseExpandCollapseState", $UIA_ExpandCollapseExpandCollapseStatePropertyId], ["MultipleViewCurrentView", $UIA_MultipleViewCurrentViewPropertyId], ["MultipleViewSupportedViews", $UIA_MultipleViewSupportedViewsPropertyId], ["WindowCanMaximize", $UIA_WindowCanMaximizePropertyId], ["WindowCanMinimize", $UIA_WindowCanMinimizePropertyId], ["WindowWindowVisualState", $UIA_WindowWindowVisualStatePropertyId], ["WindowWindowInteractionState", $UIA_WindowWindowInteractionStatePropertyId], ["WindowIsModal", $UIA_WindowIsModalPropertyId], ["WindowIsTopmost", $UIA_WindowIsTopmostPropertyId], ["SelectionItemIsSelected", $UIA_SelectionItemIsSelectedPropertyId], ["SelectionItemSelectionContainer", $UIA_SelectionItemSelectionContainerPropertyId], ["TableRowHeaders", $UIA_TableRowHeadersPropertyId], ["TableColumnHeaders", $UIA_TableColumnHeadersPropertyId], ["TableRowOrColumnMajor", $UIA_TableRowOrColumnMajorPropertyId], ["TableItemRowHeaderItems", $UIA_TableItemRowHeaderItemsPropertyId], ["TableItemColumnHeaderItems", $UIA_TableItemColumnHeaderItemsPropertyId], ["ToggleToggleState", $UIA_ToggleToggleStatePropertyId], ["TransformCanMove", $UIA_TransformCanMovePropertyId], ["TransformCanResize", $UIA_TransformCanResizePropertyId], ["TransformCanRotate", $UIA_TransformCanRotatePropertyId], ["IsLegacyIAccessiblePatternAvailable", $UIA_IsLegacyIAccessiblePatternAvailablePropertyId], ["LegacyIAccessibleChildId", $UIA_LegacyIAccessibleChildIdPropertyId], ["LegacyIAccessibleName", $UIA_LegacyIAccessibleNamePropertyId], ["LegacyIAccessibleValue", $UIA_LegacyIAccessibleValuePropertyId], ["LegacyIAccessibleDescription", $UIA_LegacyIAccessibleDescriptionPropertyId], ["LegacyIAccessibleRole", $UIA_LegacyIAccessibleRolePropertyId], ["LegacyIAccessibleState", $UIA_LegacyIAccessibleStatePropertyId], ["LegacyIAccessibleHelp", $UIA_LegacyIAccessibleHelpPropertyId], ["LegacyIAccessibleKeyboardShortcut", $UIA_LegacyIAccessibleKeyboardShortcutPropertyId], ["LegacyIAccessibleSelection", $UIA_LegacyIAccessibleSelectionPropertyId], ["LegacyIAccessibleDefaultAction", $UIA_LegacyIAccessibleDefaultActionPropertyId], ["AriaRole", $UIA_AriaRolePropertyId], ["AriaProperties", $UIA_AriaPropertiesPropertyId], ["IsDataValidForForm", $UIA_IsDataValidForFormPropertyId], ["ControllerFor", $UIA_ControllerForPropertyId], ["DescribedBy", $UIA_DescribedByPropertyId], ["FlowsTo", $UIA_FlowsToPropertyId], ["ProviderDescription", $UIA_ProviderDescriptionPropertyId], ["IsItemContainerPatternAvailable", $UIA_IsItemContainerPatternAvailablePropertyId], ["IsVirtualizedItemPatternAvailable", $UIA_IsVirtualizedItemPatternAvailablePropertyId], ["IsSynchronizedInputPatternAvailable", $UIA_IsSynchronizedInputPatternAvailablePropertyId] ]
Local $UIA_ControlArray[41][3] = [ ["UIA_AppBarControlTypeId", 50040, "Identifies the AppBar control type. Supported starting with Windows 8.1."], ["UIA_ButtonControlTypeId", 50000, "Identifies the Button control type."], ["UIA_CalendarControlTypeId", 50001, "Identifies the Calendar control type."], ["UIA_CheckBoxControlTypeId", 50002, "Identifies the CheckBox control type."], ["UIA_ComboBoxControlTypeId", 50003, "Identifies the ComboBox control type."], ["UIA_CustomControlTypeId", 50025, "Identifies the Custom control type. For more information, see Custom Properties, Events, and Control Patterns."], ["UIA_DataGridControlTypeId", 50028, "Identifies the DataGrid control type."], ["UIA_DataItemControlTypeId", 50029, "Identifies the DataItem control type."], ["UIA_DocumentControlTypeId", 50030, "Identifies the Document control type."], ["UIA_EditControlTypeId", 50004, "Identifies the Edit control type."], ["UIA_GroupControlTypeId", 50026, "Identifies the Group control type."], ["UIA_HeaderControlTypeId", 50034, "Identifies the Header control type."], ["UIA_HeaderItemControlTypeId", 50035, "Identifies the HeaderItem control type."], ["UIA_HyperlinkControlTypeId", 50005, "Identifies the Hyperlink control type."], ["UIA_ImageControlTypeId", 50006, "Identifies the Image control type."], ["UIA_ListControlTypeId", 50008, "Identifies the List control type."], ["UIA_ListItemControlTypeId", 50007, "Identifies the ListItem control type."], ["UIA_MenuBarControlTypeId", 50010, "Identifies the MenuBar control type."], ["UIA_MenuControlTypeId", 50009, "Identifies the Menu control type."], ["UIA_MenuItemControlTypeId", 50011, "Identifies the MenuItem control type."], ["UIA_PaneControlTypeId", 50033, "Identifies the Pane control type."], ["UIA_ProgressBarControlTypeId", 50012, "Identifies the ProgressBar control type."], ["UIA_RadioButtonControlTypeId", 50013, "Identifies the RadioButton control type."], ["UIA_ScrollBarControlTypeId", 50014, "Identifies the ScrollBar control type."], ["UIA_SemanticZoomControlTypeId", 50039, "Identifies the SemanticZoom control type. Supported starting with Windows 8."], ["UIA_SeparatorControlTypeId", 50038, "Identifies the Separator control type."], ["UIA_SliderControlTypeId", 50015, "Identifies the Slider control type."], ["UIA_SpinnerControlTypeId", 50016, "Identifies the Spinner control type."], ["UIA_SplitButtonControlTypeId", 50031, "Identifies the SplitButton control type."], ["UIA_StatusBarControlTypeId", 50017, "Identifies the StatusBar control type."], ["UIA_TabControlTypeId", 50018, "Identifies the Tab control type."], ["UIA_TabItemControlTypeId", 50019, "Identifies the TabItem control type."], ["UIA_TableControlTypeId", 50036, "Identifies the Table control type."], ["UIA_TextControlTypeId", 50020, "Identifies the Text control type."], ["UIA_ThumbControlTypeId", 50027, "Identifies the Thumb control type."], ["UIA_TitleBarControlTypeId", 50037, "Identifies the TitleBar control type."], ["UIA_ToolBarControlTypeId", 50021, "Identifies the ToolBar control type."], ["UIA_ToolTipControlTypeId", 50022, "Identifies the ToolTip control type."], ["UIA_TreeControlTypeId", 50023, "Identifies the Tree control type."], ["UIA_TreeItemControlTypeId", 50024, "Identifies the TreeItem control type."], ["UIA_WindowControlTypeId", 50032, "Identifies the Window control type."] ]
Func _UIA_getControlName($controlID)
Local $i
For $i = 0 To UBound($UIA_ControlArray) - 1
If($UIA_ControlArray[$i][1] = $controlID) Then
Return $UIA_ControlArray[$i][0]
EndIf
Next
Return seterror($_UIASTATUS_GeneralError,$_UIASTATUS_GeneralError,"No control with that id")
EndFunc
Func _UIA_getControlID($controlName)
Local $tName, $i
$tName = StringUpper($controlName)
If StringLeft($tName, 3) <> "UIA" Then
$tName = "UIA_" & $tName & "CONTROLTYPEID"
EndIf
For $i = 0 To UBound($UIA_ControlArray) - 1
If(StringUpper($UIA_ControlArray[$i][0]) = $tName) Then
Return $UIA_ControlArray[$i][1]
EndIf
Next
Return seterror($_UIASTATUS_GeneralError,$_UIASTATUS_GeneralError,"No control with that name")
EndFunc
Func _UIA_getPropertyIndex($propName)
Local $i
if stringinstr("indexrelative;index;instance", $propName) Then
return seterror(0,0,-1)
EndIf
For $i = 0 To UBound($UIA_propertiesSupportedArray, 1) - 1
If StringLower($UIA_propertiesSupportedArray[$i][0]) = StringLower($propName) Then
Return $i
EndIf
Next
_UIA_LOG("[FATAL] : property you use is having invalid name:=" & $propName & @CRLF, $UIA_Log_Wrapper)
return seterror($_UIASTATUS_GeneralError,$_UIASTATUS_GeneralError,"[FATAL] : property you use is having invalid name:=" & $propName & @CRLF)
EndFunc
Func _UIA_getPropertyValue($obj, $id)
Local $tval
Local $tStr
Local $i
If Not _UIA_IsElement($obj) Then
Return SetError($_UIASTATUS_GeneralError, $_UIASTATUS_GeneralError, "** NO PROPERTYVALUE DUE TO NONEXISTING OBJECT **")
EndIf
$obj.GetCurrentPropertyValue($id, $tval)
$tStr = "" & $tval
If IsArray($tval) Then
$tStr = ""
For $i = 0 To UBound($tval) - 1
$tStr = $tStr & stringstripws($tval[$i],$STR_STRIPLEADING + $STR_STRIPTRAILING )
If $i <> UBound($tval) - 1 Then
$tStr = $tStr & ";"
EndIf
Next
Return $tStr
EndIf
Return SetError($_UIASTATUS_GeneralError, $_UIASTATUS_GeneralError, $tStr)
EndFunc
Func _UIA_getAllPropertyValues($UIA_oUIElement)
Local $tStr, $tval, $tSeparator
$tStr = ""
$tSeparator = @CRLF
For $i = 0 To UBound($UIA_propertiesSupportedArray) - 1
if $UIA_propertiesSupportedArray[$i][1] <> -1 Then
$tval = _UIA_getPropertyValue($UIA_oUIElement, $UIA_propertiesSupportedArray[$i][1])
If $tval <> "" Then
$tStr = $tStr & "UIA_" & $UIA_propertiesSupportedArray[$i][0] & ":= <" & $tval & ">" & $tSeparator
EndIf
endif
Next
Return $tStr
EndFunc
Local $patternArray[21][3] = [ [$UIA_ValuePatternId, $sIID_IUIAutomationValuePattern, $dtagIUIAutomationValuePattern], [$UIA_InvokePatternId, $sIID_IUIAutomationInvokePattern, $dtagIUIAutomationInvokePattern], [$UIA_SelectionPatternId, $sIID_IUIAutomationSelectionPattern, $dtagIUIAutomationSelectionPattern], [$UIA_LegacyIAccessiblePatternId, $sIID_IUIAutomationLegacyIAccessiblePattern, $dtagIUIAutomationLegacyIAccessiblePattern], [$UIA_SelectionItemPatternId, $sIID_IUIAutomationSelectionItemPattern, $dtagIUIAutomationSelectionItemPattern], [$UIA_RangeValuePatternId, $sIID_IUIAutomationRangeValuePattern, $dtagIUIAutomationRangeValuePattern], [$UIA_ScrollPatternId, $sIID_IUIAutomationScrollPattern, $dtagIUIAutomationScrollPattern], [$UIA_GridPatternId, $sIID_IUIAutomationGridPattern, $dtagIUIAutomationGridPattern], [$UIA_GridItemPatternId, $sIID_IUIAutomationGridItemPattern, $dtagIUIAutomationGridItemPattern], [$UIA_MultipleViewPatternId, $sIID_IUIAutomationMultipleViewPattern, $dtagIUIAutomationMultipleViewPattern], [$UIA_WindowPatternId, $sIID_IUIAutomationWindowPattern, $dtagIUIAutomationWindowPattern], [$UIA_DockPatternId, $sIID_IUIAutomationDockPattern, $dtagIUIAutomationDockPattern], [$UIA_TablePatternId, $sIID_IUIAutomationTablePattern, $dtagIUIAutomationTablePattern], [$UIA_TextPatternId, $sIID_IUIAutomationTextPattern, $dtagIUIAutomationTextPattern], [$UIA_TogglePatternId, $sIID_IUIAutomationTogglePattern, $dtagIUIAutomationTogglePattern], [$UIA_TransformPatternId, $sIID_IUIAutomationTransformPattern, $dtagIUIAutomationTransformPattern], [$UIA_ScrollItemPatternId, $sIID_IUIAutomationScrollItemPattern, $dtagIUIAutomationScrollItemPattern], [$UIA_ItemContainerPatternId, $sIID_IUIAutomationItemContainerPattern, $dtagIUIAutomationItemContainerPattern], [$UIA_VirtualizedItemPatternId, $sIID_IUIAutomationVirtualizedItemPattern, $dtagIUIAutomationVirtualizedItemPattern], [$UIA_SynchronizedInputPatternId, $sIID_IUIAutomationSynchronizedInputPattern, $dtagIUIAutomationSynchronizedInputPattern], [$UIA_ExpandCollapsePatternId, $sIID_IUIAutomationExpandCollapsePattern, $dtagIUIAutomationExpandCollapsePattern] ]
Func _UIA_getPattern($obj, $patternID)
Local $pPattern, $oPattern
Local $sIID_Pattern
Local $sdTagPattern
Local $i
For $i = 0 To UBound($patternArray) - 1
If $patternArray[$i][0] = $patternID Then
$sIID_Pattern = $patternArray[$i][1]
$sdTagPattern = $patternArray[$i][2]
exitloop
EndIf
Next
$obj.getCurrentPattern($patternID, $pPattern)
$oPattern = ObjCreateInterface($pPattern, $sIID_Pattern, $sdTagPattern)
If _UIA_IsElement($oPattern) Then
Return $oPattern
Else
_UIA_LOG("UIA WARNING ** NOT ** found the pattern" & @CRLF, $UIA_Log_Wrapper)
return seterror($_UIASTATUS_GeneralError,$_UIASTATUS_GeneralError,"UIA WARNING ** NOT ** found the pattern" & @CRLF)
EndIf
EndFunc
Local Const $cRTI_Prefix="RTI."
Global $__g_hFileLog
_UIA_TFW_Init()
func _UIA_TFW_Init()
OnAutoItExitRegister("_UIA_TFW_Close")
$UIA_Vars = ObjCreate("Scripting.Dictionary")
$UIA_Vars.comparemode = 2
_UIA_LoadConfiguration()
$UIA_Vars.add("DESKTOP", $UIA_oDesktop)
_UIA_VersionInfo()
return seterror($_UIASTATUS_Success,0,0)
EndFunc
func _UIA_TFW_Close()
_UIA_logfileclose()
EndFunc
Func _UIA_LoadConfiguration()
_UIA_setVar("RTI.ACTIONCOUNT", 0)
_UIA_setVar("Global.Debug", True)
_UIA_setVar("Global.Debug.File", True)
_UIA_setVar("Global.Highlight", True)
If FileExists($UIA_CFGFileName) Then
_UIA_loadCFGFile($UIA_CFGFileName)
EndIf
EndFunc
Func _UIA_loadCFGFile($strFname)
Local $sections, $values, $strKey, $strVal, $i, $j
$sections = IniReadSectionNames($strFname)
If @error Then
_UIA_LOG("Error occurred on reading " & $strFname & @CRLF, $UIA_Log_Wrapper)
Else
For $i = 1 To $sections[0]
$values = IniReadSection($strFname, $sections[$i])
If @error Then
_UIA_LOG("Error occurred on reading " & $strFname & @CRLF, $UIA_Log_Wrapper)
Else
For $j = 1 To $values[0][0]
$strKey = $sections[$i] & "." & $values[$j][0]
$strVal = $values[$j][1]
If StringLower($strVal) = "true" Then $strVal = True
If StringLower($strVal) = "false" Then $strVal = False
If StringLower($strVal) = "on" Then $strVal = True
If StringLower($strVal) = "off" Then $strVal = False
If StringLower($strVal) = "minimized" Then $strVal = @SW_MINIMIZE
If StringLower($strVal) = "maximized" Then $strVal = @SW_MAXIMIZE
If StringLower($strVal) = "normal" Then $strVal = @SW_RESTORE
$strVal = StringReplace($strVal, "%windowsdir%", @WindowsDir)
$strVal = StringReplace($strVal, "%programfilesdir%", @ProgramFilesDir)
_UIA_setVar($strKey, $strVal)
Next
EndIf
Next
EndIf
EndFunc
Func _UIA_getVar($varName)
If $UIA_Vars.exists($varName) Then
Local $retVal = $UIA_Vars($varName)
Return $retVal
Else
SetError(1)
Return "*** WARNING: not in repository ***" & $varName
EndIf
EndFunc
Func _UIA_getVars2Array($prefix = "")
Local $keys, $it, $i
_UIA_LOG($UIA_Vars.count - 1 & @CRLF, $UIA_Log_Wrapper)
$keys = $UIA_Vars.keys
$it = $UIA_Vars.items
For $i = 0 To $UIA_Vars.count - 1
Local $oRef = $it[$i]
_UIA_LOG("[" & $keys[$i] & "]:=[" & $oRef & "]" & _UIA_IsElement($oRef) & @CRLF, $UIA_Log_Wrapper)
Next
EndFunc
Func _UIA_setVar($varName, $varValue)
If $UIA_Vars.exists($varName) Then
$UIA_Vars.remove($varName)
$UIA_Vars.add($varName, $varValue)
Else
$UIA_Vars.add($varName, $varValue)
EndIf
EndFunc
func _UIA_normalizeExpression($sPropList)
Local $asAllProperties
Local $iPropertyCount
Local $asProperties2Match[1][4]
local $i
local $aKV
Local $iMatch
local $propName, $propValue
local $bSkip
local $UIA_oUIElement
local $UIA_pUIElement
local $index
$asAllProperties = StringSplit($sPropList, ";", 1)
$iPropertyCount = $asAllProperties[0]+1
ReDim $asProperties2Match[$iPropertyCount][4]
_UIA_LOG("_UIA_normalizeExpression " & $sPropList & ";" & $iPropertyCount & " properties " & @CRLF, $UIA_Log_Wrapper)
For $i = 1 To $iPropertyCount-1
_UIA_LOG("  _UIA_getObjectByFindAll property " & $i & " " & $asAllProperties[$i] & @CRLF, $UIA_Log_Wrapper)
$aKV = StringSplit($asAllProperties[$i], ":=", 1)
$iMatch=0
$bSkip=false
If $aKV[0] = 1 Then
$aKV[1] = StringStripWS($aKV[1], 3)
$propName = $UIA_NamePropertyId
$propValue = $asAllProperties[$i]
Switch $aKV[1]
Case "active", "[active]"
$UIA_oUIAutomation.GetFocusedElement($UIA_pUIElement)
$UIA_oUIElement = ObjCreateInterface($UIA_pUIElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
$propName = "object"
$propValue=$UIA_oUIElement
$iMatch = 1
$bSkip=true
Case "last", "[last]"
$UIA_oUIElement = _UIA_getVar("RTI.LASTELEMENT")
If Not _UIA_IsElement($UIA_oUIElement) Then
$UIA_oUIElement = $UIA_oDesktop
EndIf
$propName = "object"
$propValue=$UIA_oUIElement
$iMatch = 1
$bSkip=true
Case Else
$propName = $UIA_NamePropertyId
$propValue = $asAllProperties[$i]
$iMatch=0
$bSkip=false
EndSwitch
$asProperties2Match[$i][0] = $propName
$asProperties2Match[$i][1] = $propValue
$asProperties2Match[$i][2] = $iMatch
$asProperties2Match[$i][3] = $bSkip
Else
$aKV[1] = StringStripWS($aKV[1], 3)
$aKV[2] = StringStripWS($aKV[2], 3)
$propName = $aKV[1]
$propValue = $aKV[2]
$iMatch=0
$bSkip=false
$index = _UIA_getPropertyIndex($propName)
if $index >=0 Then
Switch $UIA_propertiesSupportedArray[$index][1]
Case $UIA_ControlTypePropertyId
$propValue = Number(_UIA_getControlID($propValue))
EndSwitch
_UIA_LOG( " name:[" & $propName & "] value:[" & $propValue & "] having index " & $index & @CRLF, $UIA_Log_Wrapper)
$asProperties2Match[$i][0] = $UIA_propertiesSupportedArray[$index][1]
Else
$asProperties2Match[$i][0] = $propName
EndIf
$asProperties2Match[$i][1] = $propValue
$asProperties2Match[$i][2] = $iMatch
$asProperties2Match[$i][3] = $bSkip
EndIf
Next
$asProperties2Match[0][0]=$iPropertyCount
return $asProperties2Match
EndFunc
Func _UIA_getObjectByFindAll($obj, $str, $treeScope=$treescope_children, $p1 = 0)
Local $pElements, $iLength
Local $propertyID
Local $relPos = 0
Local $relIndex = 0
Local $iMatch = 0
Local $parentHandle
Local $properties2Match[1][2]
Local $allProperties, $propertyCount, $propName, $propValue, $bAdd, $index, $i, $arrSize, $j
Local $objParent, $propertyActualValue, $propertyVal, $oAutomationElementArray, $matchCount
local $bSkip=false
Local $tXMLLogString
$properties2Match=_UIA_normalizeExpression($str)
if $properties2Match[1][0] = "object" Then
$UIA_oUIElement = $properties2Match[1][1]
$iMatch=1
EndIf
If $iMatch = 0 Then
$arrSize = UBound($properties2Match, 1) - 1
For $j = 1 To $arrSize
$bSkip=$properties2Match[$j][3]
if $bSkip=true then
If $propName = "indexrelative" Then
$relPos = $propValue
EndIf
If($propName = "index") Or($propName = "instance") Then
$relIndex = $propValue
EndIf
EndIf
Next
If _UIA_IsElement($obj) Then
_UIA_LOG("*** Try to get a list of elements *** " & $treeScope & @CRLF, $UIA_Log_Wrapper)
$obj.FindAll($treeScope, $UIA_oTRUECondition, $pElements)
$oAutomationElementArray = ObjCreateInterface($pElements, $sIID_IUIAutomationElementArray, $dtagIUIAutomationElementArray)
EndIf
$matchCount = 0
If _UIA_IsElement($oAutomationElementArray) Then
$oAutomationElementArray.Length($iLength)
Else
_UIA_LOG("***** FATAL:???? _UIA_getObjectByFindAll no childarray found for object with following details *****" & @CRLF, $UIA_Log_Wrapper)
_UIA_LOG(_UIA_getAllPropertyValues($obj) & @CRLF, $UIA_Log_Wrapper)
$iLength = 0
EndIf
$tXMLLogString=""
$tXMLLogString = $tXMLLogString & "<propertymatching>"
For $i = 0 To $iLength - 1
$oAutomationElementArray.GetElement($i, $UIA_pUIElement)
$UIA_oUIElement = ObjCreateInterface($UIA_pUIElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
For $j = 1 To $arrSize
$bSkip=$properties2Match[$j][3]
if $bSkip=false Then
$propertyID = $properties2Match[$j][0]
$propertyVal = $properties2Match[$j][1]
$propertyActualValue = ""
Switch $propertyID
Case $UIA_ControlTypePropertyId
$propertyVal = Number($propertyVal)
EndSwitch
$propertyActualValue = _UIA_getPropertyValue($UIA_oUIElement, $propertyID)
$iMatch = StringRegExp($propertyActualValue, $propertyVal, $STR_REGEXPMATCH )
$tXMLLogString=$tXMLLogString & _UIA_EncodeHTMLString("        j:" & $j & " propID:[" & $propertyID & "] expValue:[" & $propertyVal & "]actualValue:[" & $propertyActualValue & "]" & $iMatch & @CRLF)
If $iMatch = 0 Then
If $propertyVal <> $propertyActualValue Then
ExitLoop
EndIf
EndIf
EndIf
Next
If $iMatch = 1 Then
If $relPos <> 0 Then
$oAutomationElementArray.GetElement($i + $relPos, $UIA_pUIElement)
$UIA_oUIElement = ObjCreateInterface($UIA_pUIElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
EndIf
If $relIndex <> 0 Then
$matchCount = $matchCount + 1
If $matchCount <> $relIndex Then $iMatch = 0
_UIA_LOG("Index position used " & $relIndex & $matchCount, $UIA_Log_Wrapper)
EndIf
If $iMatch = 1 Then
ExitLoop
EndIf
EndIf
Next
$tXMLLogString = $tXMLLogString & "</propertymatching>"
_UIA_LOG($tXMLLogString, $UIA_Log_Wrapper)
EndIf
If $iMatch = 1 Then
$UIA_oTW.getparentelement($UIA_oUIElement, $parentHandle)
$objParent = ObjCreateInterface($parentHandle, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
If _UIA_IsElement($objParent) = 0 Then
_UIA_LOG("No parent " & @CRLF, $UIA_Log_Wrapper)
Else
_UIA_LOG("Storing parent for found object in RTI as RTI.PARENT " & _UIA_getPropertyValue($objParent,$UIA_NamePropertyId) & @CRLF, $UIA_Log_Wrapper)
_UIA_setVar("RTI.PARENT", $objParent)
EndIf
If IsString($p1) Then
_UIA_LOG("Storing in RTI as RTI." & $p1 & @CRLF, $UIA_Log_Wrapper)
_UIA_setVar("RTI." & StringUpper($p1), $UIA_oUIElement)
EndIf
If _UIA_getVar("Global.Highlight") = True Then
_UIA_Highlight($UIA_oUIElement)
EndIf
Return $UIA_oUIElement
Else
Return ""
EndIf
EndFunc
func _UIA_getObject($obj_or_string)
Local $oElement
Local $tPhysical
Local $startElement, $oStart, $pos, $tStr
Local $xx
local $oParentHandle, $oParentBefore, $i, $parentCount
if _UIA_IsElement($obj_or_string) Then
$oElement=$obj_or_string
else
$oElement = _UIA_getVar($cRTI_Prefix & $obj_or_string)
If @error = 1 Then
$oElement = _UIA_getVar($obj_or_string)
$tPhysical = $oElement
EndIf
If @error = 1 Then
_UIA_LOG("Finding object (bypassing repository) with physical description " & $tPhysical & ":" & $obj_or_string & @CRLF, $UIA_Log_Wrapper)
$tPhysical = $obj_or_string
EndIf
EndIf
If _UIA_IsElement($oElement) Then
_UIA_LOG("Quickly referenced object " & $tPhysical & ":" & $obj_or_string & @CRLF, $UIA_Log_Wrapper)
Else
If Stringright($obj_or_string, stringlen(".mainwindow"))= ".mainwindow" Then
$startElement = "Desktop"
$oStart = $UIA_oDesktop
_UIA_LOG("Fallback finding1 object under " & $startElement & @TAB & $tPhysical & @CRLF, $UIA_Log_Wrapper)
$oElement = _UIA_getObjectByFindAll($oStart, $tPhysical, $treescope_children, $obj_or_string)
_UIA_setVar("RTI.MAINWINDOW", $oElement)
_UIA_setVar($cRTI_Prefix & stringupper($obj_or_string),$oElement)
_UIA_setVar("RTI.SEARCHCONTEXT", $oElement)
Else
$oStart = _UIA_getVar("RTI.SEARCHCONTEXT")
$startElement = "RTI.SEARCHCONTEXT"
If Not _UIA_IsElement($oStart) Then
$pos = StringInStr($obj_or_string, ".")
_UIA_LOG("_UIA_action: No RTI.SEARCHCONTEXT used for " & $obj_or_string & @CRLF, $UIA_Log_Wrapper)
If $pos > 0 Then
$tStr = $cRTI_Prefix & StringLeft(StringUpper($obj_or_string), $pos - 1) & ".MAINWINDOW"
Else
$tStr = "RTI.MAINWINDOW"
EndIf
_UIA_LOG("_UIA_action: try for " & $tStr & @CRLF, $UIA_Log_Wrapper)
$oStart = _UIA_getVar($tStr)
$startElement = $tStr
If Not _UIA_IsElement($oStart) Then
_UIA_LOG("_UIA_action: No RTI.MAINWINDOW used for " & $obj_or_string & @CRLF, $UIA_Log_Wrapper)
$xx = _UIA_getVars2Array()
$oStart = _UIA_getVar("RTI.PARENT")
$startElement = "RTI.PARENT"
If Not _UIA_IsElement($oStart) Then
_UIA_LOG("_UIA_action: No RTI.PARENT used for " & $obj_or_string & @CRLF, $UIA_Log_Wrapper)
$startElement = "Desktop"
$oStart = $UIA_oDesktop
EndIf
EndIf
EndIf
_UIA_LOG("_UIA_action: Finding object " & $obj_or_string & " object a:=" & _UIA_IsElement($obj_or_string) & " under " & $startElement & " object b:=" & _UIA_IsElement($oStart) & @CRLF, $UIA_Log_Wrapper)
_UIA_LOG("  looking for " & $tPhysical & @CRLF, $UIA_Log_Wrapper)
$oElement = _UIA_getObjectByFindAll($oStart, $tPhysical, $treescope_children)
if not _UIA_IsElement($oElement) Then
_UIA_LOG("  deep find in subtree " & $tPhysical & @CRLF, $UIA_Log_Wrapper)
$oElement = _UIA_getObjectByFindAll($oStart, $tPhysical, $treescope_subtree)
EndIf
if not _UIA_IsElement($oElement) Then
_UIA_LOG("  walking back to mainwindow and deep find in subtree " & $tPhysical & @CRLF, $UIA_Log_Wrapper)
$UIA_oUIAutomation.RawViewWalker($UIA_pTW)
$UIA_oTW=ObjCreateInterface($UIA_pTW, $sIID_IUIAutomationTreeWalker, $dtagIUIAutomationTreeWalker)
If not _UIA_IsElement($UIA_oTW) Then
_UIA_LOG("UI automation treewalker failed. UI Automation failed failed " & @CRLF, $UIA_Log_Wrapper)
EndIf
$i=0
$UIA_oTW.getparentelement($oStart,$oParentHandle)
$oParentHandle=objcreateinterface($oparentHandle,$sIID_IUIAutomationElement, $dtagIUIAutomationElement)
If _UIA_IsElement($oParentHandle) = 0 Then
_UIA_LOG("No parent: UI Automation failed could be you spy the desktop", $UIA_Log_Wrapper)
Else
while($i <=$_UIA_MAXDEPTH) and(_UIA_IsElement($oParentHandle)=true)
$i=$i+1
$oParentBefore=$oParentHandle
$UIA_oTW.getparentelement($oparentHandle,$oparentHandle)
$oParentHandle=objcreateinterface($oparentHandle,$sIID_IUIAutomationElement, $dtagIUIAutomationElement)
wend
$parentCount=$i-1
$oStart=$oParentBefore
EndIf
$oElement = _UIA_getObjectByFindAll($oStart, $tPhysical, $treescope_subtree)
EndIf
if not _UIA_IsElement($oElement) Then
If _UIA_getVar("Global.Debug") = True Then
_UIA_DumpThemAll($oStart, $treescope_subtree)
EndIf
endif
EndIf
EndIf
return $oElement
EndFunc
Func _UIA_action($obj_or_string, $strAction, $p1 = 0, $p2 = 0, $p3 = 0, $p4 = 0)
Local $obj2ActOn
Local $tPattern
Local $x, $y
Local $controlType
Local $oElement
Local $hwnd
Local $retVal=True
Local $tRect
_UIA_LOG($__UIA_debugCacheOn)
_UIA_LOG("<action>")
$oElement=_UIA_getObject($obj_or_string)
If _UIA_IsElement($oElement) Then
$obj2ActOn = $oElement
_UIA_setVar("RTI.LASTELEMENT", $oElement)
$controlType = _UIA_getPropertyValue($obj2ActOn, $UIA_ControlTypePropertyId)
Else
If Not StringInStr("exist,exists", $strAction) Then
_UIA_LOG("Not an object failing action " & $strAction & " on " & $obj_or_string & @CRLF, $UIA_Log_Wrapper)
SetError(1)
Return False
EndIf
EndIf
_UIA_setVar("RTI.ACTIONCOUNT", _UIA_getVar("RTI.ACTIONCOUNT") + 1)
_UIA_LOG("Action " & _UIA_getVar("RTI.ACTIONCOUNT") & " " & $strAction & " on " & $obj_or_string & " _UIA_IsElement:=" & _UIA_IsElement($obj2ActOn) & " par 1:=" & $p1 & " par 2:=" & $p2 & @CRLF, $UIA_Log_Wrapper)
_UIA_LOG("</action>")
_UIA_LOG($__UIA_debugCacheOff)
Switch $strAction
Case "leftclick", "left", "click", "leftdoubleclick", "leftdouble", "doubleclick", "rightclick", "right", "rightdoubleclick", "rightdouble", "middleclick", "middle", "middledoubleclick", "middledouble", "mousemove", "movemouse"
Local $clickAction = "left"
Local $clickCount = 1
If StringInStr($strAction, "right") Then $clickAction = "right"
If StringInStr($strAction, "middle") Then $clickAction = "middle"
If StringInStr($strAction, "double") Then $clickCount = 2
Local $t
$t = StringSplit(_UIA_getPropertyValue($obj2ActOn, $UIA_BoundingRectanglePropertyId), ";")
$x = Int($t[1] +($t[3] / 2))
$y = Int($t[2] + $t[4] / 2)
MouseMove($x, $y, 0)
if not stringinstr($strAction,"move") Then
MouseClick($clickAction, $x, $y, $clickCount, 0)
EndIf
Sleep($UIA_DefaultWaitTime)
Case "setvalue","settextvalue"
If($controlType = $UIA_WindowControlTypeId) Then
$hwnd = 0
$obj2ActOn.CurrentNativeWindowHandle($hwnd)
WinSetTitle(HWnd($hwnd), "", $p1)
Else
$obj2ActOn.setfocus()
Sleep($UIA_DefaultWaitTime)
$tPattern = _UIA_getPattern($obj2ActOn, $UIA_LegacyIAccessiblePatternId)
if _UIA_IsElement($tPattern) Then
$tPattern.setvalue($p1)
Else
$tPattern = _UIA_getPattern($obj2ActOn, $UIA_ValuePatternId)
if _UIA_IsElement($tPattern) Then
$tPattern.setvalue($p1)
EndIf
EndIf
EndIf
Case "setvalue using keys"
$obj2ActOn.setfocus()
Send("^a")
Send($p1)
Sleep($UIA_DefaultWaitTime)
case "getvalue"
$obj2ActOn.setfocus()
Send("^a")
Send("^c")
$retVal=clipget()
Case "sendkeys", "enterstring", "type", "typetext"
$obj2ActOn.setfocus()
Send($p1)
Case "invoke"
$obj2ActOn.setfocus()
Sleep($UIA_DefaultWaitTime)
$tPattern = _UIA_getPattern($obj2ActOn, $UIA_InvokePatternId)
$tPattern.invoke()
Case "focus", "setfocus", "activate"
_UIA_setVar("RTI.SEARCHCONTEXT", $obj2Acton)
$obj2ActOn.setfocus()
Sleep($UIA_DefaultWaitTime)
Case "close"
$tPattern = _UIA_getPattern($obj2ActOn, $UIA_WindowPatternId)
$tPattern.close()
Case "move","setposition"
$tPattern = _UIA_getPattern($obj2ActOn, $UIA_TransformPatternId)
$tPattern.move($p1, $p2)
Case "resize"
$tPattern = _UIA_getPattern($obj2ActOn, $UIA_WindowPatternId)
$tPattern.SetWindowVisualState($WindowVisualState_Normal)
$tPattern = _UIA_getPattern($obj2ActOn, $UIA_TransformPatternId)
$tPattern.resize($p1, $p2)
Case "minimize"
$tPattern = _UIA_getPattern($obj2ActOn, $UIA_WindowPatternId)
$tPattern.SetWindowVisualState($WindowVisualState_Minimized)
Case "maximize"
$tPattern = _UIA_getPattern($obj2ActOn, $UIA_WindowPatternId)
$tPattern.SetWindowVisualState($WindowVisualState_Maximized)
Case "normal"
$tPattern = _UIA_getPattern($obj2ActOn, $UIA_WindowPatternId)
$tPattern.SetWindowVisualState($WindowVisualState_Normal)
Case "close"
$tPattern = _UIA_getPattern($obj2ActOn, $UIA_WindowPatternId)
$tPattern.close()
Case "exist", "exists"
$retVal=true
Case "searchcontext", "context"
_UIA_setVar("RTI.SEARCHCONTEXT", $obj2ActOn)
Case "highlight"
_UIA_Highlight($obj2ActOn)
Case "getobject", "object"
Return $obj2ActOn
Case "attach"
Return $obj2ActOn
case "capture","screenshot","takescreenshot"
$tRect = StringSplit(_UIA_getPropertyValue($obj2ActOn, $UIA_BoundingRectanglePropertyId), ";")
consolewrite($p1 & ";" & $tRect[1] & ";" &($tRect[3] + $tRect[1]) & ";" & $tRect[2]& ";" &($tRect[4] + $tRect[2]))
_ScreenCapture_Capture($p1, $tRect[1], $tRect[2], $tRect[3] + $tRect[1], $tRect[4] + $tRect[2])
case "dump", "dumpthemall"
_UIA_DumpThemAll($obj2ActOn,$treescope_subtree)
Case Else
EndSwitch
Return $retVal
EndFunc
Func _UIA_DumpThemAll($oElementStart, $treeScope)
Local $pCondition, $pTrueCondition, $oCondition, $oAutomationElementArray
Local $pElements, $iLength, $i
local $dumpStr
local $tStr
$dumpStr="<treedump>"
$dumpStr=$dumpStr & "<treeheader>***** Dumping tree *****</treeheader>"
$oElementStart.FindAll($treeScope, $UIA_oTRUECondition, $pElements)
$oAutomationElementArray = ObjCreateInterface($pElements, $sIID_IUIAutomationElementArray, $dtagIUIAutomationElementArray)
If _UIA_IsElement($oAutomationElementArray) Then
$oAutomationElementArray.Length($iLength)
Else
$dumpStr=$dumpStr & "<fatal>***** FATAL:???? _UIA_DumpThemAll no childarray found ***** </fatal>"
$iLength = 0
EndIf
For $i = 0 To $iLength - 1
$oAutomationElementArray.GetElement($i, $UIA_pUIElement)
$UIA_oUIElement = ObjCreateInterface($UIA_pUIElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
$tStr="Title is: <" & _UIA_getPropertyValue($UIA_oUIElement, $UIA_NamePropertyId) & ">" & @TAB & "Class   := <" & _UIA_getPropertyValue($UIA_oUIElement, $UIA_ClassNamePropertyId) & ">" & @TAB & "controltype:= " & "<" & _UIA_getControlName(_UIA_getPropertyValue($UIA_oUIElement, $UIA_ControlTypePropertyId)) & ">" & @TAB & ",<" & _UIA_getPropertyValue($UIA_oUIElement, $UIA_ControlTypePropertyId) & ">" & @TAB & ", (" & Hex(_UIA_getPropertyValue($UIA_oUIElement, $UIA_ControlTypePropertyId)) & ")" & @TAB & ", acceleratorkey:= <" & _UIA_getPropertyValue($UIA_oUIElement, $UIA_AcceleratorKeyPropertyId) & ">" & @TAB & ", automationid:= <" & _UIA_getPropertyValue($UIA_oUIElement, $UIA_AutomationIdPropertyId) & ">" & @TAB
$tstr=_UIA_EncodeHTMLString($tStr)
$dumpStr=$dumpStr & "<elementinfo>" & $tStr & "</elementinfo>"
Next
$dumpStr=$dumpStr & "</treedump>"
_UIA_LOG($dumpStr)
EndFunc
Func _UIA_Highlight($oElement)
Local $t
$t = StringSplit(_UIA_getPropertyValue($oElement, $UIA_BoundingRectanglePropertyId), ";")
_UIA_DrawRect($t[1], $t[3] + $t[1], $t[2], $t[4] + $t[2])
EndFunc
func _UIA_EncodeHTMLString($str)
Local $tStr = $str
$tStr = StringReplace($tStr, "&", "&amp;")
$tStr = StringReplace($tStr, ">", "&gt;")
$tStr = StringReplace($tStr, "<", "&lt;")
$tStr = StringReplace($tStr, """", "&quot;")
$tStr = StringReplace($tStr, "'", "&apos;")
Return $tStr
EndFunc
func _UIA_LogFileClose()
filewrite($__g_hFileLog,"</log>" & @CRLF)
fileclose($__g_hFileLog)
EndFunc
Func _UIA_LOG($s, $logLevel = 0)
Local $logstr
local $bFlushCache=false
if $s=$__UIA_debugCacheOn Then
$s=""
$__l_UIA_CacheState=True
$bFlushCache=true
endif
if $s=$__UIA_debugCacheOff Then
$__l_UIA_CacheState=False
$s=$__gl_XMLCache
EndIf
if stringleft($s,1)<>"<" Then
$s=_UIA_EncodeHTMLString($s)
EndIf
if $__l_UIA_CacheState=True and $bFlushCache=false Then
$__gl_XMLCache=$__gl_XMLCache & $s
EndIf
if($__l_UIA_CacheState=false) or($bFlushCache=true) Then
if stringright($s,2)=@CRLF Then
$s=stringleft($s,stringlen($s)-2)
EndIf
$logstr= "<logline level=""" & $logLevel & """"
$logstr = $logstr & " timestamp=""" & @YEAR & @MON & @MDAY & "-" & @HOUR & @MIN & @SEC & @MSEC & """>"
$logstr = $logstr & " " & $s & "</logline>" & @CRLF
If _UIA_getVar("global.debug.file") = True Then
FileWrite($__g_hFileLog, $logStr)
Else
If _UIA_getVar("global.debug") = True Then
ConsoleWrite($logstr)
EndIf
EndIf
$__gl_XMLCache=""
EndIf
return $_UIASTATUS_Success
EndFunc
Func _UIA_DrawRect($tLeft, $tRight, $tTop, $tBottom, $color = 0xFF, $PenWidth = 4)
Local $hDC, $hPen, $obj_orig, $x1, $x2, $y1, $y2
$x1 = $tLeft
$x2 = $tRight
$y1 = $tTop
$y2 = $tBottom
$hDC = _WinAPI_GetWindowDC(0)
$hPen = _WinAPI_CreatePen($PS_SOLID, $PenWidth, $color)
$obj_orig = _WinAPI_SelectObject($hDC, $hPen)
_WinAPI_DrawLine($hDC, $x1, $y1, $x2, $y1)
_WinAPI_DrawLine($hDC, $x2, $y1, $x2, $y2)
_WinAPI_DrawLine($hDC, $x2, $y2, $x1, $y2)
_WinAPI_DrawLine($hDC, $x1, $y2, $x1, $y1)
_WinAPI_SelectObject($hDC, $obj_orig)
_WinAPI_DeleteObject($hPen)
_WinAPI_ReleaseDC(0, $hDC)
EndFunc
Func _UIA_IsElement($control)
Return IsObj($control)
EndFunc
Func _UIA_VersionInfo()
_UIA_LOG("<information> Information" & "_UIA_VersionInfo" & "version " & $__gaUIAAU3VersionInfo[0] & $__gaUIAAU3VersionInfo[1] & "." & $__gaUIAAU3VersionInfo[2] & "-" & $__gaUIAAU3VersionInfo[3] & "Release date: " & $__gaUIAAU3VersionInfo[4] & "</information>" & @CRLF,$uia_log_wrapper)
Return SetError($_UIASTATUS_Success, 0, $__gaUIAAU3VersionInfo)
EndFunc
Func _StringExplode($sString, $sDelimiter, $iLimit = 0)
If $iLimit = Default Then $iLimit = 0
If $iLimit > 0 Then
Local Const $NULL = Chr(0)
$sString = StringReplace($sString, $sDelimiter, $NULL, $iLimit)
$sDelimiter = $NULL
ElseIf $iLimit < 0 Then
Local $iIndex = StringInStr($sString, $sDelimiter, 0, $iLimit)
If $iIndex Then
$sString = StringLeft($sString, $iIndex - 1)
EndIf
EndIf
Return StringSplit($sString, $sDelimiter, $STR_ENTIRESPLIT + $STR_NOCOUNT)
EndFunc
Global Const $DMW_SHORTNAME = 1
Global Const $DMW_LOCALE_LONGNAME = 2
Global Const $LOCALE_INVARIANT = 0x007F
Global Const $LOCALE_USER_DEFAULT = 0x0400
Func _WinAPI_GetDateFormat($iLCID = 0, $tSYSTEMTIME = 0, $iFlags = 0, $sFormat = '')
If Not $iLCID Then $iLCID = 0x0400
Local $sTypeOfFormat = 'wstr'
If Not StringStripWS($sFormat, $STR_STRIPLEADING + $STR_STRIPTRAILING) Then
$sTypeOfFormat = 'ptr'
$sFormat = 0
EndIf
Local $aRet = DllCall('kernel32.dll', 'int', 'GetDateFormatW', 'dword', $iLCID, 'dword', $iFlags, 'struct*', $tSYSTEMTIME, $sTypeOfFormat, $sFormat, 'wstr', '', 'int', 2048)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, '')
Return $aRet[5]
EndFunc
Func _DateDayOfWeek($iDayNum, $iFormat = Default)
Local Const $MONDAY_IS_NO1 = 128
If $iFormat = Default Then $iFormat = 0
$iDayNum = Int($iDayNum)
If $iDayNum < 1 Or $iDayNum > 7 Then Return SetError(1, 0, "")
Local $tSYSTEMTIME = DllStructCreate($tagSYSTEMTIME)
DllStructSetData($tSYSTEMTIME, "Year", BitAND($iFormat, $MONDAY_IS_NO1) ? 2007 : 2006)
DllStructSetData($tSYSTEMTIME, "Month", 1)
DllStructSetData($tSYSTEMTIME, "Day", $iDayNum)
Return _WinAPI_GetDateFormat(BitAND($iFormat, $DMW_LOCALE_LONGNAME) ? $LOCALE_USER_DEFAULT : $LOCALE_INVARIANT, $tSYSTEMTIME, 0, BitAND($iFormat, $DMW_SHORTNAME) ? "ddd" : "dddd")
EndFunc
AutoItSetOption("WinTitleMatchMode", 2)
AutoItSetOption("MustDeclareVars", 1)
Global $sendcode = False
Global $code = ""
Global $mode = ""
Global $setmode = False
Global $nodefault = False
Global $setupCompleted = False
Global $CodeCount = 0
Global $comp = @ComputerName
Global $hour = @HOUR
Global $day = _DateDayOfWeek(@WDAY)
Global $errorCount = 0
Const $F1 = "C:\Fellowship One Check-in 2.6\AppStart.exe"
Const $F1Updater = "AppStart.exe"
Const $F1Win7 = "C:\FT\Fellowship One Check-in 2.5\AppStart.exe"
Const $process = "FellowshipTech.Application.Windows.CheckIn.exe"
Func CheckWindow()
Local $han = WinGetHandle("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church")
For $count = 0 To 5 Step +1
If WinExists("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church") Then
If WinActivate($han, "") Then
ExitLoop
Else
WinActivate($han, "")
ErrorSearch()
EndIf
Else
WinWaitActive("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church", 5)
EndIf
If $count = 5 Then
MsgBox($MB_SYSTEMMODAL, "", "Problem with Check-in" & @CRLF & "No Network or Slow Internet!" & @CRLF & "Contact IT department please", 10)
ExitLoop
EndIf
Next
EndFunc
Func Checkprocess()
Local $han = WinGetHandle("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church")
If Not ProcessExists($process) And Not ProcessExists($F1Updater) Then
If @OSVersion = "WIN_7" Then
Run($F1Win7)
Else
Run($F1)
EndIf
Else
ProcessClose($process)
ProcessClose($F1Updater)
Sleep(1000)
If @OSVersion = "WIN_7" Then
Run($F1Win7)
Else
Run($F1)
Sleep(1000)
WinWaitActive("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church", 5)
Local $han = WinGetHandle("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church")
WinActivate($han, "")
EndIf
EndIf
EndFunc
Func SendCode($activity)
Sleep(1000)
Local $oP0 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
_UIA_Action($oP0, "setfocus")
_UIA_action(".mainwindow", "sendkeys", $activity)
EndFunc
Func SelfClick()
Local $oP1 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=SwitchForm;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
Local $oP0 = _UIA_getObjectByFindAll($oP1, "Title:=Self  Check-in;controltype:=UIA_CheckBoxControlTypeId;class:=WindowsForms10.BUTTON.app.0.*", $treescope_children)
_UIA_action($oP0, "leftclick")
Local $oP1 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=SwitchForm;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
Local $oP0 = _UIA_getObjectByFindAll($oP1, "Name:=Next;controltype:=UIA_ButtonControlTypeId;class:=WindowsForms10.BUTTON.app.0.*", $treescope_children)
_UIA_action($oP0, "leftclick")
EndFunc
Func AssistClick()
Local $oP1 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=SwitchForm;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
Local $oP0 = _UIA_getObjectByFindAll($oP1, "Title:=Assisted Check-in;controltype:=UIA_CheckBoxControlTypeId;class:=WindowsForms10.BUTTON.app.0.*", $treescope_children)
_UIA_Action($oP0, "leftclick")
Local $oP2 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=SwitchForm;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
Local $oP1 = _UIA_getObjectByFindAll($oP2, "Title:=What type of check-in would you like?;controltype:=UIA_PaneControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
Local $oP0 = _UIA_getObjectByFindAll($oP1, "Title:=Enable Rapid;controltype:=UIA_CheckBoxControlTypeId;class:=WindowsForms10.BUTTON.app.0.*", $treescope_children)
_UIA_Action($oP0, "leftclick")
Local $oP1 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=SwitchForm;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
Local $oP0 = _UIA_getObjectByFindAll($oP1, "Name:=Next;controltype:=UIA_ButtonControlTypeId;class:=WindowsForms10.BUTTON.app.0.*", $treescope_children)
_UIA_action($oP0, "leftclick")
EndFunc
Func DefaultActivityClick()
Local $oP2 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=DefaultActivityForm;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
Local $oP1 = _UIA_getObjectByFindAll($oP2, "Title:=Select default activity and time for checkin.;controltype:=UIA_PaneControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
Local $oP0 = _UIA_getObjectByFindAll($oP1, "Title:=Best Fit;controltype:=UIA_CheckBoxControlTypeId;class:=WindowsForms10.BUTTON.app.0.*", $treescope_children)
_UIA_Action($oP0, "leftclick")
Local $oP1 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=DefaultActivityForm;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
Local $oP0 = _UIA_getObjectByFindAll($oP1, "Name:=Start Check-in;controltype:=UIA_ButtonControlTypeId;class:=WindowsForms10.BUTTON.app.0.*", $treescope_children)
_UIA_action($oP0, "leftclick")
EndFunc
Func WhichWindow()
Local $han = WinGetHandle("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church")
Local $text = WinGetText($han, "")
Local $tArray = _StringExplode($text, @LF)
Local $iIndex = _ArraySearch($tArray, "Activity Code:", Default)
If @error Then
Local $iIndex = _ArraySearch($tArray, "Next", Default)
If @error Then
Local $iIndex = _ArraySearch($tArray, "No Default", Default)
If @error Then
Local $iIndex = _ArraySearch($tArray, "Scan your card to check-in", Default)
If @error Then
Local $iIndex = _ArraySearch($tArray, "Return All Household Members", Default)
If @error Then
Local $iIndex = _ArraySearch($tArray, "Application Setup", Default)
If @error Then
ErrorSearch()
MsgBox($MB_SYSTEMMODAL, "", "I Cant find screen Please contact IT department and let them know.", 15)
Else
$setupCompleted = True
EndIf
Else
$setupCompleted = True
EndIf
Else
$setupCompleted = True
EndIf
Else
$nodefault = True
$sendcode = False
$setmode = False
EndIf
Else
$setmode = True
$sendcode = False
$nodefault = False
EndIf
Else
$sendcode = True
$setmode = False
$nodefault = False
EndIf
EndFunc
Func ErrorSearch()
Local $han = WinGetHandle("Error", "remote")
Local $text = WinGetText($han, "")
Local $tArray = _StringExplode($text, @LF)
Local $iIndex = _ArraySearch($tArray, "OK", Default)
If @error Then
Local $han = WinGetHandle("Connection Error", "OK")
Local $text = WinGetText($han, "")
Local $tArray = _StringExplode($text, @LF)
Local $iIndex = _ArraySearch($tArray, "OK", Default)
If @error Then
Local $han = WinGetHandle("ApplicationUpdater", "&Close program")
Local $text = WinGetText($han, "")
Local $tArray = _StringExplode($text, @LF)
Local $iIndex = _ArraySearch($tArray, "&Close program", Default)
If @error Then
Else
MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One Application Updater, Please contact IT" & @CRLF & "Trying Again", 15)
$errorCount = $errorCount + 1
Checkprocess()
If $errorCount > 3 Then
MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One, Please contact IT", 15)
Exit
EndIf
EndIf
Else
MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One, Please contact IT" & @CRLF & "Trying Again", 15)
$errorCount = $errorCount + 1
Checkprocess()
If $errorCount > 3 Then
MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One, Please contact IT", 15)
Exit
EndIf
EndIf
Else
MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One, Please contact IT" & @CRLF & "Trying Again", 15)
$errorCount = $errorCount + 1
Checkprocess()
If $errorCount > 3 Then
MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One, Please contact IT", 15)
Exit
EndIf
EndIf
EndFunc
Func Setup()
If $code = "" Or $mode = "" Then
MsgBox($MB_SYSTEMMODAL, "Error", "No settings for Service. Contact IT or enter the correct code.", 5)
Exit
Else
While $setupCompleted = False
ErrorSearch()
WhichWindow()
If $sendcode = True Then
$CodeCount = $CodeCount + 1
If $CodeCount < 5 Then
SendCode($code)
$sendcode = False
Sleep(500)
Else
MsgBox($MB_SYSTEMMODAL, "", "Invalid Code. Please Contact Jonathan or Nate on chanel 7 and let them Know", 15)
$setupCompleted = True
EndIf
ElseIf $setmode = True Then
Select
Case $mode = "Self"
SelfClick()
Case $mode = "Assisted"
AssistClick()
DefaultActivityClick()
EndSelect
$setmode = False
ElseIf $nodefault = True Then
DefaultActivityClick()
$nodefault = False
$setupCompleted = False
ElseIf $setupCompleted = True Then
ExitLoop
Else
MsgBox($MB_SYSTEMMODAL, "", "Problem with Setup, please contact Jonathan or Nate.", 15)
ExitLoop
EndIf
WEnd
EndIf
EndFunc
Func GetSettings($hour, $comp, $day)
Local $cNode = ""
Local $oXML = ObjCreate("Microsoft.XMLDOM")
$oXML.load(@ScriptDir & "\eSch.xml")
Local $nodeSpecial = $oXML.selectNodes("//Service")
For $special In $nodeSpecial
Local $nodeProperty = $special.getAttribute("Special")
If $nodeProperty = "Yes" Then
$cNode = $special
Local $nodeComputers = $cNode.childNodes
For $nodeComputer In $nodeComputers
Local $nodeProperty = $nodeComputer.getAttribute("Station")
If @error Then
ElseIf $nodeProperty = $comp Then
$cNode = $nodeComputer
Local $nodeSettings = $cNode.childNodes
For $nodeSetting In $nodeSettings
If $nodeSetting.baseName = "Code" Then
$code = $nodeSetting.text
ElseIf $nodeSetting.baseName = "Mode" Then
$mode = $nodeSetting.text
Else
EndIf
Next
EndIf
Next
Else
EndIf
Next
Local $nodeDays = $oXML.selectNodes("//Service")
For $nodeDay In $nodeDays
Local $nodeProperty = $nodeDay.getAttribute("Day")
If $nodeProperty = $day Then
$cNode = $nodeDay
Local $nodeHours = $cNode.childNodes
For $nodeHour In $nodeHours
Local $nodeProperty = $nodeHour.getAttribute("Hour")
If $nodeProperty = $hour Then
$cNode = $nodeHour
Local $nodeComputers = $cNode.childNodes
For $nodeComputer In $nodeComputers
Local $nodeProperty = $nodeComputer.getAttribute("Station")
If $nodeProperty = $comp Then
$cNode = $nodeComputer
Local $nodeSettings = $cNode.childNodes
For $nodeSetting In $nodeSettings
If $nodeSetting.baseName = "Code" Then
$code = $nodeSetting.text
ElseIf $nodeSetting.baseName = "Mode" Then
$mode = $nodeSetting.text
Else
EndIf
Next
EndIf
Next
EndIf
Next
EndIf
Next
EndFunc
If $CmdLine[0] = 0 Then
GetSettings($hour, $comp, $day)
Else
If $CmdLine[0] >= 1 Then
If $CmdLine[0] >= 2 Then
chkParam1()
Else
EndIf
EndIf
If $CmdLine[0] >= 3 Then
If $CmdLine[0] >= 4 Then
chkParam2()
Else
EndIf
EndIf
EndIf
Func chkParam1()
If $CmdLine[1] = "code" Then
If $CmdLine[2] <> "" Then
$code = $CmdLine[2]
EndIf
EndIf
If $CmdLine[1] = "mode" Then
If $CmdLine[2] <> "" Then
$mode = $CmdLine[2]
EndIf
EndIf
EndFunc
Func chkParam2()
If $CmdLine[3] = "code" Then
If $CmdLine[4] <> "" Then
$code = $CmdLine[4]
EndIf
EndIf
If $CmdLine[3] = "mode" And $CmdLine[4] <> "" Then
$mode = $CmdLine[4]
EndIf
EndFunc
Checkprocess()
WinWaitActive("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church", 30)
ErrorSearch()
CheckWindow()
Setup()
