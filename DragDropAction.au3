#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ICO\Fleche.ico
#AutoIt3Wrapper_Outfile=DragDropAction.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Description=DragDropAction
#AutoIt3Wrapper_Res_Fileversion=1.0.0
#AutoIt3Wrapper_Res_File_Add=.\Resources\play.jpg, RT_RCDATA, PLAY
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------
 AutoIt Version : 3.3.14.5
 Auteur:         OUTIN Yann
 Description du programme :
	< Outil de rangement automatique des fichiers>
#ce ----------------------------------------------------------------------------
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <SendMessage.au3>
#include <File.au3>
#include <Constants.au3>
#include <String.au3>
#include <WinAPIRes.au3>
#include <WinAPIInternals.au3>
#include <WinAPIEx.au3>
#include <Array.au3>

Global $sFichierINI = @ScriptDir & "\" & StringTrimRight(@ScriptName,4) & ".ini"
Global $sFichierLog = @ScriptDir & "\" & StringTrimRight(@ScriptName,4) & ".log"

Global Const $SC_DRAGMOVE1 = 0xF012
Global $sFichierImage = @ScriptDir & "\play.jpg"
Global $hRet

Global $Width = @DesktopWidth
Global $Heigh = @DesktopHeight

If Not FileExists($sFichierINI) Then
    _FileCreate($sFichierINI)
    IniWrite($sFichierINI,"capture","1","")
    IniWrite($sFichierINI,"move","1","")
    IniWrite($sFichierINI,"rename","1","")
EndIf

;Les sections en tableau
Global $aSectionCapture = IniReadSection($sFichierINI,"capture")
If Not IsArray($aSectionCapture) Then
    _FileWriteLog($sFichierLog,"Il n'y a aucune valeurs renseignées dans la section capture")
    Exit
EndIf
Global $aSectionMove = IniReadSection($sFichierINI,"move")
If Not IsArray($aSectionCapture) Then
    _FileWriteLog($sFichierLog,"Il n'y a aucune valeurs renseignées dans la section move")
    Exit
EndIf
Global $aSectionRename = IniReadSection($sFichierINI,"rename")
If Not IsArray($aSectionCapture) Then
    _FileWriteLog($sFichierLog,"Il n'y a aucune valeurs renseignées dans la section rename")
    Exit
EndIf

$hRet = _SearchDuplicateValue($sFichierINI)
If $hRet <> 1 Then
    _FileWriteLog($sFichierLog,"Il y a des valeurs dupliquées dans la section : " & $hRet)
    MsgBox(16,"Erreurs detectées","Vérifiez le fichier de log : "&$sFichierLog)
    Exit
EndIf

Global $aParseCapture,$aParseFileName,$aParseRename
Global $aParseCaptureRows,$aParseRenameRows
Global $nID,$nCapture,$nMove,$nRenomme,$nFichiersTraites
Global $sMoveTo,$sFinalFileName
Global $sSettings = "DragDropActionSettings.exe"

If @Compiled Then
	$hRet = _FileInstallFromResource("PLAY",$sFichierImage)
	If @error Then _FileWriteLog($sFichierLog,"_FileInstallFromResource error : " & @error)
Else
	$sFichierImage = @ScriptDir & "\Resources\play.jpg"
EndIf

#Region ### START Koda GUI section ### Form=
Global $hDragDropGui = GUICreate("WM_DROPFILES", 50, 50,$Width-100, $Heigh-100, $WS_POPUP, $WS_EX_ACCEPTFILES)
Global $hDragDropPic = GUICtrlCreatePic($sFichierImage,0, 0, 50, 50)

GUICtrlSetState($hDragDropPic, $GUI_DROPACCEPTED)
GUICtrlSetBkColor($hDragDropPic, 0x0A000)
GUICtrlSetColor($hDragDropPic, 0)

Global $hContextMenu = GuiCtrlCreateContextMenu($hDragDropPic)
Global $hContextMenuSettings = GuiCtrlCreateMenuItem("Settings", $hContextMenu)
Global $hContextMenuExit = GuiCtrlCreateMenuItem("Exit", $hContextMenu)

Global $__aGUIDropFiles = 0

GUISetState(@SW_SHOW)

GUIRegisterMsg($WM_DROPFILES, 'WM_DROPFILES')
#EndRegion ### END Koda GUI section ###

_SetWindowPos($hDragDropGui, $HWND_TOPMOST + $HWND_TOP, 0, 0, 0, 0, BitOR($SWP_NOMOVE,$SWP_NOSIZE))

While 1
    $nMsg = GUIGetMsg()
    Switch $nMsg
        Case $GUI_EVENT_PRIMARYDOWN ;on bouge l'icône avec la souris
            _SendMessage($hDragDropGui, $WM_SYSCOMMAND, $SC_DRAGMOVE1, 0)

        Case $GUI_EVENT_CLOSE
            If @Compiled = 1 Then FileDelete($sFichierImage)
            Exit

		Case $hContextMenuExit
            If @Compiled = 1 Then FileDelete($sFichierImage)
			Exit

        Case $GUI_EVENT_DROPPED
            _FileWriteLog($sFichierLog,">>>> Evenement Drop detecté")
             _FileWriteLog($sFichierLog,"Reception de "&$__aGUIDropFiles[0]&" fichiers ci-dessous")
            For $i = 1 To $__aGUIDropFiles[0]
                _FileWriteLog($sFichierLog,$__aGUIDropFiles[$i])
            Next
            _FileWriteLog($sFichierLog,">>>> Début du traitement")

             $nCapture = 0
             $nMove = 0
             $nRenomme = 0
             $nFichiersTraites = 0

             For $i = 1 To $__aGUIDropFiles[0]
                ;################################### Phase1 capture
                For $capture = 1 To $aSectionCapture[0][0]
                    If StringRegExp($__aGUIDropFiles[$i],$aSectionCapture[$capture][1],0) = 1 Then ; Est ce que $__aGUIDropFiles[$i] va matcher un pattern de la section capture ?
                        $nID = $aSectionCapture[$capture][0]
                        $aParseCapture = StringRegExp($__aGUIDropFiles[$i],$aSectionCapture[$capture][1],1) ; On regex le string
                        $nCapture += 1
                        $aParseCaptureRows = UBound($aParseCapture)-1 ; Calcul du nombre de lignes / groupes capturés
                        ;################################### Phase2 move
                        $sMoveTo = ""
                        If _ArraySearch($aSectionMove,$nID,1, 0, 0, 0, 1, 0) > 0 Then ; recherche de l'Id de la capture précedente
                            If $aSectionMove[$nID][1] <> "" Then
                                $nMove += 1
                                $sMoveTo = $aSectionMove[$nID][1] & "\"
                                $aParseFileName = StringRegExp($__aGUIDropFiles[$i],".*\\(.*)(\..*)",1) ; on sépare le nom du fichier de son extension
                            Else
                                 _FileWriteLog($sFichierLog,"Pas de déplacement pour "&$__aGUIDropFiles[$i])
                            EndIf
                        EndIf
                        ;################################### Phase3 rename
                         If _ArraySearch($aSectionRename,$nID,1, 0, 0, 0, 1, 0) > 0 Then ; recherche de l'Id de la capture précedente
                             If $aSectionRename[$nID][1] <> "" Then
                                $sFinalFileName = ""
                                $aParseRename = _StringExplodeRegex($aSectionRename[$nID][1])
                                If Not @error Then
                                    If IsArray($aParseFileName) Then
                                        $nRenomme += 1
                                        $sFinalFileName =  _Rename($aParseRename,$aParseCapture)&$aParseFileName[1]
                                        If Not @error Then
                                            If $sMoveTo <> "" Then
                                                FileMove($__aGUIDropFiles[$i],$sMoveTo&$sFinalFileName,9)
                                                $nFichiersTraites += 1
                                            EndIf
                                        Else
                                            $nRenomme -= 1
                                            _FileWriteLog($sFichierLog,"ECHEC des groupes de "&$__aGUIDropFiles[$i])
                                        EndIf
                                    EndIf
                                Else
                                     _FileWriteLog($sFichierLog,"ECHEC du renommage "&$__aGUIDropFiles[$i])
                                EndIf
                            ElseIf $sMoveTo <> "" Then ; cas ou il n'y pas de renommage mais on bouge le fichier quand même
                                _FileWriteLog($sFichierLog,"Pas de renommage pour "&$__aGUIDropFiles[$i])
                                FileMove($__aGUIDropFiles[$i],$sMoveTo&$aParseFileName[0]&$aParseFileName[1],9)
                                $nFichiersTraites += 1
                            EndIf
                         EndIf
                    ;Else
                        ;_FileWriteLog($sFichierLog,"ECHEC de la capture de "&$__aGUIDropFiles[$i])
                    EndIf
                Next
            Next
            If $nCapture = 0 Then
                MsgBox(64,"Traitement terminé","Rien a traiter")
            Else
                MsgBox(64,"Traitement terminé","Fichiers capturés : "&$nCapture&@CRLF&"Fichiers à déplacer : "&$nMove&@CRLF& _
                                                            "Fichiers renommé : "&$nRenomme&@CRLF&"Total traités : "&$nFichiersTraites)
                _FileWriteLog($sFichierLog,"Fichiers capturés : "&$nCapture&" Fichiers à déplacer : "&$nMove& _
                                                            " Fichiers renommé : "&$nRenomme&" Total traités : "&$nFichiersTraites)
                _FileWriteLog($sFichierLog,">>>> Fin du traitement")
            EndIf

        Case $hContextMenuSettings
            Run(@ScriptDir & "\" & $sSettings)
    EndSwitch
WEnd

Func _Rename($aArrayRename,$aArrayRegex)
    Local $nGroup
    Local $sRename = ""
    If  $aArrayRename[0][0] <= UBound($aArrayRegex) Then
        For $i = 1 To $aArrayRename[0][0]
            $nGroup = Number($aArrayRename[$i][0])
            If $aArrayRename[$i][1] <>"" Then
                $sRename &= $aArrayRename[$i][1]
            Else
                $sRename &= $aArrayRegex[$nGroup]
            EndIf
        Next
    Else
        Return SetError(1,0,"Il n'y a pas assez de groupes")
    EndIf
    Return($sRename)
EndFunc

Func _StringExplodeRegex($sString)
    Local $atableau[1][2]
    Local $aDelim = StringRegExp($sString, "\((\d)\)", 3)
    If IsArray($aDelim) Then
        Local $aParts = StringSplit(StringRegExpReplace($sString, "(\(\d\))", "#"), "#", 3)
        _ArrayDelete($aParts,0)
        Local $NbRows = UBound($aDelim)
        ReDim $atableau[$NbRows][2]
        For $i = 0 To $NbRows - 1
                $atableau[$i][0] = $aDelim[$i]
                $atableau[$i][1] = $aParts[$i]
        Next
        _ArrayInsert($atableau,0,$NbRows)
        Return($atableau)
    Else
        Return SetError(1,0,"Pattern introuvable")
    EndIf
EndFunc

Func _SearchDuplicateValue($sConfigFile)
    Local $LectureSectionIni = IniReadSectionNames($sConfigFile)
    Local $aTableauOrg,$aTableauUnique
    For $i = 1 To $LectureSectionIni[0]
        $aTableauOrg = IniReadSection($sConfigFile,$LectureSectionIni[$i])
        If Not IsArray($aTableauOrg) Then SetError(2,0,"La section "&$LectureSectionIni[$i]&" n'est pas remplie")
        $aTableauUnique = _ArrayUnique($aTableauOrg,0,1)
        If Number($aTableauOrg[0][0]) > Number($aTableauUnique[0]) Then Return SetError(2,0,$LectureSectionIni[$i])
    Next
    SetError(0)
    Return(1)
EndFunc

;################### pour integration fichiers dans tableau #################
Func WM_DROPFILES($hWnd, $iMsg, $wParam, $lParam)
    #forceref $hWnd, $lParam
    Switch $iMsg
        Case $WM_DROPFILES
            Local Const $aReturn = _WinAPI_DragQueryFileEx($wParam)
            If UBound($aReturn) Then
                $__aGUIDropFiles = $aReturn
            Else
                Local Const $aError[1] = [0]
                $__aGUIDropFiles = $aError
            EndIf
    EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_DROPFILES

;################### pour mise au premier plan #################
Func _SetWindowPos($hwnd, $InsertHwnd, $X, $Y, $cX, $cY, $uFlags)
    $ret = DllCall("user32.dll", "long", "SetWindowPos", "hwnd", $hwnd, "hwnd", $InsertHwnd, _
    "int", $X, "int", $Y, "int", $cX, "int", $cY, "long", $uFlags)
EndFunc

;#################### pour file install ############################
Func _FileInstallFromResource($sResName, $sDest, $isCompressed = False, $iUncompressedSize = Default)
    Local $bBytes = _GetResourceAsBytes($sResName, $isCompressed, $iUncompressedSize)
    If @error Then Return SetError(@error, 0, 0)
    FileDelete($sDest)
    FileWrite($sDest, $bBytes)
EndFunc

Func _GetResourceAsBytes($sResName, $isCompressed = False, $iUncompressedSize = Default)

    Local $hMod = _WinAPI_GetModuleHandle(Null)
    Local $hRes = _WinAPI_FindResource($hMod, 10, $sResName)
    If @error Or Not $hRes Then Return SetError(1, 0, 0)
    Local $dSize = _WinAPI_SizeOfResource($hMod, $hRes)
    If @error Or Not $dSize Then Return SetError(2, 0, 0)
    Local $hLoad = _WinAPI_LoadResource($hMod, $hRes)
    If @error Or Not $hLoad Then Return SetError(3, 0, 0)
    Local $pData = _WinAPI_LockResource($hLoad)
    If @error Or Not $pData Then Return SetError(4, 0, 0)
    Local $tBuffer = DllStructCreate("byte[" & $dSize & "]")
    _WinAPI_MoveMemory(DllStructGetPtr($tBuffer), $pData, $dSize)
    If $isCompressed Then
        Local $oBuffer
       _WinAPI_LZNTDecompress($tBuffer, $oBuffer, $iUncompressedSize)
        If @error Then Return SetError(5, 0, 0)
        $tBuffer = $oBuffer
    EndIf
    Return DllStructGetData($tBuffer, 1)
EndFunc

Func _WinAPI_LZNTDecompress(ByRef $tInput, ByRef $tOutput, $iUncompressedSize = Default)
    ; if no uncompressed size given, use 16x the input buffer
    If $iUncompressedSize = Default Then $iUncompressedSize = 16 * DllStructGetSize($tInput)
    Local $tBuffer, $ret
    $tOutput = 0
    $tBuffer = DllStructCreate("byte[" & $iUncompressedSize & "]")
    If @error Then Return SetError(1, 0, 0)
    $ret = DllCall("ntdll.dll", "long", "RtlDecompressBuffer", "ushort", 2, "struct*", $tBuffer, "ulong", $iUncompressedSize, "struct*", $tInput, "ulong", DllStructGetSize($tInput), "ulong*", 0)
    If @error Then Return SetError(2, 0, 0)
    If $ret[0] Then Return SetError(3, $ret[0], 0)
    $tOutput = DllStructCreate("byte[" & $ret[6] & "]")
    If Not _WinAPI_MoveMemory(DllStructGetPtr($tOutput), DllStructGetPtr($tBuffer), $ret[6]) Then
        $tOutput = 0
        Return SetError(4, 0, 0)
    EndIf
    Return $ret[6]
EndFunc
