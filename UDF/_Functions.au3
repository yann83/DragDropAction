#include-once

Func _InitializeFichierINI($sINIFile)
        Local $bNew = False
        If Not FileExists($sINIFile) Then
            $bNew = True
            _FileCreate($sINIFile)
        EndIf
        Local $hNIFile = _IniOpenFileEx($sINIFile)
        If Not $bNew Then
            Local $aTestSection = _IniReadSectionEx($hNIFile,"name",$INI_2DARRAYFIELD)
            If Not @error Then
                _IniCloseFileEx($hNIFile)
                Return False
            EndIf
        Else
            _IniWriteEx($hNIFile,"name","1","First entry")
            _IniWriteEx($hNIFile,"capture","1","")
            _IniWriteEx($hNIFile,"move","1","")
            _IniWriteEx($hNIFile,"rename","1","")
            _IniCloseFileEx($hNIFile)
            Return True
        EndIf
EndFunc

Func _Ajouter($sINIFile,$sName,$sRegex,$sMove,$sRename)
    Local $hNIFile = _IniOpenFileEx($sINIFile)
    Local $aSectionMove = _IniReadSectionEx($hNIFile,"name",$INI_2DARRAYFIELD)
    Local $nID = 1
    While 1
        If _ArraySearch($aSectionMove,$nID,1, 0, 0, 0, 1, 0) > -1 Then
            $nID += 1
        Else
            _IniWriteEx($hNIFile,"name",$nID,$sName)
            _IniWriteEx($hNIFile,"capture",$nID,$sRegex)
            _IniWriteEx($hNIFile,"move",$nID,$sMove)
            _IniWriteEx($hNIFile,"rename",$nID,$sRename)
            _IniCloseFileEx($hNIFile)
            Return (True)
        EndIf
    WEnd
EndFunc

Func _Modifier($sINIFile,$sID,$bStatus,$hInputName,$hInputCapture,$hInputMove,$InputRename)
    Local $hNIFile = _IniOpenFileEx($sINIFile)
    If $bStatus = True Then ; on rempli les inputs
        GUICtrlSetData($hInputName,_IniReadEx($hNIFile,"name",$sID,"Erreur lecture"))
        GUICtrlSetData($hInputCapture,_IniReadEx($hNIFile,"capture",$sID,"Erreur lecture"))
        GUICtrlSetData($hInputMove,_IniReadEx($hNIFile,"move",$sID,"Erreur lecture"))
        GUICtrlSetData($InputRename,_IniReadEx($hNIFile,"rename",$sID,"Erreur lecture"))
        _IniCloseFileEx($hNIFile)
        Return True
    ElseIf $bStatus = False Then; on enregistre les modis
        _IniWriteEx($hNIFile,"name",$sID,GUICtrlRead($hInputName))
        _IniWriteEx($hNIFile,"capture",$sID,GUICtrlRead($hInputCapture))
        _IniWriteEx($hNIFile,"move",$sID,GUICtrlRead($hInputMove))
        _IniWriteEx($hNIFile,"rename",$sID,GUICtrlRead($InputRename))
        _IniCloseFileEx($hNIFile)
        Return False
    EndIf
EndFunc

Func _Supprimer($sINIFile,$sID)
    Local $hNIFile = _IniOpenFileEx($sINIFile)
    Local $aSections = _IniReadSectionNamesEx($hNIFile)
    For $i = 1 To $aSections[0]
        _IniDeleteEx($hNIFile,$aSections[$i],$sID)
    Next
    $aSections = _IniReadSectionEx($hNIFile,"name",$INI_2DARRAYFIELD)
    If @error Then
        _IniWriteEx($hNIFile,"name","1","Name me")
        _IniCloseFileEx($hNIFile)
        Return False
    EndIf
    _IniCloseFileEx($hNIFile)
    Return True
EndFunc

Func _IniRearange($sINIFile)
    Local $hNIFile = _IniOpenFileEx($sINIFile)
    Local $aSections = _IniReadSectionNamesEx($hNIFile)
    Local $aSection
    Local $nSomme = 0
    Local $nNbLignes = 0
    For $i = 1 To $aSections[0]
        $aSection = _IniReadSectionEx($hNIFile,$aSections[$i],$INI_2DARRAYFIELD)
        If @error Then
            _IniCloseFileEx($hNIFile)
            Return SetError(1,0,$aSection&@error)
        EndIf
        $nNbLignes = Number($aSection[0][0])
        ;ConsoleWrite($aSection[0][0]&@CRLF)
        ;_ArrayDisplay($aSection)
        # Verification du nombre d'entrée
        If $nSomme = 0 Then
            $nSomme = $nNbLignes
        ElseIf $nSomme <> $nNbLignes Then
            _IniCloseFileEx($hNIFile)
            Return SetError(1,0,"La somme des entrées de "&$aSections[$i]&" ("&$nNbLignes&") est differente de la précédente : "&$aSections[$i-1]&" ("&$nSomme&")")
        Else
            $nSomme = $nNbLignes
        EndIf
        # Fin Verification du nombre d'entrée

        # Convertion de l'ID en nombre pour le tri
        For $k= 1 To $aSection[0][0]
            $aSection[$k][0] = Number($aSection[$k][0])
        Next
        # Fin Convertion de l'ID en nombre pour le tri

        _ArraySort($aSection,0,1,0,0) ; Tri par ID croissant
    Next

    For $i = 1 To $aSections[0]  ;Step -1 ; lecture de bas vers le haut
        $aSection = _IniReadSectionEx($hNIFile,$aSections[$i],$INI_2DARRAYFIELD)
        # Verification du de l'ordre des ID
        For $k= 1 To $aSection[0][0]
            If $k <> Number($aSection[$k][0]) Then $aSection[$k][0] = $k
        Next
        # Fin Verification du de l'ordre des ID

        # Reecriture de la section
        _IniDeleteEx($hNIFile,$aSections[$i],$INI_REMOVE )
        For $k= 1 To $aSection[0][0]
            _IniWriteEx($hNIFile,$aSections[$i],$aSection[$k][0],$aSection[$k][1])
        Next
        # Fin Reecriture de la section
    Next
    _IniCloseFileEx($hNIFile)
    SetError(0)
    Return True
EndFunc

Func _Rename($aArrayRename,$sString,$sRegex)
    Local $nGroup
    Local $sRename = ""
    Local $aArrayRegex = StringRegExp($sString,$sRegex,1)
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
    SetError(0)
    Return($sRename)
EndFunc

Func _GroupsOrder($sString)
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
        SetError(0)
        Return($atableau)
    Else
        Return SetError(1,0,"Pattern introuvable")
    EndIf
EndFunc

Func _AutoRegex($sString)
    Local $aParseCapture
    Local $nLength = 0
    Local $aString[8][2] = [[7,""],[".", ".*?"],[" ", "\s"],["-", "\-"],["(", "\("],[")", "\)"],["[", "\["],["]", "\]"]]
    For $i = 1 To $aString[0][0]
        $sString = StringReplace ($sString, $aString[$i][0], $aString[$i][1])
    Next
    If StringRegExp($sString,"(\d{1,})",0) = 1 Then
        $aParseCapture = StringRegExp($sString,"(\d{1,})",3) ; On regex le string
        $sString = StringRegExpReplace($sString,"(\d{1,})","%i%"); on remplace les chiffres par %i%
        For $i = 0 To UBound($aParseCapture) - 1
            $nLength = StringLen($aParseCapture[$i]) ;on calcul le nombre de chiffres
            $sString = StringReplace($sString,"%i%","\d{"&$nLength&"}",1); on remplace progressivement de gauche vers la droite %i%
        Next
    EndIf
    SetError(0)
    Return($sString)
EndFunc

Func _TestReg($sString,$sRegex)
    Local $aParseCapture
    If StringRegExp($sString,$sRegex,0) = 1 Then
        $aParseCapture = StringRegExp($sString,$sRegex,1) ; On regex le string
        SetError(0)
        Return ($aParseCapture)
    Else
        Return SetError(1,0,"Pas de capture")
    EndIf
EndFunc

Func _CountGroup($sString,$sRegex)
    Local $aParseCapture
    Local $nGroup
    If StringRegExp($sString,$sRegex,0) = 1 Then
        $aParseCapture = StringRegExp($sString,$sRegex,1) ; On regex le string
        $nGroup = UBound($aParseCapture)
        SetError(0)
        Return ($nGroup)
    Else
        Return SetError(1,0,"Pas de capture")
    EndIf
EndFunc

Func _GUICtrlRichEdit_WriteLine($hWnd, $sText, $iIncrement = 0, $sAttrib = "", $iColor = -1)
    ; Count the @CRLFs
    StringReplace(_GUICtrlRichEdit_GetText($hWnd, True), "", "")
    Local $iLines = @extended
    ; Adjust the text char count to account for the @CRLFs
    Local $iEndPoint = _GUICtrlRichEdit_GetTextLength($hWnd, True, True) - $iLines
    ; Add new text
    _GUICtrlRichEdit_AppendText($hWnd, $sText )
    ; Select text between old and new end points
    _GuiCtrlRichEdit_SetSel($hWnd, $iEndPoint, -1)
    ; Convert colour from RGB to BGR
    $iColor = Hex($iColor, 6)
    $iColor = '0x' & StringMid($iColor, 5, 2) & StringMid($iColor, 3, 2) & StringMid($iColor, 1, 2)
    ; Set colour
    If $iColor <> -1 Then _GuiCtrlRichEdit_SetCharColor($hWnd, $iColor)
    ; Set size
    If $iIncrement <> 0 Then _GUICtrlRichEdit_ChangeFontSize($hWnd, $iIncrement)
    ; Set weight
    If $sAttrib <> "" Then _GUICtrlRichEdit_SetCharAttributes($hWnd, $sAttrib)
    ; Clear selection
    _GUICtrlRichEdit_Deselect($hWnd)
EndFunc





