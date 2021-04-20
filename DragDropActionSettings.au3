#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ICO\settings.ico
#AutoIt3Wrapper_Outfile=DragDropActionSettings.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Description=DragDropActionSettings
#AutoIt3Wrapper_Res_Fileversion=1.0.0
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------
 AutoIt Version : 3.3.14.5
 Programme Version :
 Auteur:         OUTIN Yann
 Description du programme :
	<Réglages du programme DragDropAction>
#ce ----------------------------------------------------------------------------
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <File.au3>
#include <GuiRichEdit.au3>
#include <GuiListView.au3>

Global $rRetour

Global $sFichierINI = @ScriptDir & "\DragDropAction.ini"
If Not FileExists($sFichierINI) Then
    _FileCreate($sFichierINI)
    IniWrite($sFichierINI,"name","","")
    IniWrite($sFichierINI,"capture","","")
    IniWrite($sFichierINI,"move","","")
    IniWrite($sFichierINI,"rename","","")
Else
    $rRetour = _IniRearange($sFichierINI)
    If @error Then
        MsgBox(16,"Erreur",$rRetour)
        Exit
    EndIf
EndIf


#Region ### START Koda GUI section ### Form=G:\DOCUMENTS\PROGRAMMATION\AUTOIT\DragDropAction\Settings.kxf
Global $DragDropAction = GUICreate("DragDropAction Settings", 660, 560, -1, -1)

Global $Input0Name = GUICtrlCreateInput("Entrez le nom à transformer", 24, 20, 522, 30)
GUICtrlSetFont($Input0Name, 16, 400, 0, "MS Sans Serif")
Global $Input1Regex = GUICtrlCreateInput("", 24, 60, 522, 30)
GUICtrlSetFont($Input1Regex, 16, 400, 0, "MS Sans Serif")
Global $Input2Destination = GUICtrlCreateInput("", 24, 140, 522, 30)
GUICtrlSetFont($Input2Destination, 16, 400, 0, "MS Sans Serif")
Global $Input3Rename = GUICtrlCreateInput("", 24, 180, 522, 30)
GUICtrlSetFont($Input3Rename, 16, 400, 0, "MS Sans Serif")
Global $Input4Name = GUICtrlCreateInput("Nom de l'action", 265, 260, 280, 30)
GUICtrlSetFont($Input4Name, 16, 400, 0, "MS Sans Serif")


Global $Label1Regex = _GUICtrlRichEdit_Create($DragDropAction, "", 24, 100, 522, 30,BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL,$ES_READONLY))
_GUICtrlRichEdit_SetText($Label1Regex,  "")
Global $Label2Rename = GUICtrlCreateLabel("", 24, 220, 522, 30)
GUICtrlSetFont($Label2Rename, 16, 400, 0, "MS Sans Serif")
Global $Label3Name = GUICtrlCreateLabel("Nom : ", 198, 262, 60, 30)
GUICtrlSetFont($Label3Name, 16, 400, 0, "MS Sans Serif")

Global $Group1Rename  = GUICtrlCreateGroup("", 20, 215, 527, 35)
Global $Group2Edit = GUICtrlCreateGroup("Edition", 20, 255, 100, 45)
GUICtrlSetBkColor($Group2Edit, 0xFF8000)

Global $Button1AutoReg = GUICtrlCreateButton("Auto Regex", 560, 60, 70, 30)
Global $Button2TestReg = GUICtrlCreateButton("Test Regex", 560, 100, 70, 30)
Global $Button3Destination = GUICtrlCreateButton("Destination", 560, 140, 70, 30)
Global $Button4AutoRename = GUICtrlCreateButton("Auto Rename", 560, 180, 70, 30)
Global $Button5TestRename = GUICtrlCreateButton("Test Rename", 560, 220, 70, 30)

Global $Button6Add = GUICtrlCreateIcon (".\ICO\add.ico",-1, 30, 275, 20,20)
Global $Button7Mod = GUICtrlCreateIcon (".\ICO\edit.ico",-1, 60, 275, 20,20)
Global $Button8Del = GUICtrlCreateIcon (".\ICO\del.ico",-1, 90, 275, 20,20)

Global $ListView = GUICtrlCreateListView("ID|Nom du programme", 24, 310, 522, 240)
$rRetour = _DisplayListView($sFichierINI,"name",$ListView)
If Not $rRetour Then
    MsgBox(16,"Erreur","Fichier ini non valide")
    Exit
EndIf

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

#Region ### Declaration des Global ###
Global $nMsg
Global $MainExe = "DragDropAction.exe"
Global $ReadInput0Name
Global $ReadInput1Regex
Global $aGroupsOrder
Global $aColorLabel1Regex
Global $aColors[4] = [0xFF0000,0xF2B727,0x008040,0x001020]
Global $nColors
Global $SaveInput1Regex = ""
Global $sFileSelectFolder = @ScriptDir
Global $nGroupes
Global $sRename
Global $sSelection = ""
Global $bIsModif = False
#EndRegion ### END Declaration des Global###

While 1
    $nMsg = GUIGetMsg()
    Switch $nMsg
        Case $GUI_EVENT_CLOSE
            If ProcessExists($MainExe) Then
                ProcessClose($MainExe)
                Sleep(1000)
                Run(@ScriptDir&"\"&$MainExe)
            EndIf
            GUIDelete($DragDropAction)
            Exit
        Case $Button1AutoReg ;automatique regex
            If GUICtrlRead($Input0Name) = "" Then
                MsgBox(16,"Erreur","Veuillez renseigner le 1er champ en haut")
            Else
                $ReadInput0Name = GUICtrlRead($Input0Name)
                GUICtrlSetData($Input1Regex,_AutoRegex($ReadInput0Name))
            EndIf
        Case $Button2TestReg ;test du regex
            If GUICtrlRead($Input1Regex) = "" And GUICtrlRead($Input0Name) = "" Then
                MsgBox(16,"Erreur","Veuillez renseigner les champs en haut.")
            Else
                $ReadInput0Name = GUICtrlRead($Input0Name)
                $ReadInput1Regex = GUICtrlRead($Input1Regex)
                $aColorLabel1Regex = _TestReg($ReadInput0Name,$ReadInput1Regex)
                If @error Then
                    _GUICtrlRichEdit_SetText($Label1Regex, "") ;vide l'input
                    _GUICtrlRichEdit_WriteLine($Label1Regex, $aColorLabel1Regex , "", "", 0xFF0000)
                Else
                    _GUICtrlRichEdit_SetText($Label1Regex, "") ;vide l'input
                    $nColors = 0
                    For $i = 0 To UBound($aColorLabel1Regex) - 1
                        If $i = 0 Then
                            _GUICtrlRichEdit_WriteLine($Label1Regex, $aColorLabel1Regex[$i] , +4, "", $aColors[$nColors])
                        Else
                            _GUICtrlRichEdit_WriteLine($Label1Regex, $aColorLabel1Regex[$i] , "", "", $aColors[$nColors])
                        Endif
                        If $i >= 4 Then
                            $nColors = 0
                        Else
                            $nColors += 1
                        EndIf
                    Next
                EndIf
            EndIf
        Case $Button3Destination
            $sFileSelectFolder = FileSelectFolder("Choisir le dossier de destination", $sFileSelectFolder)
            If @error Then
                MsgBox(16,"", "Aucun dossier n'a été sélectionné.")
            Else
                GUICtrlSetData($Input2Destination,$sFileSelectFolder)
            EndIf
        Case $Button4AutoRename
            If GUICtrlRead($Input1Regex) = "" And GUICtrlRead($Input0Name) = "" Then
                MsgBox(16,"Erreur","Veuillez renseigner les deux premiers champs en haut.")
            Else
                $ReadInput0Name = GUICtrlRead($Input0Name)
                $ReadInput1Regex = GUICtrlRead($Input1Regex)
                If StringInStr($ReadInput1Regex,"(") > 0 And StringInStr($ReadInput1Regex,")") > 0 Then
                    $nGroupes = _CountGroup($ReadInput0Name,$ReadInput1Regex)
                    $sRename = ""
                    For $i = 0 To $nGroupes - 1
                        $sRename &= "("&$i&")"
                    Next
                    GUICtrlSetData($Input3Rename,$sRename)
                Else
                    MsgBox(16,"Erreur","Aucun groupe semble créé.")
                EndIf
            EndIf
        Case $Button5TestRename
            If GUICtrlRead($Input1Regex) = "" And GUICtrlRead($Input0Name) = "" And GUICtrlRead($Input3Rename) = "" Then
                MsgBox(16,"Erreur","Veuillez renseigner les champs en haut.")
            Else
                $aGroupsOrder = _GroupsOrder(GUICtrlRead($Input3Rename))
                GUICtrlSetData($Label2Rename,_Rename($aGroupsOrder,GUICtrlRead($Input0Name),GUICtrlRead($Input1Regex)))
            EndIf
        Case $Button6Add
            If GUICtrlRead($Input1Regex) = "" And GUICtrlRead($Input3Rename) = "" And GUICtrlRead($Input4Name) = ""  Then
                MsgBox(16,"Erreur","Veuillez renseigner les champs Regex et Rename.")
            Else
                _Ajouter($sFichierINI,GUICtrlRead($Input4Name) ,GUICtrlRead($Input1Regex) ,GUICtrlRead($Input2Destination),GUICtrlRead($Input3Rename))
                _DisplayListView($sFichierINI,"name",$ListView)
                MsgBox(64,"","L'entrée "&GUICtrlRead($Input4Name)&" a été ajoutée")
            EndIf
        Case $Button7Mod
            If $bIsModif = True Then
                $bIsModif = False
                ;si la slection est pas vide on enregistre
                If $sSelection <> "" Then _Modifier($sFichierINI,$sSelection,$bIsModif,$Input4Name,$Input1Regex,$Input2Destination,$Input3Rename)
                 GUICtrlSetState($Button6Add,$GUI_ENABLE)
                GUICtrlSetState($Button8Del,$GUI_ENABLE)
                $sSelection = ""
                _DisplayListView($sFichierINI,"name",$ListView)
                MsgBox(64,"","L'entrée "&$sSelection&" a été modifiée")
            ElseIf $bIsModif = False Then
                $sSelection = _FormatSeclect($ListView)
                If @error then
                    MsgBox(16,"","selectionner une ligne")
                Else
                    $bIsModif = True
                    ;on met les entrée dans les input
                    _Modifier($sFichierINI,$sSelection,$bIsModif,$Input4Name,$Input1Regex,$Input2Destination,$Input3Rename)
                    GUICtrlSetState($Button6Add,$GUI_DISABLE)
                    GUICtrlSetState($Button8Del,$GUI_DISABLE)
                    MsgBox(64,"","L'entrée "&$sSelection&" a été affichée")
                EndIf
            EndIf
        Case $Button8Del
            $sSelection = _FormatSeclect($ListView)
            If @error then
                MsgBox(16,"","selectionner une ligne")
            Else
                _Supprimer($sFichierINI,$sSelection)
                $rRetour = _IniRearange($sFichierINI)
                If @error Then
                    MsgBox(16,"Erreur",$rRetour)
                    Exit
                EndIf
                _DisplayListView($sFichierINI,"name",$ListView)
                $sSelection = ""
                MsgBox(64,"","L'entrée "&$sSelection&" a été supprimée")
            EndIf

    EndSwitch
WEnd

#Region ### Fonctions ###
Func _Ajouter($sINIFile,$sName,$sRegex,$sMove,$sRename)
    Local $aSectionMove = IniReadSection($sINIFile,"name")
    Local $nID = 1
    While 1
        If _ArraySearch($aSectionMove,$nID,1, 0, 0, 0, 1, 0) > -1 Then
            $nID += 1
        Else
            IniWrite($sINIFile,"name",$nID,$sName)
            IniWrite($sINIFile,"capture",$nID,$sRegex)
            IniWrite($sINIFile,"move",$nID,$sMove)
            IniWrite($sINIFile,"rename",$nID,$sRename)
            Return (True)
        EndIf
    WEnd
EndFunc

Func _Modifier($sINIFile,$sID,$bStatus,$hInputName,$hInputCapture,$hInputMove,$InputRename)
    If $bStatus = True Then ; on rempli les inputs
        GUICtrlSetData($hInputName,IniRead($sINIFile,"name",$sID,"Erreur lecture"))
        GUICtrlSetData($hInputCapture,IniRead($sINIFile,"capture",$sID,"Erreur lecture"))
        GUICtrlSetData($hInputMove,IniRead($sINIFile,"move",$sID,"Erreur lecture"))
        GUICtrlSetData($InputRename,IniRead($sINIFile,"rename",$sID,"Erreur lecture"))
        Return True
    ElseIf $bStatus = False Then; on enregistre les modis
        IniWrite($sINIFile,"name",$sID,GUICtrlRead($hInputName))
        IniWrite($sINIFile,"capture",$sID,GUICtrlRead($hInputCapture))
        IniWrite($sINIFile,"move",$sID,GUICtrlRead($hInputMove))
        IniWrite($sINIFile,"rename",$sID,GUICtrlRead($InputRename))
        Return False
    EndIf
EndFunc

Func _Supprimer($sINIFile,$sID)
    Local $aSections = IniReadSectionNames($sINIFile)
    For $i = 1 To $aSections[0]
        IniDelete($sINIFile,$aSections[$i],$sID)
    Next
EndFunc

Func _IniRearange($sINIFile)
    Local $aSections = IniReadSectionNames($sINIFile)
    Local $aSection
    Local $nSomme = 0
    For $i = 1 To $aSections[0]
        $aSection = IniReadSection($sINIFile,$aSections[$i])

        # Verification du nombre d'entrée
        If $nSomme = 0 Then
            $nSomme = Number($aSection[0][0])
        ElseIf $nSomme <> Number($aSection[0][0]) Then
            Return SetError(1,0,"La somme des entrées de "&$aSections[$i]&" est differente de la précédente")
        Else
            $nSomme = Number($aSection[0][0])
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
        $aSection = IniReadSection($sINIFile,$aSections[$i])
        # Verification du de l'ordre des ID
        For $k= 1 To $aSection[0][0]
            If $k <> Number($aSection[$k][0]) Then $aSection[$k][0] = $k
        Next
        # Fin Verification du de l'ordre des ID

        # Reecriture de la section
        IniDelete($sINIFile,$aSections[$i])
        For $k= 1 To $aSection[0][0]
            IniWrite($sINIFile,$aSections[$i],$aSection[$k][0],$aSection[$k][1])
        Next
        # Fin Reecriture de la section
    Next

    SetError(0)
    Return True
EndFunc

Func _FormatSeclect($hListView)
    Local $sLecture = GUICtrlRead(GUICtrlRead($hListView))
    If $sLecture = "0" Then Return SetError(1,0,"pas de ligne selectionne")
    $sLecture = StringTrimRight($sLecture,1)
    Local $aSplitRow = StringSplit($sLecture,"|")
    SetError(0)
    Return($aSplitRow[1])
EndFunc

Func _DisplayListView($sINIFile,$sSectionName,$hListView)
    _GUICtrlListView_DeleteAllItems($hListView)
    Local $aINISection = IniReadSection($sINIFile,$sSectionName)
    If Not @error Then
        If IsArray($aINISection) Then
            For $i = 1 To $aINISection[0][0]
                GUICtrlCreateListViewItem($aINISection[$i][0]&"|"&$aINISection[$i][1], $hListView)
            Next
            SetError(0)
            Return(True)
        EndIf
    Else
        Return SetError(1,0,"Erreur lecture du fichier ini")
    EndIf
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
    $sString = StringReplace ($sString, ".", "\.")
    $sString = StringReplace ($sString, " ", "\W")
    $sString = StringReplace ($sString, "-", "\W")
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

#EndRegion ### END Fonctions###
