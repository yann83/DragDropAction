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
 Auteur:         Yann83
 Description du programme :
	<Réglages du programme DragDropAction>
#ce ----------------------------------------------------------------------------
#NoTrayIcon
#include <.\UDF\_Languages.au3>
#include <.\UDF\IniEx.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <File.au3>
#include <GuiRichEdit.au3>
#include <GuiListView.au3>
#include <Array.au3>
#include <.\UDF\_Functions.au3>

Global $TipHelp = "."&@TAB&@TAB&"Any character, except newline"&@CRLF& _
                            "\w"&@TAB&@TAB&"Word"&@CRLF& _
                            "\d"&@TAB&@TAB&"Digit"&@CRLF& _
                            "\s"&@TAB&@TAB&"Whitespace"&@CRLF& _
                            "\W"&@TAB&@TAB&"Not word"&@CRLF& _
                            "\D"&@TAB&@TAB&"Not digit"&@CRLF& _
                            "\S"&@TAB&@TAB&"Not whitespace"&@CRLF& _
                            "[abc]"&@TAB&@TAB&"Any of a, b, or c"&@CRLF& _
                            "[a-e]"&@TAB&@TAB&"Characters between a and e"&@CRLF& _
                            "[1-9]"&@TAB&@TAB&"Digit between 1 and 9"&@CRLF& _
                            "[^abc]"&@TAB&@TAB&"Any character except a, b or c"&@CRLF& _
                            "\. \\"&@TAB&@TAB&"Escape special character used by regex"&@CRLF& _
                            "\t"&@TAB&@TAB&"Tab"&@CRLF& _
                            "\n"&@TAB&@TAB&"Newline"&@CRLF& _
                            "\r"&@TAB&@TAB&"Carriage return"&@CRLF& _
                            "(abc)"&@TAB&@TAB&"Capture group"&@CRLF& _
                            "(a|b)"&@TAB&@TAB&"Match a or b"&@CRLF& _
                            "(?:abc)"&@TAB&@TAB&"Match abc, but don’t capture"&@CRLF& _
                            "a*"&@TAB&@TAB&"Match 0 or more"&@CRLF& _
                            "a+"&@TAB&@TAB&"Match 1 or more"&@CRLF& _
                            "a?"&@TAB&@TAB&"Match 0 or 1"&@CRLF& _
                            "a{5}"&@TAB&@TAB&"Match exactly 5"&@CRLF& _
                            "a{,3}"&@TAB&@TAB&"Match up to 3"&@CRLF& _
                            "a{3,}"&@TAB&@TAB&"Match 3 or more"&@CRLF& _
                            "a{1,3}"&@TAB&@TAB&"Match between 1 and 3"&@CRLF& _
                            "a(?=b)"&@TAB&@TAB&"Match a in baby but not in bay"&@CRLF& _
                            "a(?!b)"&@TAB&@TAB&"Match a in Stan but not in Stab"&@CRLF& _
                            "(?<=a)b"&@TAB&@TAB&"Match b in crabs but not in cribs"&@CRLF& _
                            "(?<!a)b"&@TAB&@TAB&"Match b in fib but not in fab"&@CRLF& _
                            "(?i)"&@TAB&@TAB&"Case-insensitive mode"

Global $rRetour

Global $sFichierINI = @ScriptDir & "\DragDropAction.ini"
If @error Then MsgBox(16,"Erreur","Le fichier "&$sFichierINI&" renvoi "&@error)
Global $bFichierINIError = _InitializeFichierINI($sFichierINI)

If Not $bFichierINIError Then
    $rRetour = _IniRearange($sFichierINI)
    If @error Then
        MsgBox(16,"Erreur",$rRetour)
        Exit
    EndIf
EndIf

#Region ### START Koda GUI section ### Form=G:\DOCUMENTS\PROGRAMMATION\AUTOIT\DragDropAction\Settings.kxf
Global $DragDropAction = GUICreate($_Gui_MainName, 660, 590, -1, -1)

Global $Input0Name = GUICtrlCreateInput($_Gui_Input0Name, 24, 20, 522, 30)
GUICtrlSetFont($Input0Name, 16, 400, 0, "MS Sans Serif")
Global $Input1Regex = GUICtrlCreateInput("", 24, 60, 522, 30)
GUICtrlSetFont($Input1Regex, 16, 400, 0, "MS Sans Serif")
GUICtrlSetTip( $Input1Regex, $TipHelp)
Global $Input2Destination = GUICtrlCreateInput("", 24, 140, 522, 30)
GUICtrlSetFont($Input2Destination, 16, 400, 0, "MS Sans Serif")
Global $Input3Rename = GUICtrlCreateInput("", 24, 180, 522, 30)
GUICtrlSetFont($Input3Rename, 16, 400, 0, "MS Sans Serif")
Global $Input4Name = GUICtrlCreateInput($_Gui_Input4Name, 265, 260, 280, 30)
GUICtrlSetFont($Input4Name, 16, 400, 0, "MS Sans Serif")

Global $Label1Regex = _GUICtrlRichEdit_Create($DragDropAction, "", 24, 100, 522, 30,BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL,$ES_READONLY))
_GUICtrlRichEdit_SetText($Label1Regex,  "")
Global $Label2Rename = GUICtrlCreateLabel("", 24, 220, 522, 30)
GUICtrlSetFont($Label2Rename, 16, 400, 0, "MS Sans Serif")
Global $Label3Name = GUICtrlCreateLabel($_Gui_Label3Name, 156, 262, 100, 30)
GUICtrlSetFont($Label3Name, 16, 400, 0, "MS Sans Serif")
Global $Label4Search = GUICtrlCreateLabel($_Gui_Label4Search, 156, 302, 100, 30)
GUICtrlSetFont($Label4Search, 16, 400, 0, "MS Sans Serif")

Global $Group1Rename  = GUICtrlCreateGroup("", 20, 215, 527, 35)
Global $Group2Edit = GUICtrlCreateGroup($_Gui_Group2Edit, 20, 255, 100, 45)
GUICtrlSetBkColor($Group2Edit, 0xFF8000)

Global $Button1AutoReg = GUICtrlCreateButton($_Gui_Button1AutoReg, 560, 60, 70, 30)
Global $Button2TestReg = GUICtrlCreateButton($_Gui_Button2TestReg, 560, 100, 70, 30)
Global $Button3Destination = GUICtrlCreateButton($_Gui_Button3Destination, 560, 140, 70, 30)
Global $Button4AutoRename = GUICtrlCreateButton($_Gui_Button4AutoRename, 560, 180, 70, 30)
Global $Button5TestRename = GUICtrlCreateButton($_Gui_Button5TestRename, 560, 220, 70, 30)

Global $Button6Add = GUICtrlCreateIcon (".\ICO\add.ico",-1, 30, 275, 20,20)
GUICtrlSetTip($Button6Add, $_Gui_Add)
Global $Button7Mod = GUICtrlCreateIcon (".\ICO\edit.ico",-1, 60, 275, 20,20)
GUICtrlSetTip($Button7Mod, $_Gui_Modify)
Global $Button8Del = GUICtrlCreateIcon (".\ICO\del.ico",-1, 90, 275, 20,20)
GUICtrlSetTip($Button8Del, $_Gui_Delete)

; Creation de la zone de recherche
Global $idEdit = GUICtrlCreateEdit( "", 265, 300, 280, 30, BitXOR( $GUI_SS_DEFAULT_EDIT, $WS_HSCROLL, $WS_VSCROLL ) )
GUICtrlSetFont($idEdit, 16, 400, 0, "MS Sans Serif")
Global $hEdit = GUICtrlGetHandle($idEdit)
Global $idEditSearch = GUICtrlCreateDummy()

; Handle $WM_COMMAND messages from Edit control
; To be able to read the search string dynamically while it's typed in
GUIRegisterMsg( $WM_COMMAND, "WM_COMMAND" )

Global $idListView = GUICtrlCreateListView("", 24, 340, 522, 240,$LVS_OWNERDATA)
Global $hListView = GUICtrlGetHandle( $idListView )
_GUICtrlListView_SetExtendedListViewStyle( $hListView, BitOR( $LVS_EX_DOUBLEBUFFER, $LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES ) )
_GUICtrlListView_AddColumn( $hListView, $_Gui_ListViewCol1, $_Gui_ListViewCol1Size)
_GUICtrlListView_AddColumn( $hListView, $_Gui_ListViewCol2, $_Gui_ListViewCol2Size )

; Handle $WM_NOTIFY messages from ListView
; Necessary to display the rows in a virtual ListView
GUIRegisterMsg( $WM_NOTIFY, "WM_NOTIFY" )

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

;initialyze search array
Global $aMoteurRecherche = _TableauObjetsEtRecherche($sFichierINI,"name")
If @error Then MsgBox(16,"Erreur",$aMoteurRecherche)
                            ;$aObjets ;$iNbObjets $aTableauRecherche $iNbTableauRecherche
Global $aObjets = $aMoteurRecherche[0]
Global $aTableauRecherche = $aMoteurRecherche[2]
; Display items
GUICtrlSendMsg( $idListView, $LVM_SETITEMCOUNT, $aMoteurRecherche[3], 0 )

#Region ### Declaration des Global ###
Global $iIndiceListView,$bIsObjectSelected,$aDataListView
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
Global $sLectureRecherche = ""
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

#Region ### Search section ###
        Case $idEditSearch
            $sLectureRecherche = GUICtrlRead( $idEdit )
            If $sLectureRecherche = "" Then
                ; Empty search string, display all rows
                For $i = 0 To $aMoteurRecherche[1] - 1
                    $aTableauRecherche[$i] = $i
                Next
                $aMoteurRecherche[3] = $aMoteurRecherche[1]
            Else
                ; Find rows matching the search string
                $aMoteurRecherche[3] = 0
                For $i = 0 To $aMoteurRecherche[1] - 1
                    If StringInStr( $aObjets[$i][1], $sLectureRecherche ) Then ; Normal search
                        $aTableauRecherche[$aMoteurRecherche[3]] = $i
                        $aMoteurRecherche[3] += 1
                    EndIf
                Next
            EndIf
            ; Display items
            GUICtrlSendMsg( $idListView, $LVM_SETITEMCOUNT, $aMoteurRecherche[3], 0 )
            ;ConsoleWrite( StringFormat( "%4d", $aMoteurRecherche[3] ) & " rows matching """ & $sLectureRecherche & """" & @CRLF )
#EndRegion ### END Search section ###

#Region ### Right Buttons ###
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
                        If $i = 3 And Mod($i,3) = 0 Then
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
#EndRegion ### END Right Buttons ###

#Region ### CRUD section ###
        Case $Button6Add
            If GUICtrlRead($Input1Regex) = "" And GUICtrlRead($Input3Rename) = "" And GUICtrlRead($Input4Name) = ""  Then
                MsgBox(16,"Erreur","Veuillez renseigner les champs Regex et Rename.")
            Else
                _Ajouter($sFichierINI,GUICtrlRead($Input4Name) ,GUICtrlRead($Input1Regex) ,GUICtrlRead($Input2Destination),GUICtrlRead($Input3Rename))
                _Recharge($sFichierINI)
                MsgBox(64,"","L'entrée "&GUICtrlRead($Input4Name)&" a été ajoutée")
            EndIf
        Case $Button7Mod
            If $bIsModif = True Then
                $bIsModif = False
                ;si la slection est pas vide on enregistre
                If $aDataListView[1] <> "" Then _Modifier($sFichierINI,$aDataListView[1],$bIsModif,$Input4Name,$Input1Regex,$Input2Destination,$Input3Rename)
                 GUICtrlSetState($Button6Add,$GUI_ENABLE)
                GUICtrlSetState($Button8Del,$GUI_ENABLE)
                $aDataListView[1] = ""
                 _Recharge($sFichierINI)
                MsgBox(64,"","L'entrée "&$sSelection&" a été modifiée")
            ElseIf $bIsModif = False Then
                $aDataListView = _Selection()
                If $aDataListView = 0 Then
                    MsgBox(64,"Attention","Veuillez selectionner une ligne")
                Else
                    $bIsModif = True
                    ;on met les entrée dans les input
                    _Modifier($sFichierINI,$aDataListView[1],$bIsModif,$Input4Name,$Input1Regex,$Input2Destination,$Input3Rename)
                    GUICtrlSetState($Button6Add,$GUI_DISABLE)
                    GUICtrlSetState($Button8Del,$GUI_DISABLE)
                    MsgBox(64,"","L'entrée "&$aDataListView[1]&" a été affichée")
                EndIf
            EndIf
        Case $Button8Del
            $aDataListView = _Selection()
            If $aDataListView = 0 Then
                MsgBox(64,"Attention","Veuillez selectionner une ligne")
            Else
                If _Supprimer($sFichierINI,$aDataListView[1]) Then
                    $rRetour = _IniRearange($sFichierINI)
                    If @error Then
                        MsgBox(16,"Erreur",$rRetour)
                        Exit
                    EndIf
                EndIf
                MsgBox(64,"","L'entrée "&$aDataListView[1]&" a été supprimée")
                _Recharge($sFichierINI)
            EndIf
#EndRegion ### END CRUD section ###
    EndSwitch
WEnd

#Region ### Functions with globals###
Func _Selection()
    $iIndiceListView = Int(_GUICtrlListView_GetSelectedIndices($idListView)); récupére l'indice de l'item de la liste
    $bIsObjectSelected = _GUICtrlListView_GetItemSelected( $idListView, $iIndiceListView); renvoie True si l'item est selectionné
    If $bIsObjectSelected = True Then Return _GUICtrlListView_GetItemTextArray($idListView, $iIndiceListView);récupére dans un tableau les données de l'item
    Return (0)
EndFunc

Func _Recharge($sPathINI)
    GUICtrlSetData($idEdit,"")
     _GUICtrlListView_DeleteAllItems($hListView)
    $aMoteurRecherche = _TableauObjetsEtRecherche($sPathINI,"name")
    $aObjets = $aMoteurRecherche[0]
    $aTableauRecherche = $aMoteurRecherche[2]
    GUICtrlSendMsg( $idListView, $LVM_SETITEMCOUNT, $aMoteurRecherche[3], 0 )
EndFunc

Func _TableauObjetsEtRecherche($sPathINI,$sIniSection)
    Local $ifNbObjets = 0, $afObjets[1000][2]
    $hPathINI = _IniOpenFileEx($sPathINI)
    Local $aDATA = _IniReadSectionEx($hPathINI,$sIniSection,$INI_2DARRAYFIELD)
    If Not @error  Then
        ; Enumerate through the array displaying the section names.
        For $i = 1 To $aDATA[0][0]
            $afObjets[$i-1][0] = $aDATA[$i][0]
            $afObjets[$i-1][1] = $aDATA[$i][1]
        Next
        $ifNbObjets = $aDATA[0][0]
        Local $afRetTableau [4] = [$afObjets,$ifNbObjets,"",""]

        Local $afTableauRecherche[$ifNbObjets]
        Local $ifNbTableauRecherche = 0
          ; Set search array to include all items
        For $i = 0 To $ifNbObjets - 1
            $afTableauRecherche[$i] = $i
        Next
        $ifNbTableauRecherche = $ifNbObjets
        $afRetTableau[2] = $afTableauRecherche
        $afRetTableau[3] = $ifNbTableauRecherche
    Else
        _IniCloseFileEx($hPathINI)
        Return SetError(1,0,"Erreur [_TableauObjetsEtRecherche] lecture tableau "&@error)
    EndIf
     _IniCloseFileEx($hPathINI)
    SetError(0)
     Return $afRetTableau
EndFunc

Func WM_COMMAND( $hWnd, $iMsg, $wParam, $lParam )
    Local $hWndFrom = $lParam
    Local $iCode = BitShift( $wParam, 16 ) ; High word
    Switch $hWndFrom
        Case $hEdit;variable
            Switch $iCode
                Case $EN_CHANGE
                    GUICtrlSendToDummy( $idEditSearch ); variable
            EndSwitch
    EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc

Func WM_NOTIFY( $hWnd, $iMsg, $wParam, $lParam )
    Local Static $tText = DllStructCreate( "wchar[50]" )
    Local Static $pText = DllStructGetPtr( $tText )

    Local $tNMHDR, $hWndFrom, $iCode
    $tNMHDR = DllStructCreate( $tagNMHDR, $lParam )
    $hWndFrom = HWnd( DllStructGetData( $tNMHDR, "hWndFrom" ) )
    $iCode = DllStructGetData( $tNMHDR, "Code" )

    Switch $hWndFrom
        Case $hListView; variable
            Switch $iCode
                Case $LVN_GETDISPINFOW
                    Local $tNMLVDISPINFO = DllStructCreate( $tagNMLVDISPINFO, $lParam )
                    If BitAND( DllStructGetData( $tNMLVDISPINFO, "Mask" ), $LVIF_TEXT ) Then
                        Local $sItem = $aObjets[$aTableauRecherche[DllStructGetData($tNMLVDISPINFO,"Item")]][DllStructGetData($tNMLVDISPINFO,"SubItem")]; variable
                        DllStructSetData( $tText, 1, $sItem ); variable
                        DllStructSetData( $tNMLVDISPINFO, "Text", $pText )
                        DllStructSetData( $tNMLVDISPINFO, "TextMax", StringLen( $sItem ) ); variable
                    EndIf
            EndSwitch
    EndSwitch

    Return $GUI_RUNDEFMSG
EndFunc
#EndRegion ### END Fonctions###
