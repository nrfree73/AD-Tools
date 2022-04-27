#pragma compile(FileVersion, 6.0)
#pragma compile(ProductVersion, 6.0)
#pragma compile(FileDescription, [AD Tools v6.0])
#pragma compile(ProductName, [AD Tools v6.0])
#pragma compile(LegalCopyright, (21/04/2022) - Nicolas RISTOVSKI)
#pragma compile(CompanyName, Nicolas RISTOVSKI / EAPI69)

Global $lastdatecompile = "	(c) 2018~2022 " ;about box
Global $lastdateupdate = " 21-04-2022 " ;main routine information
Global $ADTVersion="v6.0"

#Region n1

    #Region Includes autoit
#include <AD.au3>
#include <TabConstants.au3>
#include <GUIConstants.au3>
#include <Array.au3>
#include <Date.au3>
#include <FileConstants.au3>
#include <String.au3>
#include <Misc.au3>
#include <File.au3>
#include <security.au3>
#include <GuiTreeView.au3>
#include <GUIListBox.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiEdit.au3>
#include <GuiRichEdit.au3>
#include "resources.au3"
   #EndRegion Include

	#Region Obfuscator Script autoit
		#AutoIt3Wrapper_Run_Obfuscator=Y
		#Obfuscator_Parameters=/SF /SV /OM /CS=0 /CN=0
	#EndRegion

	#Region Directives **** AutoIt3Wrapper_GUI ****
		#AutoIt3Wrapper_Compression=4
		#AutoIt3Wrapper_Res_Description=[ AD Tools ]
		#AutoIt3Wrapper_Res_Fileversion=2.0
		#AutoIt3Wrapper_Res_LegalCopyright=UTA69
		#AutoIt3Wrapper_Res_Language=1036
		#AutoIt3Wrapper_Res_SaveSource=N
		#Au3Stripper_On
	#EndRegion Directives

Opt("TrayMenuMode", 3)
	#RequireAdmin
	Global $adroot
	Global $dcdomaintmp
	Global $dcdomain = "DC=dct,DC=adt,DC=local"
	Global $racinead = "DCT"
	Global $defautdc = "DCT.adt.local"
	Global $defautdcinit
	Global $sad_useridparam, $racinead, $adrootadmin
	Global $sad_passwordparam, $adrootpwd
	Global $sad_dnsdomainparam, $dcdomain
	Global $sad_hostserverparam
	Global $sad_configurationparam
	Global $hgui
	Global $iobject = ""
	Global $ousourcedir
	Global $ou
	Global $sout
	Global $groupsidrh1
	Global $iterad = 1
	Global $sout
	Global $isdct = 0
	Global $isdct2 = ""
	Global $lastiter = ""
	Global $readad = 0
	Global $aresult
	Global $directives = ""
	Global $directivesDSIBA = ""
	Global $scomboreaddirectives = ""
	Global $restoredirectives = ""
	Global $actiongroups = ""
	Global $restoredirectives = ""
	Global $groupsidrh1
	Global $idrh1
    Global $domainname=""
    Global $aff
    Global $date_expire = ""
    Global $date_prolonger = ""
    Global $cyclesIDRG=0
    Global $auserids ;ANR
    Global $numdirmetier =""
    Global $isDomainOK=0
    Global $ouDirmetier="" ;check dir-metier OU

	Func _terminate()
	   $t = MsgBox(4, "Quitter ?", "[Oui]: Terminer AD Tools !" & @CRLF & @CRLF & "[Non]: ne fait rien")
					If $t = 6 Then
						Exit
					EndIf
	EndFunc

	Global $hdll = DllOpen("user32.dll")
	Global $historik
	Global $hguilg

	Func about()
;
	EndFunc

	_ad_open()
	Global $atemproot = _ad_listrootdseattributes()
	If IsArray($atemproot) Then
		$isdct = $atemproot[3][1]
		$isdct = StringReplace($isdct, "DC=", "")
		$isdct = StringReplace($isdct, ",", ".")
		$isdct2 = $isdct
		Global $domainname = $isdct
		$domainname = StringSplit($domainname, ".")
		If IsArray($domainname) Then
			$domainname = $domainname[1]
			$isDomainOK=1
		EndIf
		$defautdc = $isdct
		$defautdcinit = $defautdc
	EndIf
	If StringCompare($isdct, "dct.adt.local") = 0 Then
		$isdct = 1
	Else
		$isdct = 0
	 EndIf


;Boucle Principale _Extended()
		_extended()

	Func __arraydisplay(Const ByRef $avarray, $stitle = "Array: ListView Display", $iitemlimit = -1, $itranspose = 0, $sseparator = "", $sreplace = "|", $sheader = "")
		If NOT IsArray($avarray) Then Return SetError(1, 0, 0)
		Local $idimension = UBound($avarray, 0), $iubound = UBound($avarray, 1) - 1, $isubmax = UBound($avarray, 2) - 1
		If $idimension > 2 Then Return SetError(2, 0, 0)
		If $sseparator = "" Then $sseparator = Chr(124)
		If _arraysearch($avarray, $sseparator, 0, 0, 0, 1) <> -1 Then
			For $x = 1 To 255
				If $x >= 32 AND $x <= 127 Then ContinueLoop
				Local $sfind = _arraysearch($avarray, Chr($x), 0, 0, 0, 1)
				If $sfind = -1 Then
					$sseparator = Chr($x)
					ExitLoop
				EndIf
			Next
		EndIf
		Local $vtmp, $ibuffer = 64
		Local $icollimit = 250
		Local $ioneventmode = Opt("GUIOnEventMode", 0), $sdataseparatorchar = Opt("GUIDataSeparatorChar", $sseparator)
		If $isubmax < 0 Then $isubmax = 0
		If $itranspose Then
			$vtmp = $iubound
			$iubound = $isubmax
			$isubmax = $vtmp
		EndIf
		If $isubmax > $icollimit Then $isubmax = $icollimit
		If $iitemlimit < 1 Then $iitemlimit = $iubound
		If $iubound > $iitemlimit Then $iubound = $iitemlimit
		If $sheader = "" Then
			$sheader = "Row  "
			For $i = 0 To $isubmax
				$sheader &= $sseparator & "Col " & $i
			Next
		EndIf
		Local $avarraytext[$iubound + 1]
		For $i = 0 To $iubound
			$avarraytext[$i] = "[" & $i & "]"
			For $j = 0 To $isubmax
				If $idimension = 1 Then
					If $itranspose Then
						$vtmp = $avarray[$j]
					Else
						$vtmp = $avarray[$i]
					EndIf
				Else
					If $itranspose Then
						$vtmp = $avarray[$j][$i]
					Else
						$vtmp = $avarray[$i][$j]
					EndIf
				EndIf
				$vtmp = StringReplace($vtmp, $sseparator, $sreplace, 0, 1)
				$avarraytext[$i] &= $sseparator & $vtmp
				$vtmp = StringLen($vtmp)
				If $vtmp > $ibuffer Then $ibuffer = $vtmp
			Next
		Next
		$ibuffer += 1
		Local Const $_arrayconstant_gui_dockborders = 102
		Local Const $_arrayconstant_gui_dockbottom = 64
		Local Const $_arrayconstant_gui_dockheight = 512
		Local Const $_arrayconstant_gui_dockleft = 2
		Local Const $_arrayconstant_gui_dockright = 4
		Local Const $_arrayconstant_gui_event_close = -3
		Local Const $_arrayconstant_lvif_param = 4
		Local Const $_arrayconstant_lvif_text = 1
		Local Const $_arrayconstant_lvm_getcolumnwidth = (4096 + 29)
		Local Const $_arrayconstant_lvm_getitemcount = (4096 + 4)
		Local Const $_arrayconstant_lvm_getitemstate = (4096 + 44)
		Local Const $_arrayconstant_lvm_insertitemw = (4096 + 77)
		Local Const $_arrayconstant_lvm_setextendedlistviewstyle = (4096 + 54)
		Local Const $_arrayconstant_lvm_setitemw = (4096 + 76)
		Local Const $_arrayconstant_lvs_ex_fullrowselect = 32
		Local Const $_arrayconstant_lvs_ex_gridlines = 1
		Local Const $_arrayconstant_lvs_showselalways = 8
		Local Const $_arrayconstant_ws_ex_clientedge = 512
		Local Const $_arrayconstant_ws_maximizebox = 65536
		Local Const $_arrayconstant_ws_minimizebox = 131072
		Local Const $_arrayconstant_ws_sizebox = 262144
		Local Const $_arrayconstant_taglvitem = "int Mask;int Item;int SubItem;int State;int StateMask;ptr Text;int TextMax;int Image;int Param;int Indent;int GroupID;int Columns;ptr pColumns"
		Local $iaddmask = BitOR($_arrayconstant_lvif_text, $_arrayconstant_lvif_param)
		Local $tbuffer = DllStructCreate("wchar Text[" & $ibuffer & "]"), $pbuffer = DllStructGetPtr($tbuffer)
		Local $titem = DllStructCreate($_arrayconstant_taglvitem), $pitem = DllStructGetPtr($titem)
		DllStructSetData($titem, "Param", 0)
		DllStructSetData($titem, "Text", $pbuffer)
		DllStructSetData($titem, "TextMax", $ibuffer)
		Local $iwidth = 1572, $iheight = 680
		Local $hguiv3 = GUICreate($stitle, $iwidth, $iheight, Default, Default, BitOR($_arrayconstant_ws_sizebox, $_arrayconstant_ws_minimizebox, $_arrayconstant_ws_maximizebox))
		Local $aiguisize = WinGetClientSize($hguiv3)
		Local $hlistview = GUICtrlCreateListView($sheader, 0, 0, $aiguisize[0], $aiguisize[1] - 26, $_arrayconstant_lvs_showselalways)
		Local $hcopy = GUICtrlCreateButton("Copie Selection", 3, $aiguisize[1] - 23, $aiguisize[0] - 6, 20)
		GUICtrlSetResizing($hlistview, $_arrayconstant_gui_dockborders)
		GUICtrlSetResizing($hcopy, $_arrayconstant_gui_dockleft + $_arrayconstant_gui_dockright + $_arrayconstant_gui_dockbottom + $_arrayconstant_gui_dockheight)
		GUICtrlSendMsg($hlistview, $_arrayconstant_lvm_setextendedlistviewstyle, $_arrayconstant_lvs_ex_gridlines, $_arrayconstant_lvs_ex_gridlines)
		GUICtrlSendMsg($hlistview, $_arrayconstant_lvm_setextendedlistviewstyle, $_arrayconstant_lvs_ex_fullrowselect, $_arrayconstant_lvs_ex_fullrowselect)
		GUICtrlSendMsg($hlistview, $_arrayconstant_lvm_setextendedlistviewstyle, $_arrayconstant_ws_ex_clientedge, $_arrayconstant_ws_ex_clientedge)
		Local $aitem
		For $i = 0 To $iubound
			If GUICtrlCreateListViewItem($avarraytext[$i], $hlistview) = 0 Then
				$aitem = StringSplit($avarraytext[$i], $sseparator)
				DllStructSetData($tbuffer, "Text", $aitem[1])
				DllStructSetData($titem, "Item", $i)
				DllStructSetData($titem, "SubItem", 0)
				DllStructSetData($titem, "Mask", $iaddmask)
				GUICtrlSendMsg($hlistview, $_arrayconstant_lvm_insertitemw, 0, $pitem)
				DllStructSetData($titem, "Mask", $_arrayconstant_lvif_text)
				For $j = 2 To $aitem[0]
					DllStructSetData($tbuffer, "Text", $aitem[$j])
					DllStructSetData($titem, "SubItem", $j - 1)
					GUICtrlSendMsg($hlistview, $_arrayconstant_lvm_setitemw, 0, $pitem)
				Next
			EndIf
		Next
		GUICtrlSendMsg($hlistview, 4126, 1, -1)
		If $idimension = 2 Then
			For $i = 2 To UBound($avarray, 2)
				GUICtrlSendMsg($hlistview, 4126, $i, -1)
			Next
		EndIf
		$iwidth = 0
		For $i = 0 To $isubmax + 1
			$iwidth += GUICtrlSendMsg($hlistview, $_arrayconstant_lvm_getcolumnwidth, $i, 0)
		Next
		If $iwidth < 250 Then $iwidth = 230
		$iwidth += 40
		If $iwidth > @DesktopWidth Then $iwidth = @DesktopWidth - 50 ;100
		WinMove($hguiv3, "", (@DesktopWidth - $iwidth) / 2, Default, $iwidth)
		GUISetState(@SW_SHOW, $hguiv3)
		While 1
			Switch GUIGetMsg()
				Case $_arrayconstant_gui_event_close
					ExitLoop
				Case $hcopy
					Local $sclip = ""
					Local $aicuritems[1] = [0]
					For $i = 0 To GUICtrlSendMsg($hlistview, $_arrayconstant_lvm_getitemcount, 0, 0)
						If GUICtrlSendMsg($hlistview, $_arrayconstant_lvm_getitemstate, $i, 2) Then
							$aicuritems[0] += 1
							ReDim $aicuritems[$aicuritems[0] + 1]
							$aicuritems[$aicuritems[0]] = $i
						EndIf
					Next
					If NOT $aicuritems[0] Then
						For $sitem In $avarraytext
							$sclip &= $sitem & @CRLF
						Next
					Else
						For $i = 1 To UBound($aicuritems) - 1
							$sclip &= $avarraytext[$aicuritems[$i]] & @CRLF
						Next
					EndIf
					ClipPut($sclip)
			EndSwitch
		WEnd
		GUIDelete($hguiv3)
		Opt("GUIOnEventMode", $ioneventmode)
		Opt("GUIDataSeparatorChar", $sdataseparatorchar)
		Return 1
	EndFunc

#EndRegion n1

#Region _Extended() ;main routine

	Func _extended()
	   while 1
	    Global $defautdc = $defautdcinit
		Global $allou = 0

		ToolTip("",5,5,"")

		  #Region GUI
			Global $hguiP = GUICreate("    [ AD Tools " & $ADTVersion & " ]---[ La POSTE/DSEM/EAPI69/NR ]---[ màj. " & $lastdateupdate & " ]         domaine: [   " & $isdct2 & "   ]" , 1240, 595, 50, 50, $ws_sysmenu)
			GUISetBkColor(14745599)
			$pic1 = GUICtrlCreatePic("",15,80,320,400)
			Global Const $Bmp_Logo = _GDIPlus_BitmapCreateFromMemory(_image(), True)
			_WinAPI_DeleteObject(GUICtrlSendMsg($pic1, $STM_SETIMAGE, $IMAGE_BITMAP, $Bmp_Logo))
			;Global $aff = GUICtrlCreateEdit("", 350, 5, 880, 555,$ES_READONLY , 512)
			Global $aff = _GUICtrlRichEdit_Create($hGuiP, "", 350, 5, 880, 562, BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL, $ES_READONLY ) )
				_ad_open()

			If $isDomainOK=1 Then
			_GUICtrlRichEdit_SetCharColor($aff, 0x00ff0000)      ;   set color blue (BBGGRR format)
			;_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
            _GUICtrlRichEdit_SetFont($aff, 10)                   ;   set size
            _GUICtrlRichEdit_SetCharAttributes($aff, '+bo')      ;   set bold
			_GUICtrlRichEdit_AppendText($aff, "> link ready !  ( connecté au Domaine:  "  & $isdct2 & "  )" & @CRLF)
		 Else
			_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
            _GUICtrlRichEdit_SetFont($aff, 10)                   ;   set size
            _GUICtrlRichEdit_SetCharAttributes($aff, '+bo')      ;   set bold
			_GUICtrlRichEdit_AppendText($aff, "> link not ready !  ( Etablissez une liaison reseau puis relancez AD Tools )" & @CRLF)
			EndIf

			Global $idabout = GUICtrlCreateButton(" ? ", 1, 1, 35, 20)
			Global $buttonlocalgroup = GUICtrlCreateButton(" Groupes Locaux", 263, 1, 85, 20)
			GUICtrlSetTip($buttonlocalgroup, "Gere les groupes locaux d'un Ordinateur" & @CRLF & "comparer,lister, [+] ou [-] groupes locaux...", "", 0, 1)
			Global $buttoncomputer = GUICtrlCreateButton("Ordinateur", 105, 1, 60, 20)
			GUICtrlSetTip($buttoncomputer, "[+] ou [-] 'Groupe(s)' pour Ordinateur(s)", "", 0, 1)
			Global $buttonbitlocker = GUICtrlCreateButton("Bitlocker", 41, 1, 60, 20)
			GUICtrlSetTip($buttonbitlocker, "'Bitlocker' pour Ordinateur(s), non sur MBAM, mais sur un AD !", "", 0, 1)
			;
			Global $buttonidrh = GUICtrlCreateButton("Utilisateur (Idrh)", 175, 1, 80, 20)
			GUICtrlSetTip($buttonidrh, "Gestion des comptes utilisateurs (Idrh)", "", 0, 1)
			;
			GUISetState(@SW_SHOW, $hgui)
			Send("{TAB}")
		 #EndRegion GUI

Local $idmsg
		While 1
		   	;  tooltip("Boucle Principale",5,5,"")
			$idmsg = GUIGetMsg()
			Switch $idmsg
				Case $buttonlocalgroup
					If _ispressed("10", $hdll) Then
					Else
						readlocalgroup2()
					EndIf
					GUICtrlDelete($hguilg)
				Case $buttonbitlocker
					$t = MsgBox(4, "Clé Bitlocker (AD) ?", "[Oui]: clé Bitlocker ordinateur" & @CRLF & @CRLF & "[Non]: ne fait rien")
					If $t = 6 Then
						bitlocker()
					EndIf
				Case $buttoncomputer
					$t = MsgBox(4, "Computer ?", "[Oui]: [+/-]  'Groupe(s)' => Computer(s)" & @CRLF & @CRLF & "[Non]: ne fait rien")
					If $t = 6 Then
						computergroup()
					 EndIf
			    Case $buttonidrh
					   _Func_Idrh()
				Case $gui_event_close
					 _terminate()

				Case $idabout
					about()
					$t = MsgBox(0, "A propos d' 'AD Tools'", "AD Tools v6 - Nicolas RISTOVSKI" & @CRLF & @CRLF & $lastdatecompile)
					_Exit()
			EndSwitch
		WEnd

wend
EndFunc

#EndRegion

Func _Func_Idrh()
;$historik=""
$cyclesIDRG=$cyclesIDRG+1
if $cyclesIDRG>15 Then
   $cyclesIDRG=0
   _GUICtrlRichEdit_SetText($aff, "" ) 														; _GUICtrlRichEdit_AppendText($aff, "" )
   ;_GUICtrlRichEdit_AppendText($aff, "Erase history: Ready !" & @crlf & @CRLF);, 1)
EndIf

$defautdc = $defautdcinit

 Local $aPos = WinGetPos($hguiP)
    ; Display the array values returned by WinGetPos.
    ;MsgBox($MB_SYSTEMMODAL, "", "X-Pos: " & $aPos[0] & @CRLF & "Y-Pos: " & $aPos[1] & @CRLF )

   #Region GUI
			Global $hgui = GUICreate("    -[ Utilisateur ]--[ IDRH ]-	",355, 580, $aPos[0], $aPos[1]+7, $ws_sysmenu)
			GUISetBkColor(14745599)
			;Global $aff = GUICtrlCreateEdit("", 350, 5, 880, 555, -1, 512)
			;_GUICtrlRichEdit_AppendText($aff, $historik, 1)
			_ad_open()
			$buttonidrh1 = GUICtrlCreateInput("", 50, 50, 85, 35)
			GUICtrlSetFont(-1, 14, 400, 0, "MS Sans Serif")
			GUICtrlSetBkColor(-1, 6908265)
			GUICtrlSetTip($buttonidrh1, "Idrh Source", "", 0, 1)
			$labelidrh1 = GUICtrlCreateLabel("Srce", 10, 50, 40, 15)
			$idcheckboxcommonname = GUICtrlCreateCheckbox("", 25, 16, 14, 14)
			$idcheckboxcommonnamelabel = GUICtrlCreateLabel("[ANR] cherche 'Srce' par Nom", 15, 30, 170, 15)
			GUICtrlSetTip($idcheckboxcommonname, "Si coché, recherche par le nom (ANR) dans Idrh 'Srce'" & @CRLF & "", "", 0, 1)
			$buttonidrh2 = GUICtrlCreateInput("", 50, 100, 85, 35)
			GUICtrlSetFont(-1, 14, 400, 0, "MS Sans Serif")
			GUICtrlSetBkColor(-1, 6908265)
			GUICtrlSetTip($buttonidrh2, "Idrh destination", "", 0, 1)
			$labelidrh2 = GUICtrlCreateLabel("Dest", 10, 100, 40, 15)
			$buttonidgroup = GUICtrlCreateInput("", 22, 300, 125, 35)
			GUICtrlSetFont(-1, 14, 400, 0, "MS Sans Serif")
			GUICtrlSetBkColor(-1, 6908265)
			GUICtrlSetTip($buttonidgroup, "Groupe ou Objet ? : cocher la case [Source] ! " & @CRLF & "separateur [;]  " & " (Idrh, SID key, nom ... ex: nicolas ;san ;)" & @CRLF & @CRLF & "touche [Shift] pressée => champs formatés: 'SID key'|'objectname'", "", 0, 1)
			$labelidgroup = GUICtrlCreateLabel("lister clé SID ?", 35, 280, 100, 15)
			$buttonidgroup2 = GUICtrlCreateInput("", 22, 400, 125, 35)
			GUICtrlSetFont(-1, 14, 400, 0, "MS Sans Serif")
			GUICtrlSetBkColor(-1, 6908265)
			GUICtrlSetTip($buttonidgroup2, "[+] multiples Groupes [;] à l' Idrh 'Source' " & @CRLF & "separateur [;]  " & "" & @CRLF & @CRLF & "touche [Shift] pressée : [-] groupes", "", 0, 1)
			$labelidgroup2 = GUICtrlCreateLabel("[+] groupes [;]", 20, 380, 120, 20)
			$labelcreategroup = GUICtrlCreateLabel("Gr?", 115, 380, 20, 14)
			$idcheckboxcreategroups = GUICtrlCreateRadio("", 135, 380, 14, 14)
			GUICtrlSetTip($idcheckboxcreategroups, "membres d'un Groupe ? , recuperer la liste des membres" & @CRLF & @CRLF & "touche [SHIFT]: creation groupe" & @CRLF & " Si separator [;] creation en Masse des groupes", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labeldrive = GUICtrlCreateLabel("cpy drives ?", 190, 20, 90, 12)
			$idcheckboxdrives = GUICtrlCreateCheckbox("", 174, 20, 14, 15)
			GUICtrlSetTip($idcheckboxdrives, "copier Lecteurs reseaux depuis Idrh 'Srce' vers Idrh 'Dest'", "", 0, 1)
			If $isdct = 1 Then
				GUICtrlSetState(-1, $gui_checked)
			Else
				GUICtrlSetState(-1, $gui_unchecked)
			EndIf
			$labelscan = GUICtrlCreateLabel("cpy Mes scan ?", 190, 50, 90, 12)
			$idcheckscancpy = GUICtrlCreateCheckbox("", 174, 50, 14, 14)
			GUICtrlSetTip($idcheckscancpy, "copier le 'comment' Mes scan depuis Idrg 'Srce' vers  Idrh 'Dest'", "", 0, 1)
			If $isdct = 1 Then
				GUICtrlSetState(-1, $gui_checked)
			Else
				GUICtrlSetState(-1, $gui_unchecked)
			EndIf
			$labelgroups = GUICtrlCreateLabel("cpy Groupes ?", 190, 80, 90, 12)
			$idcheckboxgroups = GUICtrlCreateCheckbox("", 174, 80, 14, 14)
			GUICtrlSetTip($idcheckboxgroups, "copier tous les groupes depuis Idrh 'Srce' vers Idrh 'Dest'", "", 0, 1)
			If $isdct = 1 Then
				GUICtrlSetState(-1, $gui_unchecked)
			Else
				GUICtrlSetState(-1, $gui_checked)
			EndIf
			$labelreinitpwd = GUICtrlCreateLabel("Pwd Reset 'Srce' ?", 190, 110, 90, 12)
			$idcheckboxpwdreset = GUICtrlCreateRadio("", 174, 110, 14, 14)
			GUICtrlSetTip($idcheckboxpwdreset, "Reinit du mot de passe Idrh 'Srce'", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelairwatch = GUICtrlCreateLabel("Airwatch 'Srce' ?", 190, 140, 90, 12)
			$idcheckboxairwatch = GUICtrlCreateRadio("", 174, 140, 14, 14)
			GUICtrlSetTip($idcheckboxairwatch, "Rajoute groupe Airwatch, et son @mail à l'Idrh 'Srce'", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelairwatchbpe = GUICtrlCreateLabel("AW ?", 293, 140, Default, 12)
			$idcheckboxbpe = GUICtrlCreateCheckbox("", 280, 140, 14, 14)
			GUICtrlSetTip($idcheckboxbpe, "AW ? = selectionne un autre groupe Airwatch (ex: BPE, Ma French Banque...)", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelairgroup = GUICtrlCreateLabel("[+] groupe 'Srce' ?", 190, 170, 110, 12)
			$idcheckboxgroup = GUICtrlCreateRadio("", 174, 170, 14, 14)
			GUICtrlSetTip($idcheckboxgroup, "Rajoute un ou plusieurs groupes à l'Idrh 'Srce'", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelmultiplesrce = GUICtrlCreateLabel("[a;b]", 148, 155, 38, 16)
			$multiplesrce = GUICtrlCreateRadio("", 154, 170, 14, 14)
			GUICtrlSetTip($multiplesrce, "multiple Idrh 'Source' pour [+] groupes   -  separateur [;]" & @CRLF & @CRLF & "[Shift] key pressée avec [OK ] => Pour retirer les groupes..." & @CRLF & @CRLF & "(Idrh, name, SID key)", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelrole = GUICtrlCreateLabel("[+] Role 'Srce' ?", 190, 200, 90, 12)
			$idcheckboxrole = GUICtrlCreateRadio("", 174, 200, 14, 14)
			GUICtrlSetTip($idcheckboxrole, "Rajoute un role à l'Idrh 'Srce'", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labellbpai = GUICtrlCreateLabel("LBPAI/Iliade ?", 40, 200, 90, 12)
			$idcheckboxlbpai = GUICtrlCreateRadio("", 24, 200, 14, 14)
			If $isdct = 1 Then
				GUICtrlSetTip($idcheckboxlbpai, "[+] [Categories groupes] pour: LBPAI/Iliade -  [SHIFT] key pour retirer [-] " & @CRLF & "serveur: 752474SN-FI01\_Iliade$", "", 0, 1)
			Else
				GUICtrlSetTip($idcheckboxlbpai, "[+] [Categories groupes] pour: LBPAI/Iliade -  [SHIFT] key pour retirer [-] ", "", 0, 1)
			EndIf
			GUICtrlSetState(-1, $gui_unchecked)
			$labeldirective = GUICtrlCreateLabel("Directive 'Srce' ?", 40, 230, 90, 12)
			$idcheckboxdirective = GUICtrlCreateRadio("", 24, 230, 14, 14)
			GUICtrlSetTip($idcheckboxdirective, "Appliquer une Directive Metier à l'Idrh 'Srce', l'ancienne Directive est automatiquement retirée si la nouvelle est appliquée. Si Directive non appliquée, un retour en arriere est effectué (rollback)" & @crlf & " note: la liste des groupes de l'ancienne Directive retirée est listée sous la ligne : Groupes Directive Obsolete ", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)

			$labeldsiba = GUICtrlCreateLabel("DSIBA 'Srce' ?", 40, 255, 90, 12)
			$idcheckboxdirectiveDSIBA = GUICtrlCreateRadio("", 24, 255, 14, 14)
			GUICtrlSetTip($idcheckboxdirectiveDSIBA, "Appliquer une Directive Metier DSIBA à l'Idrh 'Srce', l'ancienne Directive est automatiquement retirée si la nouvelle est appliquée, sinon un retour en arriere est effectué" & @crlf & " note: la liste des groupes de l'ancienne Directive retirée est listée sous la ligne : Groupes Directive Obsolete ", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)

			$labelairgroupremove = GUICtrlCreateLabel("[-] groupe 'Srce' ?", 40, 170, 110, 12)
			$idcheckboxgroupremove = GUICtrlCreateRadio("", 24, 170, 14, 14)
			GUICtrlSetTip($idcheckboxgroupremove, "Retirer des groupes ou roles à l'Idrh 'Srce'", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelprolonge = GUICtrlCreateLabel("Date 'Srce' ?", 190, 230, 100, 12)
			$idcheckboxprolonge = GUICtrlCreateRadio("", 174, 230, 14, 14)
			GUICtrlSetTip($idcheckboxprolonge, "Prolonge la date de l'Idrh 'Srce'" & @CRLF & @CRLF & "[shift] key pressée avec [OK]: enlève l'expiration pour un compte en 'x' !", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelcreate = GUICtrlCreateLabel("Creer 'Srce' ?", 190, 260, 90, 12)
			$idcheckboxcreate = GUICtrlCreateRadio("", 174, 260, 14, 14)
			GUICtrlSetTip($idcheckboxcreate, "Creer nouveau compte utilisateur 'Srce' " & @CRLF & @CRLF & "[SHIFT] key pressée avec [OK]:  Permet de renommer le compte Source uniquement !", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelmove = GUICtrlCreateLabel("Deplace 'Srce' ?", 190, 290, 110, 12)
			$idcheckboxmove = GUICtrlCreateRadio("", 174, 290, 14, 14)
			GUICtrlSetTip($idcheckboxmove, "Deplacer un compte 'Srce' dans une autre OU " & @CRLF & @CRLF & "Si pas d'Idrh 'Srce' renseigné alors cherche un fichier texte à importer pour deplacer en masse" & @CRLF & " ex fichier: pabc123|BPLY" & @CRLF & @CRLF & "traiter en Masse LBPAI/Iliade" & @CRLF & " ex fichier LBPAI: pabc123|MIPS|Standard;commun full users;Chargé de Développement" & @CRLF & "	lbpai? or iliade? à saisir dans Idrh 'Srce' et cocher la case [ANR] pour voir liste des categories LBPAI", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labeldrive = GUICtrlCreateLabel("[+] Drive 'Srce' ?", 190, 320, 90, 12)
			$idcheckboxdrive = GUICtrlCreateRadio("", 174, 320, 14, 14)
			GUICtrlSetTip($idcheckboxdrive, "Rajoute un ou plusieurs lecteurs à l'Idrh 'Srce' " & @CRLF & @CRLF & "multiple idrh 'Srce', separateur [;] : lister tous les comptes 'srce' avec leurs lecteurs DCT", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labeldriveremove = GUICtrlCreateLabel("[-] Drive 'Srce' ?", 190, 350, 110, 12)
			$idcheckboxdriveremove = GUICtrlCreateRadio("", 174, 350, 14, 14)
			GUICtrlSetTip($idcheckboxdriveremove, "Retirer un lecteur à l'Idrh 'Srce' ", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labeldesactivesrce = GUICtrlCreateLabel("Desactive 'Srce' ?", 190, 380, 110, 12)
			$idcheckboxdesactivesrce = GUICtrlCreateRadio("", 174, 380, 14, 14)
			GUICtrlSetTip($idcheckboxdesactivesrce, "Desactiver l'Idrh 'Srce'" & @CRLF & @CRLF & "[SHIFT] pressée avec [OK]:  Pour supprimer le compte Idrh 'Srce'", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labeldescription = GUICtrlCreateLabel("Description 'Srce' ?", 190, 410, 110, 12)
			$idcheckboxdescription = GUICtrlCreateRadio("", 174, 410, 14, 14)
			GUICtrlSetTip($idcheckboxdescription, "Changer la description Idrh 'Srce' ", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelscan = GUICtrlCreateLabel("Comment 'Srce' ?", 190, 440, 110, 12)
			$idcheckboxscan = GUICtrlCreateRadio("", 174, 440, 14, 14)
			GUICtrlSetTip($idcheckboxscan, "change le 'Comment' champ Mes scan pour l'Idrh 'Srce'", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			Global $idok2 = GUICtrlCreateButton("OK", 55, 480, 225, 65)
			GUICtrlSetTip($idok2, "Presser OK pour valider (certaines options fonctionnent avec la touche [Shift] ", "", 0, 1)
			GUISetState(@SW_SHOW, $hgui)
			GUICtrlSetData($buttonidrh1, "")
			GUICtrlSetBkColor($buttonidrh1, 6908265)
			GUICtrlSetData($buttonidrh2, "")
			GUICtrlSetBkColor($buttonidrh2, 6908265)
			$result = ""
			$result2 = ""
			$idrh1 = ""
			$idrh2 = ""
			$outputidrh1 = ""
			Send("{TAB}")
		#EndRegion GUI

	    Global $idrh1 = "", $idrh2 = "", $idgroup = "", $idgroup2 = ""
		Global $actiongroups = ""
		Global $actiongroupsar = ""
		Global $driveismanual = ""
		Global $result, $result2, $outputidrh1
	    ;_GUICtrlRichEdit_AppendText($aff, @crlf & "---[ IDRH ]---  Ready !" & @CRLF);, 1)

				While 1
		   	;  tooltip("Boucle Idrh",5,5,"")
			$idmsg2 = GUIGetMsg()
			Switch $idmsg2
			    Case $buttonidgroup
					$idgroup = GUICtrlRead($buttonidgroup)
					GUICtrlSetBkColor($buttonidgroup, 16776960)
				Case $buttonidgroup2
					$idgroup2 = GUICtrlRead($buttonidgroup2)
					GUICtrlSetBkColor($buttonidgroup2, 16776960)
				Case $buttonidrh1
					$idrh1 = GUICtrlRead($buttonidrh1)
					GUICtrlSetBkColor($buttonidrh1, 16776960)
					$idrh1 = StringStripWS($idrh1, 8)
				Case $buttonidrh2
					$idrh2 = GUICtrlRead($buttonidrh2)
					GUICtrlSetBkColor($buttonidrh2, 16776960)
					$idrh2 = StringStripWS($idrh2, 8)

			    Case $idok2
				  ;$idrh1=GUICtrlRead($buttonidrh1)
				  ;$idrh2=GUICtrlRead($buttonidrh2)
				  ;$idgroup = GUICtrlRead($buttonidgroup)
				  ;$idgroup2 = GUICtrlRead($buttonidgroup2)
					$idcheckboxcreate = BitAND(GUICtrlRead($idcheckboxcreate), $gui_checked)
					$idcheckboxprolonge = BitAND(GUICtrlRead($idcheckboxprolonge), $gui_checked)
					$idcheckboxmove = BitAND(GUICtrlRead($idcheckboxmove), $gui_checked)
					$idcheckboxdrive = BitAND(GUICtrlRead($idcheckboxdrive), $gui_checked)
					$userexist = _ad_objectexists($idrh1)
					$idcheckboxdirective = BitAND(GUICtrlRead($idcheckboxdirective), $gui_checked)
					$idcheckboxdirectiveDSIBA  = BitAND(GUICtrlRead($idcheckboxdirectiveDSIBA), $gui_checked)
					$idcheckboxgroupremove = BitAND(GUICtrlRead($idcheckboxgroupremove), $gui_checked)
					$idcheckboxdriveremove = BitAND(GUICtrlRead($idcheckboxdriveremove), $gui_checked)
					$idcheckboxdesactivesrce = BitAND(GUICtrlRead($idcheckboxdesactivesrce), $gui_checked)
					$idcheckboxdescription = BitAND(GUICtrlRead($idcheckboxdescription), $gui_checked)
					$idcheckboxscan = BitAND(GUICtrlRead($idcheckboxscan), $gui_checked)
					$idcheckscancpy = BitAND(GUICtrlRead($idcheckscancpy), $gui_checked)
					$multiplesrce = BitAND(GUICtrlRead($multiplesrce), $gui_checked)
					$idcheckboxcommonname = BitAND(GUICtrlRead($idcheckboxcommonname), $gui_checked)
					$idcheckboxcreategroups = BitAND(GUICtrlRead($idcheckboxcreategroups), $gui_checked)
					$bpe = BitAND(GUICtrlRead($idcheckboxbpe), $gui_checked)
					$idcheckboxlbpai = BitAND(GUICtrlRead($idcheckboxlbpai), $gui_checked)
					$listmanualdrives = ""
					Global $scomboreaddirectives = ""
					;_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "IDRH: [OK]" & @CRLF);, 1)
					If $idcheckboxmove = 1 AND $isdct = 1 AND StringLen($idrh1) = 0 Then
						massive_move()
						_GUICtrlRichEdit_SetCharColor($aff, 0x00ff0000)      ;   set color blue (BBGGRR format)
						_GUICtrlRichEdit_SetFont($aff, 10)                   ;   set size
						_GUICtrlRichEdit_SetCharAttributes($aff, '+bo')      ;   set bold
						_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Massive Move Users - (deplacer users)" & @CRLF);, 1)
						Guidelete($hGUI)
						Return 0
					EndIf
					If $idcheckboxprolonge = 1 AND StringLen($idrh1) = 0 Then
						massive_date()
						_GUICtrlRichEdit_SetCharColor($aff, 0x00ff0000)      ;   set color blue (BBGGRR format)
						_GUICtrlRichEdit_SetFont($aff, 10)                   ;   set size
						_GUICtrlRichEdit_SetCharAttributes($aff, '+bo')      ;   set bold
						_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Massive Extend Date Users - (Prolongation comptes)" & @CRLF);, 1)
						Guidelete($hGUI)
						Return 0
					EndIf
					If $idcheckboxdrive = 1 AND $isdct = 1 AND StringInStr($idrh1, ";") Then
						massive_drive()
						_GUICtrlRichEdit_SetCharColor($aff, 0x00ff0000)      ;   set color blue (BBGGRR format)
						_GUICtrlRichEdit_SetFont($aff, 10)                   ;   set size
						_GUICtrlRichEdit_SetCharAttributes($aff, '+bo')      ;   set bold
						_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Massive Users Drives Export - (Exporter la liste des lecteurs reseux de chaque user)" & @CRLF);, 1)
						Guidelete($hGUI)
						Return 0
					EndIf
					If StringCompare($idrh1, $idrh2) = 0 AND $idrh1 <> "" Then
						MsgBox(0, "Info !", "User 'Srce' et  User 'Dest' sont identiques !" & @CRLF & "Abandon !", 7)
						_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
						_GUICtrlRichEdit_SetFont($aff, 10)                   ;   set size
						_GUICtrlRichEdit_SetCharAttributes($aff, '+bo')      ;   set bold
						_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "User 'Srce' et User 'Dest' sont identiques : Abandon ! "  & @CRLF);, 1)
						Guidelete($hGUI)
						Return 0
					 EndIf

					#Region Select Case

						Select
							Case $idcheckboxcreategroups = 1
								$defautdc = $defautdcinit
								If _ispressed("10", $hdll) Then
									$idgroup2 = StringStripWS($idgroup2, 3)
									If $idgroup2 <> "" AND NOT StringInStr($idgroup2, ";") Then
										If StringLen($idgroup2) < 7 Then
											ToolTip("", 5, 5)
											MsgBox(0, "Warning !", $idgroup2 & " , longeur <7 !" & @CRLF & "Abandon ...", 7)
											 _GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
											 _GUICtrlRichEdit_SetFont($aff, 10)                   ;   set size
											 _GUICtrlRichEdit_SetCharAttributes($aff, '+bo')      ;   set bold
											 _GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Nom du groupe doit etre > ou egal à 7 chr$" & @CRLF);, 1)
											 Guidelete($hGUI)
											Return 0
										Else
										EndIf
										Global $unite = ""
										ToolTip("Création d'un nouveau groupe...", 5, 5)
										If $isdct = 1 Then
											$unite = InputBox("default OU ? 2 or 4 chr$", "ex: MILY, MITE, BPLY, GAUB, MI ... virtuos" & @CRLF & "?? : scan all OUs...", "MI")
											If StringLen($unite) < 2 OR StringLen($unite) > 4 AND NOT ($unite = "virtuos" OR $unite = "??") Then
												ToolTip("", 5, 5)
												ExitLoop
											EndIf
										Else
											$unite = "??"
											ToolTip("", 5, 5)
										EndIf
										If $unite = "??" Then
											$allou = 1
											treeview_affiche()
											treeselect()
										Else
											$allou = 0
										EndIf
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										ToolTip("", 5, 5)
										If $allou = 1 Then
											$unite = ""
											$allou = 0
											Global $sou = $sout
										Else
											Global $sou = "OU=Groupes,OU=" & $unite & ",OU=" & StringMid($unite, 1, 2) & $defautdc
										EndIf
										If $unite = "MI" Then
											Global $sou = "OU=Groupes,OU=_Support Transverse,OU=MI" & $defautdc
										EndIf
										If $unite = "virtuos" Then
											Global $sou = "OU=FunctionalGroups,OU=_Flexible Workspace Commun" & $defautdc
										EndIf
										Global $ivalue = _ad_creategroup($sou, $idgroup2)
										If $ivalue = 1 Then
											Global $groupdesc = InputBox("Definir 'Description' groupe !", $idgroup2 & @CRLF & "Si necessaire, facultatif...", "")
											$ivalue = _ad_modifyattribute($idgroup2, "description", $groupdesc, 2)
											MsgBox(64, "Info !", "Groupe '" & $idgroup2 & "' crée" & @CRLF & $sou & @CRLF & @CRLF & "description est:" & @CRLF & "'" & $groupdesc & "'")
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Groupe '" & $idgroup2 & "' crée dans l'OU: " & "  " & $sou & "    -   " & "description:" & "  " & "'" & $groupdesc & "'" & @CRLF);, 1)
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Group '" & $idgroup2 & "' crée dans l'OU: " & "  " & $sou & "    -   " & "description:" & "  " & "'" & $groupdesc & "'" & @CRLF
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Groupe '" & $idgroup2 & "' crée dans l'OU: " & "  " & $sou & "    -   " & "description:" & "  " & "'" & $groupdesc & "'" & @CRLF);, 1)
										ElseIf @error = 1 Then
											MsgBox(64, "Info !", "Groupe '" & $idgroup2 & "' existe déjà dans l'AD !")
										ElseIf @error = 2 Then
											MsgBox(64, "Warning !", "Cette OU= '" & $sou & "' n'existe pas dans l'AD !")
										Else
											MsgBox(64, "Warning !", "Return code '" & @error & "' de l'AD..." & @CRLF & " (acces refusé ? ou nom de groupe trop long ! (>64chr$) pour creer le groupe")
										EndIf
										ToolTip("", 5, 5)
										Guidelete($hGUI)
										Return 0
									ElseIf $idgroup2 <> "" AND StringInStr($idgroup2, ";") Then
										Global $unite = ""
										ToolTip("Creation des groupes en masse...", 5, 5)
										If $isdct = 1 Then
											$unite = InputBox("default OU ? 2 or 4 chr$", "ex: MILY, MITE, BPLY, GAUB, MI ... virtuos" & @CRLF & "?? : scan all OUs...", "MI")
											If StringLen($unite) < 2 OR StringLen($unite) > 4 AND NOT ($unite = "virtuos" OR $unite = "??") Then
												ToolTip("", 5, 5)
												ExitLoop
											EndIf
										Else
											$unite = "??"
											ToolTip("", 5, 5)
										EndIf
										If $unite = "??" Then
											$allou = 1
											treeview_affiche()
											treeselect()
										Else
											$allou = 0
										EndIf
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										ToolTip("", 5, 5)
										If $allou = 1 Then
											$unite = ""
											$allou = 0
											Global $sou = $sout
										Else
											Global $sou = "OU=Groupes,OU=" & $unite & ",OU=" & StringMid($unite, 1, 2) & $defautdc
										EndIf
										If $unite = "MI" Then
											Global $sou = "OU=Groupes,OU=_Support Transverse,OU=MI" & $defautdc
										EndIf
										If $unite = "virtuos" Then
											Global $sou = "OU=FunctionalGroups,OU=_Flexible Workspace Commun" & $defautdc
										EndIf
										Global $result = ""
										Global $result2 = ""
										$idgroup2 = StringSplit($idgroup2, ";")
										For $z = 1 To $idgroup2[0]
											If StringLen($idgroup2[$z]) < 7 Then
												ToolTip("", 5, 5)

												MsgBox(0, "Warning !", $idgroup[$z] & " , lenth<7" & @CRLF & "rien effectué... abandon !", 7)
												Guidelete($hGUI)
												Return 0
											Else
											EndIf
										Next
										For $z = 1 To $idgroup2[0]
											Global $ivalue = _ad_creategroup($sou, $idgroup2[$z])
											If $ivalue = 1 Then
												$result2 = $result2 & $idgroup2[$z] & ";"
												$groupdesc = "" & $idgroup2[$z]
												$ivalue = _ad_modifyattribute($idgroup2[$z], "description", $groupdesc, 2)
												$result = $result & "Groupe '" & $idgroup2[$z] & "' cree dans l'OU:  " & $sou & "	/ description:	" & "'" & $idgroup2 & "'" & @CRLF
											ElseIf @error = 1 Then
												$result = $result & "Groupe '" & $idgroup2[$z] & "' existe déjà dans l'AD !" & @CRLF
											ElseIf @error = 2 Then
												$result = $result & "Warning !  " & "Cette OU= '" & $sou & "' n'existe pas dans l'AD !" & @CRLF
											Else
												$result = $result & "Warning !  " & "Return code '" & @error & "' de l'AD... " & " (acces refusé ? ou nom de groupe trop long ! (>64chr$) pour creer le groupe: '" & $idgroup2[$z] & "'" & @CRLF
											EndIf
										Next
										ToolTip("", 5, 5)
										$result = "		Creation des nouveaux groupes en masse dans l'OU:  " & $sou & @CRLF & @CRLF & $result
										$result2 = StringTrimRight($result2, 1)
										$result = $result & @CRLF & @CRLF & $result2
										  _GUICtrlRichEdit_SetCharColor($aff, 0x00ff0000)      ;   set color blue (BBGGRR format)
										  _GUICtrlRichEdit_SetFont($aff, 10)                   ;   set size
										  _GUICtrlRichEdit_SetCharAttributes($aff, '+bo')      ;   set bold
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF)
										_GUICtrlRichEdit_AppendText($aff, $result & @CRLF);, 1)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $result & @CRLF
										Guidelete($hGUI)
										Return 0
									Else
										MsgBox(0, "Warning !", "Le nom du groupe à creer ne peut etre <vide> !", 7)
										ToolTip("", 5, 5)
										Guidelete($hGUI)
										Return 0
									EndIf
								EndIf
								If $idgroup2 <> "" AND NOT StringInStr($idgroup2, ";") Then
									$test = _ad_objectexists($idgroup2)
									If $test = 1 Then
										$win = _ad_getgroupmembers($idgroup2)
										$amembersaddsam = ""
										If IsArray($win) Then
											For $i = 1 To UBound($win) - 1
												$amembersaddsam = $amembersaddsam & _ad_fqdntosamaccountname($win[$i]) & ";"
											Next
										EndIf
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idgroup2 & @CRLF & $amembersaddsam & @CRLF);, 1)
										If $amembersaddsam <> "" Then
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idgroup2 & @CRLF & "Le groupe contient [ " & $win[0] & " ] membres." & @CRLF & $amembersaddsam & @CRLF & @CRLF
										Else
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idgroup2 & @CRLF & "Le groupe n'a aucun membre ! [ " & $win[0] & " ]." & @CRLF & $amembersaddsam & @CRLF & @CRLF
										EndIf
									Else
										$win = "Le Groupe:  " & $idgroup2 & "  n'existe pas !"
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idgroup2 & @CRLF & $win & @CRLF);, 1)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idgroup2 & @CRLF & $win & @CRLF
									EndIf
								EndIf
								Guidelete($hGUI)
								Return 0
						    Case $idcheckboxcommonname = 1
	  if $domainname="" Then
		   MsgBox(0,"Info !","Pas d'Active Directory trouvé ! impossible de scanner...  Etablissez une liaison reseau et relancez  'AD Tools' ..." & @crlf & $domainname,15)
		   ToolTip("",5,5,"")

		   Guidelete($hGUI)
		   Return 0
	  EndIf
								ToolTip("", 5, 5,"")
								$defautdc = $defautdcinit
								$filtreperso = 0
								If _ispressed("10", $hdll) Then
									$filtreperso = 1
								Else
									$filtreperso = 0
								EndIf
								If $idrh1 = "" AND $idgroup = "" AND NOT (StringRegExp($idgroup, "^S-\d-\d+-(\d+-){1,14}\d+$") OR StringRegExp($idgroup, "^S-\d-(\d+-){1,14}\d+$") OR StringInStr($idgroup, ";")) Then
									MsgBox(64, "Info !", "Si user Srce = 'disabled' ou 'locked', recherche des comptes 'désactivés' ou 'vérouillés' dans l'AD..." & @CRLF & @CRLF & "Et aussi... 'SID' (groupes,personnes)," & @CRLF & "'SIDh' cherche tous les SIDh (history) pour les categories (groupes,personnes...)" & @CRLF & "'lbpai?' ou 'iliade?' pour afficher les 'categories de groupes' pour exploiter via import fichier...  A utiliser avec 'Deplacer Srce' sans preciser d'idrh 'srce' ! (import fichier texte)")
									ToolTip("", 5, 5)
									_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Si 'Srce' = 'disabled' ou 'locked', rechercher des comptes 'désactivés' ou 'vérouillés' dans l'AD..." & @CRLF & @CRLF & "Et aussi... 'SID' (groupes,personnes)," & @CRLF & "'SIDh' recherche des SIDh (SID history) groupes,personnes..." & @CRLF & "'lbpai?' ou 'iliade?' : pour afficher les 'categories de groupes' pour exploiter via import fichier...  A utiliser avec 'Deplacer Srce' sans preciser d'idrh 'srce' ! (import fichier texte)" & @CRLF);, 1)
									   Guidelete($hGUI)
									   Return 0
								ElseIf $idgroup <> "" AND $idrh1 = "" AND NOT StringInStr($idgroup, ";") Then
									Global $displayname = ""
									If StringRegExp($idgroup, "\h") AND StringLen($idgroup) > 2 Then
										Global $displayname = $idgroup
										Global $sou = ""
										Global $sldapfilter = "(&(objectCategory=Person)(ANR=" & $idgroup & " ))"
										Global $sdatatoretrieve = "sAMAccountName"
										Global $auserids = _ad_getobjectsinou($sou, $sldapfilter, 2, $sdatatoretrieve, "displayName")
										If IsArray($auserids) Then
											$idgroup = $auserids[1]
										EndIf
									EndIf
									Global $testbinaryh2 = ""
									$testbinaryh = _ad_getobjectproperties($idgroup, "SIDhistory")
									If IsArray($testbinaryh) AND UBound($testbinaryh) >= 2 Then
										For $z = 1 To UBound($testbinaryh) - 1
											$testbinaryh2 = $testbinaryh2 & $testbinaryh[$z][1] & "|"
										Next
										$testbinaryh = $testbinaryh2
										$testbinaryh = StringTrimRight($testbinaryh, 1)
									Else
										$testbinaryh = ""
									EndIf
									$sidkeyg = _ad_getobjectproperties($idgroup, "ObjectSID")
									If IsArray($sidkeyg) Then
										$sidkeyg1 = _arraytostring($sidkeyg)
										$sidkeyg1 = StringSplit($sidkeyg1, "|")
										$sidkeyg1 = $sidkeyg1[3]
										If $testbinaryh = "" Then
											InputBox("SID key !", "SID key pour: " & $displayname & @CRLF & $idgroup & @CRLF & $testbinaryh, $sidkeyg1)
											ClipPut($sidkeyg1)
										Else
											$result2 = "SID key pour: " & $displayname & "  " & $idgroup & "     " & $sidkeyg1 & @CRLF & "SIDhistory : " & $testbinaryh & @CRLF
											ClipPut($result2)
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $result2 & @CRLF & @CRLF & "-- SID keys:" & @CRLF & $sidkeyg1 & @CRLF & "--" & @CRLF & @CRLF & "			=====	=====	=====	=====	=====			" & @CRLF & @CRLF);, 1)
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $result2 & @CRLF & @CRLF & "-- SID keys:" & @CRLF & $sidkeyg1 & @CRLF & "--" & @CRLF
											ClipPut($sidkeyg1)
										EndIf
										ToolTip("", 5, 5)
										Guidelete($hGUI)
										Return 0
									ElseIf StringInStr($idgroup, "-") AND StringLen($idgroup) >= 20 AND $idrh1 = "" Then
										$sidkeymatch = _ad_getobjectsinou("", "(&(objectSID=" & $idgroup & "))", 2, "sAMAccountName")
										If @error > 0 AND NOT IsArray($idgroup) Then
											MsgBox(64, "Info !", "SID key non trouvé, err: " & @error & @CRLF & "  recherche pour SID key:  " & $idgroup)
											$historik = $historik & "SID key non trouvé, err: " & @error & @CRLF & "  recherche pour SID key:  " & $idgroup
											ToolTip("", 5, 5)

									   Guidelete($hGUI)
									   Return 0
										ElseIf NOT IsArray($idgroup) Then
											ToolTip("", 5, 5)
											InputBox("search SID key", "SAM Account Name:  " & $displayname & @CRLF & $sidkeymatch[1], $idgroup)
											ToolTip("", 5, 5)
											Guidelete($hGUI)
											Return 0
										EndIf
										ToolTip("", 5, 5)
										Guidelete($hGUI)
										Return 0
									EndIf
								Else
									ToolTip("", 5, 5)
								EndIf
								If StringInStr($idgroup, ";") AND $idgroup <> "" AND $idrh1 = "" Then
									$idgroup = StringSplit($idgroup, ";")
									If NOT IsArray($idgroup) Then
										MsgBox(0, "Warning ! '0873' ", "no SIDkeys or Objects in array !", 7)
										ToolTip("", 5, 5)
										Guidelete($hGUI)
										Return 0
									EndIf
									Global $test = 0
									If _ispressed("10", $hdll) Then
										$test = 1
									EndIf
									Global $result2 = ""
									For $z = 1 To $idgroup[0]
										If $idgroup[$z] <> "" AND $idrh1 = "" AND NOT (StringRegExp($idgroup[$z], "^S-\d-\d+-(\d+-){1,14}\d+$") OR StringRegExp($idgroup[$z], "^S-\d-(\d+-){1,14}\d+$")) Then
											Global $displayname = ""
											If StringRegExp($idgroup[$z], "\h") AND StringLen($idgroup[$z]) > 2 Then
												Global $displayname = $idgroup[$z]
												Global $sou = ""
												Global $sldapfilter = "(&(objectCategory=Person)(ANR=" & $idgroup[$z] & " ))"
												Global $sdatatoretrieve = "sAMAccountName"
												Global $auserids = _ad_getobjectsinou($sou, $sldapfilter, 2, $sdatatoretrieve, "displayName")
												If IsArray($auserids) Then
													$idgroup[$z] = $auserids[1]
												EndIf
											EndIf
											Global $testbinaryh2 = ""
											$testbinaryh = _ad_getobjectproperties($idgroup[$z], "SIDhistory")
											If IsArray($testbinaryh) AND UBound($testbinaryh) >= 2 Then
												For $y = 1 To UBound($testbinaryh) - 1
													$testbinaryh2 = $testbinaryh2 & $testbinaryh[$y][1] & "|"
												Next
												$testbinaryh = $testbinaryh2
												$testbinaryh = StringTrimRight($testbinaryh, 1)
												$testbinaryh = "  SIDhistory:  " & $testbinaryh
											Else
												$testbinaryh = ""
											EndIf
											If NOT StringInStr($idgroup[$z], "S-1") Then
												$sidkeyg = _ad_getobjectproperties($idgroup[$z], "ObjectSID")
												If IsArray($sidkeyg) Then
													$sidkeyg1 = _arraytostring($sidkeyg)
													$sidkeyg1 = StringSplit($sidkeyg1, "|")
													$sidkeyg1 = $sidkeyg1[3]
													If $test = 1 Then
														$result2 = $result2 & "|" & $sidkeyg1 & "|" & $idgroup[$z] & @CRLF
													Else
														$result2 = $result2 & "SIDkey trouvé:  " & $sidkeyg1 & "  pour objectID:  " & $idgroup[$z] & "  " & $displayname & $testbinaryh & @CRLF
													EndIf
													ToolTip("", 5, 5)
												ElseIf NOT IsArray($sidkeyg) Then
													$result2 = $result2 & "SIDkey non trouvé pour objectID:  " & $idgroup[$z] & @CRLF
												EndIf
											EndIf
										EndIf
										If $idgroup[$z] <> "" AND $idrh1 = "" AND (StringRegExp($idgroup[$z], "^S-\d-\d+-(\d+-){1,14}\d+$") OR StringRegExp($idgroup[$z], "^S-\d-(\d+-){1,14}\d+$")) Then
											$sidkeymatch = _ad_getobjectsinou("", "(&(objectSID=" & $idgroup[$z] & "))", 2, "sAMAccountName")
											If @error > 0 Then
												$result2 = $result2 & "Compte  'Idrh' non trouvé pour SIDkey :  " & $idgroup[$z] & @CRLF
												ToolTip("", 5, 5)
											Else
												ToolTip("", 5, 5)
												$result2 = $result2 & "Compte  'Idrh' non trouvé pour SIDkey :  " & $sidkeymatch[1] & "  " & $idgroup[$z] & @CRLF
												ToolTip("", 5, 5)
											EndIf
											ToolTip("", 5, 5)
										EndIf
										ToolTip("", 5, 5)
									Next
									ToolTip("", 5, 5)
									If $result2 <> "" Then
										ClipPut($result2)
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $result2 & @CRLF & @CRLF & "			=====	=====	=====	=====	=====			" & @CRLF & @CRLF);, 1)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $result2
									 EndIf
									 ToolTip("", 5, 5)
									Guidelete($hGUI)
									Return 0
								EndIf
								$idrh1 = GUICtrlRead($buttonidrh1)
								If StringInStr($idrh1, "locked") Then
									ToolTip("Cherche les comptes vérouillés dans l' AD...", 5, 5)
									$alocked = _ad_getobjectslocked()
									If @error > 0 Then
										MsgBox(64, "Info !", "Aucun compte vérouillé n'a été trouvé !, err: " & @error)
									Else
										ToolTip("", 5, 5)
										__arraydisplay($alocked, "Comptes utilisateurs 'Locked' / 'verouillés'")
									EndIf
									ToolTip("", 5, 5)
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  recherche des comptes vérouillés"
									Guidelete($hGUI)
									Return 0
								EndIf
								$idrh1 = GUICtrlRead($buttonidrh1)
								If StringInStr($idrh1, "disabled") Then
									ToolTip("Recherche des comptes désactivés 'disabled' ...", 5, 5)
									$adisabled = _ad_getobjectsdisabled()
									If @error > 0 Then
										MsgBox(64, "Info !", "Aucun compte désactivé n'a été trouvé !, err: " & @error)
									Else
										ToolTip("", 5, 5)
										__arraydisplay($adisabled, "Comptes users 'désactivés'")
									EndIf
									ToolTip("", 5, 5)
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  recherche des comptes désactivés 'disabled'"
									Guidelete($hGUI)
									Return 0
								EndIf
								$idrh1 = GUICtrlRead($buttonidrh1)
								If StringInStr(($idrh1), "SID") AND NOT StringInStr($idrh1, "SIDh") AND StringLen($idrh1) = 3 Then
									ToolTip("Recherche des SID dans l'AD... (all category: group,person etc...)", 5, 5)
									$sid = _ad_getobjectsinou("", "(&(objectSID=*))", 2, "sAMAccountName,cn,displayName,objectSid,sidhistory")
									If @error > 0 Then
										MsgBox(64, "Info !", "Aucun SID n'a été trouvé !, err: " & @error)
									Else
										ToolTip("", 5, 5)
										ClipPut(_arraytostring($sid, ";"))
										__arraydisplay($sid, "comptes SID")
									EndIf
									ToolTip("", 5, 5)
									Guidelete($hGUI)
									Return 0
								EndIf
								$idrh1 = GUICtrlRead($buttonidrh1)
								If StringInStr(($idrh1), "SIDh") AND StringLen($idrh1) = 4 Then
									ToolTip("Recherche des SID history dans l'AD...(all category: group,person etc...)", 5, 5)
									$sidh = _ad_getobjectsinou("", "(&(SIDhistory=*))", 2, "sAMAccountName,cn,displayName,objectSid,sidhistory")
									If @error > 0 Then
										MsgBox(64, "Info !", "Aucun SID history n'a été trouvé !, err: " & @error)
									Else
										ToolTip("", 5, 5)
										ClipPut(_arraytostring($sidh, ";"))
										__arraydisplay($sidh, "SID history Accounts")
									EndIf
									ToolTip("", 5, 5)
									Guidelete($hGUI)
									Return 0
								EndIf
								$idrh1 = GUICtrlRead($buttonidrh1)
								If StringInStr($idrh1, "iliade?") OR StringInStr($idrh1, "lbpai?") Then
									$groupslbpailiade = "Chargé de Développement" & @CRLF & "Direction des Ressources Humaines" & @CRLF & "Direction adminstrative et financière" & @CRLF & "Direction Technique" & @CRLF & "Direction juridique et risque" & @CRLF & "Direction des systèmes d'information" & @CRLF & "Direction Marketing et Commerciale" & @CRLF & "Standard" & @CRLF & "commun full users" & @CRLF & "Direction générale" & @CRLF & @CRLF & "Autres categories:" & @CRLF & "----------------------------" & @CRLF & "BDD_DAF" & @CRLF & "Bibliotheque_DP" & @CRLF & "Comité de Direction" & @CRLF & "Communication Interne" & @CRLF & "Contrats" & @CRLF & "DFRC" & @CRLF & "Direction Financiere" & @CRLF & "Direction Systèmes Informatiques" & @CRLF & "DRA" & @CRLF & "Echange_DAF-DT" & @CRLF & "Echange_DM-DT" & @CRLF & "Echange_DRA-DT" & @CRLF & "Gourvernance" & @CRLF & "Risques et Conformité" & @CRLF & "Services Generaux" & @CRLF & "Supports de présentation"
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & @CRLF & "		LBPAI/Iliade [Categories]" & @CRLF & @CRLF & $groupslbpailiade & @CRLF & @CRLF
									_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "		LBPAI/Iliade [Categories]" & @CRLF & @CRLF & $groupslbpailiade & @CRLF & @CRLF);, 1)
									Guidelete($hGUI)
									Return 0
								EndIf
								ToolTip("Recherche par nom commun ambigu (ANR) : " & $idrh1 & " patience... 'Ambiguous Name Resolution'", 5, 5)
								Global $sou = ""
								Global $sldapfilter = "(&(objectCategory=Person)(ANR=" & $idrh1 & " ))"
								If $filtreperso = 1 Then
									$sdatatoretrieve = InputBox("data to retrieve", "Ex: cn,sAMAccountName, initials,mail..", "Displayname,cn,SAMaccountName")
									_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "ANR: filtre modifié = " & $sdatatoretrieve & @CRLF);, 1)
								Else
									Global $sdatatoretrieve = "sAMAccountName,cn,DisplayName,initials,title,mail,telephoneNumber,Mobile,physicalDeliveryOfficeName,comment,Laposte00-NetworkDrive"
									_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "ANR: filtre par defaut = " & $sdatatoretrieve & @CRLF);, 1)
								EndIf
								Global $auserids = _ad_getobjectsinou($sou, $sldapfilter, 2, $sdatatoretrieve, "displayName")
								If IsArray($auserids) Then
									ToolTip("", 5, 5, "")
									ClipPut(_arraytostring($auserids, ";"))
									__arraydisplay($auserids, "M.1/ common usernames found ! for keyworld,  '" & $idrh1 & "'")
								EndIf
								If IsArray($auserids) = 0 Then
								   ToolTip("", 5, 5)
									$readtexbox = $idrh1
									Global $sou = ""
									Global $sldapfilter = "(&(objectCategory=Person)(ANR=" & $readtexbox & "))"
									If $filtreperso = 1 Then
										$sdatatoretrieve = InputBox("data to retrieve", "Ex: cn,sAMAccountName, initials,mail..", "Displayname,cn,SAMaccountName")
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "ANR: filtre modifié = " & $sdatatoretrieve & @CRLF);, 1)
									Else
										Global $sdatatoretrieve = "sAMAccountName,cn,DisplayName,initials,title,mail,telephoneNumber,Mobile,physicalDeliveryOfficeName,comment"
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "ANR: filtre par defaut = " & $sdatatoretrieve & @CRLF);, 1)
									EndIf
									Global $auserids = _ad_getobjectsinou($sou, $sldapfilter, 2, $sdatatoretrieve, "displayName")
									If IsArray($auserids) Then
										ToolTip("", 5, 5, "")
										ClipPut(_arraytostring($auserids, ";"))
										__arraydisplay($auserids, "M.2/ common usernames found ! for keyworld,  '" & $readtexbox & "'")
									EndIf
								EndIf
								If IsArray($auserids) = 0 Then
									ToolTip("", 5, 5)
									MsgBox(64, "Warning !", "Aucun user trouvé avec filtre = " & $idrh1 & @CRLF & @CRLF & "  ... recherche dans les groupes :  ObjectClass= Groups  !", 20)
									Local $vargroup = _ad_getobjectsinou("", "(&(objectClass=group)(name=*" & $idrh1 & "*))", 2, "samaccountname")
									If IsArray($vargroup) Then
										$hwnd = GUICreate("Liste des groupes avec filtre = " & $idrh1, 1000, 60)
										$combo = GUICtrlCreateCombo("", 10, 10, 980, 60)
										For $d = 1 To UBound($vargroup) - 1
											GUICtrlSetData($combo, $vargroup[$d] & "  ==>  " & "[" & _ad_getobjectattribute($vargroup[$d], "distinguishedName") & "]", $vargroup[1] & "  ==>  " & "[" & _ad_getobjectattribute($vargroup[$d], "distinguishedName") & "]")
										Next
										GUISetState()
										Do
										Until GUIGetMsg() = $gui_event_close
										GUIDelete()
										ClipPut(_arraytostring($vargroup))
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & " recherche des noms communs 'ANR'" & @CRLF & _arraytostring($vargroup)
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & " recherche des noms communs 'ANR'" & @CRLF & _arraytostring($vargroup) & @CRLF);, 1)
										ToolTip("", 5, 5)
										Guidelete($hGUI)
										Return 0
									 Else
										ToolTip("", 5, 5,"")
										MsgBox(64, "Warning !", "Aucun groupe avec le filtre = " & $idrh1)
									 EndIf
									 ToolTip("", 5, 5,"")
									Guidelete($hGUI)
									Return 0
								EndIf
								Global $adatatoretrieve = StringSplit($sdatatoretrieve, ",", $str_nocount)
								For $i = 0 To UBound($auserids, 2) - 1
									$auserids[0][$i] = $adatatoretrieve[$i]
								Next
								ToolTip("", 5, 5)
								Guidelete($hGUI)
								Return 0
							Case $multiplesrce = 1
								$defautdc = $defautdcinit
								If _ispressed("10", $hdll) Then
									$test = 1
									ToolTip("selection multiple Idrh  et [-] groupes...", 5, 5)
								Else
									$test = 0
									ToolTip("selection multiple Idrh  et [+] groupes...", 5, 5)
								EndIf
								$multipleidrh = GUICtrlRead($buttonidrh1)
								If $multipleidrh = "" Then
									ToolTip("", 5, 5)
									MsgBox(0, "Info !", "abandon ! pas de user defini avec séparateur ';' dans 'Source' ... " & $multipleidrh, 7)
									Guidelete($hGUI)
									Return 0
								EndIf
								If $isdct = 1 Then
									$unite = InputBox("default OU ? 2 or 4 chr$", "ex: MILY, MITE, BPLY, GAUB, MI ... virtuos" & @CRLF & "?? : scan all OUs...", "MI")
									If StringLen($unite) < 2 OR StringLen($unite) > 4 AND NOT ($unite = "virtuos" OR $unite = "??") Then
										ToolTip("", 5, 5)
										MsgBox(0, "Info !", "abandon ! pas d'OU cible definie... " & $unite, 7)
										Guidelete($hGUI)
										Return 0
									EndIf
								Else
									$unite = "??"
								EndIf
								If $unite = "??" Then
									$allou = 1
									treeview_affiche()
									treeselect()
								Else
									$allou = 0
								EndIf
								$defautdc = StringReplace($defautdc, ".", ",DC=")
								$defautdc = ",DC=" & $defautdc
								If $allou = 1 Then
									$unite = ""
									$allou = 0
									Global $aous = _ad_getobjectsinou("" & $sout)
								Else
									Global $aous = _ad_getobjectsinou("OU=Groupes,OU=" & $unite & ",OU=" & StringMid($unite, 1, 2) & $defautdc)
								EndIf
								If $unite = "MI" Then
									Global $aous = _ad_getobjectsinou("OU=Groupes,OU=_Support Transverse,OU=MI" & $defautdc)
								EndIf
								If $unite = "virtuos" Then
									Global $aous = _ad_getobjectsinou("OU=FunctionalGroups,OU=_Flexible Workspace Commun" & $defautdc)
								EndIf
								$sout = ""
								$filez = FileOpen(@ScriptDir & "\Groups-AD_OU-" & $unite & ".txt", 2)
								If @error > 0 Then
									MsgBox(64, $adroot, "Pas de groupes trouvés ! ")
									ToolTip("", 5, 5)
									Guidelete($hGUI)
									Return 0
								Else
									$filterou = InputBox("filtrer groupes ?", "vide = tous les groupes , ou filtre: _cbk, _dpi ... ", "")
									Global $aselected[1]
									If $test = 1 Then
										Global $hgui2 = GUICreate("Selection groupes [-]  et [ Apply ] !  (" & $filterou & ")", 1000, 500)
									Else
										Global $hgui2 = GUICreate("Selection groupes [+]  et [ Apply ] !  (" & $filterou & ")", 1000, 500)
									EndIf
									Global $hlistbox = _guictrllistbox_create($hgui2, "String upon creation", 10, 30, 980, 470, BitOR($ws_border, $lbs_hasstrings, $ws_vscroll, $lbs_multiplesel))
									_guictrllistbox_beginupdate($hlistbox)
									_guictrllistbox_resetcontent($hlistbox)
									_guictrllistbox_initstorage($hlistbox, 100, 4096)
									_guictrllistbox_endupdate($hlistbox)
									If IsArray($aous) = 1 Then
										For $z = 1 To $aous[0] - 1
											FileWriteLine($filez, $aous[$z] & @CRLF)
											If StringLen($filterou) = 0 Then
												_guictrllistbox_addstring($hlistbox, $aous[$z])
											Else
												If NOT StringInStr($filterou, ";") Then
													If StringInStr($aous[$z], $filterou) Then
														_guictrllistbox_addstring($hlistbox, $aous[$z])
													EndIf
												Else
													Local $filterou2 = StringSplit($filterou, ";")
													For $w = 1 To $filterou2[0]
														If StringInStr($aous[$z], $filterou2[$w]) Then
															_guictrllistbox_addstring($hlistbox, $aous[$z])
														EndIf
													Next
												EndIf
											EndIf
										Next
									EndIf
									FileClose($filez)
									If _ispressed("10", $hdll) Then
									Else
										FileDelete(@ScriptDir & "\Groups-AD_OU-" & $unite & ".txt")
									EndIf
								EndIf
								If IsArray($aous) = 0 Then
									$idcheckboxgroup = 2
									GUIDelete($hgui2)
									ToolTip("", 5, 5)
									MsgBox(0, "Info !", "abandon, pas de groupes dans l'OU cible: " & $unite, 7)
									Guidelete($hGUI)
									Return 0
								EndIf
								Global $hbutton = GUICtrlCreateButton("[ Apply ]", 20, 3, 80, 25)
								GUISetState()
								While 1
									Switch GUIGetMsg()
										Case $hbutton
											$aselected = _guictrllistbox_getselitems($hlistbox)
											If $aselected[0] = 1 Then
												$sitem = " item"
											Else
												$sitem = " items"
											EndIf
											$sitems = ""
											For $i = 1 To $aselected[0]
												$sitems &= _guictrllistbox_gettext($hlistbox, $aselected[$i]) & @CRLF
											Next
											If $aselected[0] <> 0 Then
												$sitems = StringSplit($sitems, @CRLF)
												Global $sidkeymatch = ""
												Global $multipleidrhbis = ""
												$multipleidrh = StringSplit($multipleidrh, ";")
												If $test = 0 Then
													$result = "		[+]   multiple users groupes selectionnés...	" & @CRLF
												Else
													$result = "		[-]	  multiple users groupes selectionnés...	" & @CRLF
												EndIf
												For $z = 1 To $sitems[0]
													If StringLen($sitems[$z]) <> 0 Then
														If $test = 0 Then
															$result = $result & @CRLF & "[+] " & $sitems[$z] & @CRLF
														Else
															$result = $result & @CRLF & "[-] " & $sitems[$z] & @CRLF
														EndIf
														For $i = 1 To $multipleidrh[0]
															Global $sidkeymatch = ""
															Global $multipleidrhbis = ""
															Global $displayname = ""
															$multipleidrhbis = $multipleidrh[$i]
															If StringRegExp($multipleidrhbis, "\h") AND StringLen($multipleidrhbis) > 2 Then
																Global $displayname = $multipleidrhbis
																Global $sou = ""
																Global $sldapfilter = "(&(objectCategory=Person)(ANR=" & $multipleidrhbis & " ))"
																Global $sdatatoretrieve = "sAMAccountName"
																Global $auserids = _ad_getobjectsinou($sou, $sldapfilter, 2, $sdatatoretrieve, "displayName")
																If IsArray($auserids) Then
																	$multipleidrhbis = $auserids[1]
																EndIf
															EndIf
															If StringInStr($multipleidrhbis, "-") AND StringLen($multipleidrhbis) >= 20 Then
																$sidkey = _ad_getobjectsinou("", "(&(objectSID=" & $multipleidrhbis & "))", 2, "sAMAccountName")
																If @error > 0 Then
																EndIf
																If IsArray($sidkey) Then
																	$multipleidrhbis = $sidkey[1]
																EndIf
															EndIf
															$userexist = _ad_objectexists($multipleidrhbis)
															If $userexist = 1 AND $test = 0 Then
																Global $ivalue = _ad_addusertogroup($sitems[$z], $multipleidrhbis)
																If $ivalue = 1 Then
																	$result = $result & $multipleidrhbis & "   " & $displayname & "  [+] ok !"
																ElseIf @error = 2 Then
																	$result = $result & "l'OU cible n'a pas été trouvée !"
																ElseIf @error = 3 Then
																	$result = $result & $multipleidrhbis & "   " & $displayname & "  deja membre !"
																ElseIf @error = 3 Then
																	$result = $result & $multipleidrhbis & "   " & $displayname & "  n'est pas membre !"
																ElseIf @error >= 4 OR @error = -2147352567 Then
																	$result = $result & $multipleidrhbis & "   " & $displayname & "  accès refusé !"
																EndIf
															Else
																If $multipleidrhbis <> "" AND $userexist = 0 Then
																	$result = $result & $multipleidrhbis & " pas trouvé sur l'AD ! "
																EndIf
															EndIf
															If $userexist = 1 AND $test = 1 Then
																Global $ivalue = _ad_removeuserfromgroup($sitems[$z], $multipleidrhbis)
																If $ivalue = 1 Then
																	$result = $result & $multipleidrhbis & "   " & $displayname & "  [-] ok !"
																ElseIf @error = 2 Then
																	$result = $result & "l'OU cible n'a pas été trouvée !"
																ElseIf @error = 3 Then
																	$result = $result & $multipleidrhbis & "   " & $displayname & "  n'est pas membre !"
																ElseIf @error = 3 Then
																	$result = $result & $multipleidrhbis & "   " & $displayname & "  deja membre !"
																ElseIf @error >= 4 OR @error = -2147352567 Then
																	$result = $result & $multipleidrhbis & "   " & $displayname & "  accès refusé !"
																EndIf
															Else
																If $multipleidrhbis <> "" AND $userexist = 0 Then
																	$result = $result & $multipleidrhbis & " pas trouvé sur l'AD ! "
																EndIf
															EndIf
															$result = $result & @CRLF
														Next
													EndIf
												Next
												ToolTip("", 5, 5)
												GUIDelete($hgui2)
												ClipPut($result)
												_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $result & @CRLF & @CRLF & "			=====	=====	  		=====	=====			" & @CRLF & @CRLF);, 1)
												$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $result
												;_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $result & @CRLF);, 1)
												ExitLoop
											Else
												MsgBox(4160, "warning !", "Pas de selection... réessayer ou refermer la fenetre [x]...")
											EndIf
										Case $gui_event_close
											$idcheckboxgroup = 2
											GUIDelete($hgui2)
											ExitLoop
									EndSwitch
								WEnd
								ToolTip("", 5, 5)
								Guidelete($hGUI)
								Return 0
							Case StringUpper($idrh1) = "" OR $userexist = 0 AND $multiplesrce = 0 AND $idcheckboxcreate = 0
								$defautdc = $defautdcinit
								MsgBox(0, "Attention !", "utilisateur 'Srce'  " & $idrh1 & "  = <vide> ou <Idrh> inconnu !", 7)

								Guidelete($hGUI)
								Return 0
							Case $idcheckboxcreate = 1 AND $userexist = 0 OR ($idcheckboxcreate = 1 AND _ispressed("10", $hdll))
								$defautdc = $defautdcinit
								If _ispressed("10", $hdll) Then
									If _ad_objectexists($idrh1) = 0 Then
										MsgBox(0, "Warning !", "Id: " & $idrh1 & " not found...", 15)
										Guidelete($hGUI)
										Return 0
									EndIf
									Global $ireply = MsgBox(308, "Renommer user Srce ?", "voulez-vous renommer user Srce:  '" & $idrh1 & "'  ?")
									If $ireply <> 6 Then
										Guidelete($hGUI)
										Return 0
									EndIf
									Global $prenomidrh1 = ""
									Global $nomidrh1 = ""
									Global $samidrh1 = ""
									Global $formr = GUICreate("Renommer user Srce:  [ " & $idrh1 & " ]  -  " & _ad_getobjectattribute($idrh1, "givenName") & " " & _ad_getobjectattribute($idrh1, "sn"), 814, 180)
									GUICtrlCreateLabel("IDRH: ", 38, 12, 231, 17)
									GUICtrlCreateLabel("Prenom: ", 38, 42, 231, 17)
									GUICtrlCreateLabel("Nom: ", 38, 82, 231, 17)
									GUICtrlCreateLabel("Description: ", 38, 122, 231, 17)
									Global $iobject = GUICtrlCreateInput(StringUpper($idrh1), 141, 8, 559, 21)
									Global $inewnamegn = GUICtrlCreateInput(_ad_getobjectattribute($idrh1, "givenName"), 141, 40, 559, 21)
									Global $inewnamesn = GUICtrlCreateInput(_ad_getobjectattribute($idrh1, "sn"), 141, 80, 559, 21)
									Global $inewdescription = GUICtrlCreateInput(_ad_getobjectattribute($idrh1, "description"), 141, 120, 559, 21)
									Global $initialprenom = _ad_getobjectattribute($idrh1, "givenName")
									Global $initialnom = _ad_getobjectattribute($idrh1, "sn")
									Global $initialdesc = _ad_getobjectattribute($idrh1, "description")
									Global $bok = GUICtrlCreateButton("Rename user", 250, 143, 130, 33)
									Global $bcancel = GUICtrlCreateButton("Cancel", 500, 143, 73, 33, BitOR($gui_ss_default_button, $bs_defpushbutton))
									GUISetState(@SW_SHOW)
									While 1
										Global $nmsg = GUIGetMsg()
										Switch $nmsg
											Case $gui_event_close, $bcancel
												$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "   renommer un compte utilisateur AD:  " & $idrh1 & "  commande annulée !"
												_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
												_GUICtrlRichEdit_SetFont($aff, 10)                   ;   set size
												_GUICtrlRichEdit_SetCharAttributes($aff, '+bo')      ;   set bold
												_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "   renommer un compte utilisateur AD:  " & $idrh1 & "  commande annulée !" & @CRLF);, 1)
												GUIDelete($formr)
												Guidelete($hGUI)
												Return 0
											Case $bok
												$samidrh1 = GUICtrlRead($iobject)
												$prenomidrh1 = GUICtrlRead($inewnamegn)
												$nomidrh1 = GUICtrlRead($inewnamesn)
												$descriptionidrh1 = GUICtrlRead($inewdescription)
												GUIDelete($formr)
												$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "   renommer un compte utilisateur AD:  " & $idrh1 & "  commande validée !"
												_GUICtrlRichEdit_SetCharColor($aff, 0x00ff0000)      ;   set color blue (BBGGRR format)
												_GUICtrlRichEdit_SetFont($aff, 10)                   ;   set size
												_GUICtrlRichEdit_SetCharAttributes($aff, '+bo')      ;   set bold
												_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "   renommer un compte utilisateur AD:  " & $idrh1 & "  commande validée !" & @CRLF);, 1)
												ExitLoop
										EndSwitch
									WEnd
									If _ad_objectexists($samidrh1) = 0 OR ($samidrh1 = $idrh1) Then
										Global $ivalue = _ad_renameobject($idrh1, $prenomidrh1 & " " & $nomidrh1)
										If $ivalue = 1 Then
											MsgBox(64, "Renommer user Srce...", "User [" & $idrh1 & "] renommé en..." & @CRLF & @CRLF & "[" & $samidrh1 & "] -  " & $prenomidrh1 & " " & $nomidrh1 & "  -  '" & $descriptionidrh1 & "'" & @CRLF & @CRLF & "initial:" & @CRLF & @CRLF & "[" & StringUpper($idrh1) & "] -  " & $initialprenom & " " & $initialnom & "  -  '" & $initialdesc & "'")
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "   Renommer user Srce..." & "  User [" & $idrh1 & "] renommé en..." & @CRLF & @CRLF & "[" & $samidrh1 & "] -  " & $prenomidrh1 & " " & $nomidrh1 & "  -  '" & $descriptionidrh1 & "'" & @CRLF & @CRLF & " initial:" & @CRLF & @CRLF & "[" & StringUpper($idrh1) & "] -  " & $initialprenom & " " & $initialnom & "  -  '" & $initialdesc & "'"
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "   Renomme user Srce..." & "  User [" & $idrh1 & "] renommé en..." & @CRLF & @CRLF & "[" & $samidrh1 & "] -  " & $prenomidrh1 & " " & $nomidrh1 & "  -  '" & $descriptionidrh1 & "'" & @CRLF & @CRLF & " initial:" & @CRLF & @CRLF & "[" & StringUpper($idrh1) & "] -  " & $initialprenom & " " & $initialnom & "  -  '" & $initialdesc & "'" & @CRLF);, 1)
										ElseIf @error = 1 Then
											MsgBox(64, "Renomme user Srce...", "user Srce '" & $samidrh1 & "' n'existe pas !")
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "   utilisateur  " & $samidrh1 & "  inexistant !"
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "   utilisateur  " & $samidrh1 & "  inexistant !" & @CRLF);, 1)
										Else
											MsgBox(64, "Renomme user Srce...", "Return code '" & @error & "' from Active Directory")
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "   code erreur: " & @error & "  accès refusé ?!"
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "   code erreur: " & @error & "  accès refusé ?!" & @CRLF);, 1)
										EndIf
										_ad_modifyattribute($idrh1, "givenName", $prenomidrh1, 2)
										_ad_modifyattribute($idrh1, "sn", $nomidrh1, 2)
										_ad_modifyattribute($idrh1, "cn", $prenomidrh1 & " " & $nomidrh1, 2)
										_ad_modifyattribute($idrh1, "displayName", $prenomidrh1 & " " & $nomidrh1, 2)
										_ad_modifyattribute($idrh1, "description", $descriptionidrh1, 2)
										$iobject = _ad_samaccountnametofqdn($idrh1)
										$iobject = _stringexplode($iobject, "DC", 1)
										$iobject = $iobject[1]
										$iobject = StringReplace($iobject, "DC=", "")
										$iobject = StringReplace($iobject, ",", ".")
										$iobject = StringReplace($iobject, "=", "")
										_ad_modifyattribute($idrh1, "userPrincipalName", $samidrh1 & "@" & $iobject, 2)
										_ad_modifyattribute($idrh1, "sAMAccountName", $samidrh1, 2)
									Else
										MsgBox(0, "Warning !", "'" & $idrh1 & "'" & " impossible de changer Idrh Srce '" & $samidrh1 & "' !" & @CRLF & " Déjà utilisé ...", 15)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "   Warning !  " & "  '" & $idrh1 & "'" & " impossible de changer Idrh Srce '" & $samidrh1 & "' !" & "   " & " Déjà utilisé !"
										  _GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
										  _GUICtrlRichEdit_SetFont($aff, 10)                   ;   set size
										  _GUICtrlRichEdit_SetCharAttributes($aff, '+bo')      ;   set bold
									_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "   Warning !  " & "  '" & $idrh1 & "'" & " ne peut etre renommé en '" & $samidrh1 & "' !" & "   " & " l'idrh existe déjà  !" & @CRLF);, 1)
									EndIf
									Guidelete($hGUI)
									Return 0
								Else
									If $isdct = 1 Then
										$unite = InputBox("default OU ?", "ex: MILY, MITE, BPLY, GAUB ... " & @CRLF & "?? : scan all OUs...", "")
										If @error Then
											MsgBox(0, "Info !", "Abandon... creation user Srce !")
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  annulation creation nouvel utilisateur !"
											 _GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
											 _GUICtrlRichEdit_SetFont($aff, 10)                   ;   set size
											 _GUICtrlRichEdit_SetCharAttributes($aff, '+bo')      ;   set bold
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  annulation creation nouvel utilisateur !" & @CRLF);, 1)
											Guidelete($hGUI)
											Return 0
										EndIf
										If StringLen($unite) < 2 OR StringLen($unite) > 4 AND ($unite <> "??") Then
											Guidelete($hGUI)
											Return 0
										EndIf
									Else
										$unite = "??"
									EndIf
									If $unite = "??" Then
										If StringUpper($domainname) = "COURRIER" Then
											ToolTip("OUP/OUP-Z/OUP-Z(nn)", 5, 5, "domaine COURRIER: placez-vous dans l'OU cible OUP-Z(nn)..")
										EndIf
										$allou = 1
										treeview_affiche()
										treeselect()
										ToolTip("", 5, 5, "")
									EndIf
									If $allou = 1 Then
										$sou = $sout
										$unite = ""
										$allou = 0
									Else
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sou = "OU=Utilisateurs,OU=" & $unite & ",OU=" & StringMid($unite, 1, 2) & $defautdc
									EndIf
									$sout = ""
									$idrh1 = StringUpper($idrh1)
									$prenomidrh1 = InputBox("prenom pour " & $idrh1, " ", "")
									$nomidrh1 = InputBox("Nom pour " & $idrh1, " ", "")
									$nomidrh1 = StringUpper($nomidrh1)
									Global $passwordtoset = ""
									If StringInStr($unite, "CF") Then
										$passwordtoset = InputBox(" Pwd initial pour users du [Cnah]", "Initial Pwd set for: " & $idrh1, "Cnah" & @YEAR)
									ElseIf StringUpper($domainname) = "COURRIER" Then
										$passwordtoset = InputBox("initial Password", "Initial Pwd set for: " & $idrh1, "Abcd1234*")
									Else
										$passwordtoset = InputBox("initial Password", "Initial Pwd set for: " & $idrh1, "W!nd0ws10")
									EndIf
									If $passwordtoset = "" Then
										$passwordtoset = "W!nd0ws10"
									EndIf

									Local $jetonC=0
									Global $ivalue = _ad_createuser($sou, $idrh1, $prenomidrh1 & " " & $nomidrh1)
									If $ivalue = 1 Then
										MsgBox(64, "", "User Srce '" & $idrh1 & "' dans OU '" & $sou & "' crée" & @CRLF & @CRLF & "mot de passe initial (à changer): " & $passwordtoset)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & "' dans OU '" & $sou & "' à été crée, " & "le mot de passe initial (à changer): " & $passwordtoset & @CRLF & @CRLF
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idrh1 & "' in OU '" & $sou & "' à été crée, " & "le mot de passe initial (à changer): " & $passwordtoset & @CRLF  & @CRLF);, 1)
										If StringUpper($domainname) = "COURRIER" Then
											$oudep = InputBox("code regate", " code regate à 6 chiffres...", "")
											_ad_addusertogroup("GLD-Z-UtilEtab-Etab" & $oudep, $idrh1)
											_ad_addusertogroup("GGD-Z-DansEtab-Etab" & $oudep, $idrh1)
										 EndIf

										 If StringUpper($domainname) = "CORP" Then
											$oudep = InputBox("code regate", " code regate à 6 chiffres...", "")
											$test=StringUpper( StringLeft($idrh1,1) )
											if StringInStr($test,"X") Then	;compte en 'X'
											_ad_addusertogroup("Profil-ORG-" & $oudep & "-Externe_Local", $idrh1)
										    Else
											_ad_addusertogroup("Profil-ORG-" & $oudep & "-Interne_Local", $idrh1)
											EndIf
										 EndIf

										If $isdct = 1 Then
											$historik = $historik & "Autopilote, groupe rajouté: rg-pitr_cm_dir-tertiaire_standard_std" & @CRLF & @CRLF
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Autopilote, groupe rajouté: rg-pitr_cm_dir-tertiaire_standard_std" & @CRLF  & @CRLF);, 1)
										EndIf
										ToolTip("Travail en cours...", 5, 5, "")
										Sleep(3000)
										ToolTip("", 5, 5, "")
										_ad_modifyattribute($idrh1, "givenName", $prenomidrh1, 2)
										_ad_modifyattribute($idrh1, "sn", $nomidrh1, 2)
										_ad_modifyattribute($idrh1, "displayName", $prenomidrh1 & " " & $nomidrh1, 2)
										If $isdct = 1 Then
											_ad_addusertogroup("rg-pitr_cm_dir-tertiaire_standard_std", $idrh1)
										EndIf
										If StringInStr(StringLeft($idrh1, 1), "x") Then
											$hguicalendar = GUICreate("compte " & $idrh1, 300, 100, 350, 120)
											Local $iddate = GUICtrlCreateDate("", 30, 30, 100, 20, $dts_shortdateformat)
											Local $idbtn = GUICtrlCreateButton("OK", 120, 65, 60, 20)
											GUISetState(@SW_SHOW)
											While 1
												Switch GUIGetMsg()
													Case $gui_event_close
														GUIDelete($hguicalendar)
														Guidelete($hGUI)
														Return 0
													Case $idbtn
														Global $dateprolonge = GUICtrlRead(($iddate))
														GUIDelete($hguicalendar)
														ExitLoop
												EndSwitch
											WEnd
											$joursd = StringSplit($dateprolonge, "/")
											Global $jourd = $joursd[1]
											Global $moisd = $joursd[2]
											Global $anneed = $joursd[3]
											$dated = $anneed & "/" & $moisd & "/" & $jourd
											$ivalue = _ad_setaccountexpire($idrh1, $dated)
											_ad_setpassword($idrh1, $passwordtoset, 1)
											$pwdreset = @CRLF & @CRLF & "Compte Srce: " & $idrh1 & " , reinit. mot de passe à:  " & $passwordtoset & "  (Le mot de passe devra etre modifié à l'ouverture de session !)"
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $pwdreset
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $pwdreset & @CRLF);, 1)
										Else
											_ad_setpassword($idrh1, $passwordtoset, 1)
											$pwdreset = @CRLF & @CRLF & "Compte Srce: " & $idrh1 & " , reinit. mot de passe à:  " & $passwordtoset & "  (Le mot de passe devra etre modifié à l'ouverture de session !)"
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $pwdreset
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $pwdreset & @CRLF);, 1)
										EndIf
									ElseIf @error = 1 Then
										MsgBox(64, "", "User Srce '" & $idrh1 & "' existe déjà !")
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & " , existe deja !"
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idrh1 & " , existe deja !" & @CRLF);, 1)
									    $jetonC=1

									ElseIf @error = 2 Then
										MsgBox(64, "", "OU '" & $sou & "' n'existe pas")
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & " , n'existe pas !"
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idrh1 & " , n'existe pas !" & @CRLF);, 1)
									ElseIf @error = 3 Then
										MsgBox(64, "", "Value for CN (e.g. Lastname Firstname) est manquant !")
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & " , le nom ou le prenom n'a pas été saisi !"
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idrh1 & " , le nom ou le prenom n'a pas été saisi !" & @CRLF);, 1)
									ElseIf @error = 4 Then
										MsgBox(64, "", "Value for $sAD_User is missing")
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & " , manque une valeur !"
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idrh1 & " , manque une valeur !" & @CRLF);, 1)
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & " , code erreur " & @error & " !"
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF &  $idrh1 & " , code erreur " & @error & " !" & @CRLF);, 1)
										$jetonC=1
									 EndIf

									$userexist = _ad_objectexists($idrh1)
									If $userexist = 0 and $jetonC=1 Then
										MsgBox(0, "Warning !", "La creation à échouée car ce  Nom/prénom  existe sous un autre 'Idrh'" & @CRLF & "...ou droits dans cette OU !", 15)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  La création à echouée car l'utilisateur existe deja sous un autre 'idrh' !"
										  _GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
										  _GUICtrlRichEdit_SetFont($aff, 10)                   ;   set size
										  _GUICtrlRichEdit_SetCharAttributes($aff, '+bo')      ;   set bold
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF &  "  La création à echouée car l'utilisateur existe deja sous un autre 'idrh' !" & @CRLF);, 1)
										Global $rechercheanrc = ""
										$rechercheanrc = InputBox("Recherche par 'ANR'", "Creation du compte échoué !" & @CRLF & "recherche par 'Nom' ... ", "")
										Global $sou = ""
										Global $sldapfilter = "(&(objectCategory=Person)(ANR=" & $rechercheanrc & " ))"
										Global $sdatatoretrieve = "sAMAccountName,cn,displayName,title,distinguishedName,comment,Laposte00-NetworkDrive"
										Global $auserids = _ad_getobjectsinou($sou, $sldapfilter, 2, $sdatatoretrieve, "displayName")
										If @error Then
											ToolTip("", 5, 5)
											;MsgBox(64, "Warning !", "No users found ! " & $rechercheanrc)
											;Guidelete($hGUI)
											;Return 0
										 EndIf


										 If IsArray($auserids) Then ; ANR M.1
									ToolTip("", 5, 5, "")
									ClipPut(_arraytostring($auserids, ";"))
									__arraydisplay($auserids, "M.1/ ANR usernames trouvés ! avec mot clé,  '" & $idrh1 & "'")
								;EndIf
									     ElseIf IsArray($auserids) = 0 Then ;ANR M.2
								   ToolTip("", 5, 5)

									Global $rechercheanrc = ""
										;$rechercheanrc = InputBox("Recherche par 'ANR'", "Creation du compte échoué !" & @CRLF & "recherche par 'Nom' ... ", "")
										Global $sou = ""
										Global $sldapfilter = "(&(objectCategory=Person)(ANR=" & $rechercheanrc & " ))"
										Global $sdatatoretrieve = "sAMAccountName,cn,displayName,title,distinguishedName"
										Global $auserids = _ad_getobjectsinou($sou, $sldapfilter, 2, $sdatatoretrieve, "displayName")
										If @error Then
											ToolTip("", 5, 5)
										;	MsgBox(64, "Warning !", "No users found ! " & $rechercheanrc)
										;	Guidelete($hGUI)
										;	Return 0
										 EndIf
										 ;

										 If IsArray($auserids) Then
									ToolTip("", 5, 5, "")
									ClipPut(_arraytostring($auserids, ";"))
									__arraydisplay($auserids, "M.2/ ANR usernames trouvés ! avec mot clé,  '" & $idrh1 & "'")
								EndIf
										ToolTip("", 5, 5)
										;ClipPut(_arraytostring($auserids, ";"))
										;$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & " noms trouvés par recherche ANR... " & @CRLF & _arraytostring($auserids, ";" & @crlf)
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF &  " noms trouvés par recherche ANR... copiés vers le presse-papier" & @CRLF ) ;& _arraytostring($auserids, ";" & @CRLF) );, 1)
										;__arraydisplay($auserids, "common usernames found ! for keyworld,  '" & $rechercheanrc & "'")
										Guidelete($hGUI)
										Return 0
									EndIf
								 EndIf
							  EndIf ; fin elseif isarray $auserids

								$userexist = _ad_objectexists($idrh1)
								If $isdct = 1 and $userexist=1 Then
									$t = MsgBox(4, "Apply Directive ?", "Rajouter une directive après la création du compte  " & $idrh1 & "  ?")
									If $t = 6 Then
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  Rajouter une 'Directive' pour  " & $idrh1 & @CRLF & @CRLF
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  Rajouter une 'Directive' pour  " & $idrh1 & @CRLF & @CRLF) ;, 1)
										directives()
									Else
									EndIf
								 EndIf

							Case $idcheckboxcreate = 1 AND $userexist = 1
								$defautdc = $defautdcinit
								MsgBox(0, "warning !", " Idrh Srce [" & $idrh1 & "] existe déjà dans l'AD !", 7)
								$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  L'idrh " & $idrh1 & "  existe deja !"
								_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  L'idrh " & $idrh1 & "  existe deja !" & @CRLF);, 1)
								Guidelete($hGUI)
								Return 0
								$userexist = _ad_objectexists($idrh1)
							Case $idcheckboxprolonge = 1 AND StringInStr(StringLeft($idrh1, 1), "x")
								$defautdc = $defautdcinit
								Global $date_expire = ""
								$date_expire = _ad_getobjectproperties($idrh1, "accountExpires")
								$date_expire = $date_expire[1][1]
								$date_expire = _dateadd("d", -1, $date_expire)
								If $date_expire = "0000/00/00 00:00:00" OR $date_expire = "1601/01/01 00:00:00" Then
									MsgBox(0, "warning !", "Le compte '" & $idrh1 & "' , n'expire jamais...", 7)
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "warning !  , " & "Le compte '" & $idrh1 & "' , n'expire jamais..."
									_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF &  "warning !  , " & "Le compte '" & $idrh1 & "' , n'expire jamais..." & @CRLF);, 1)
								EndIf
								If $date_expire = "0000/00/00 00:00:00" OR $date_expire = "1601/01/01 00:00:00" Then
									$t = MsgBox(4, "compte AD en 'x' ", "Le compte " & $idrh1 & " est un compte en 'x' , souhaitez-vous mettre une date d'expiration ? (Oui)" & @CRLF & "   ou (NON), le compte n'expire jamais")
									If $t = 7 Then
										MsgBox(0, "no expiration date !", "[" & $idrh1 & "] : no expire date ...", 7)
										$date_expire = "01/01/1970 00:00:00"
										$ivalue = _ad_setaccountexpire($idrh1, $date_expire)
									Else
										;$date_expire = _dateadd("d", -1, $date_expire)
										$hguicalendar = GUICreate("compte " & $idrh1, 300, 100, 350, 120)
										Local $iddate = GUICtrlCreateDate($date_expire, 30, 30, 100, 20, $dts_shortdateformat)
										Local $idbtn = GUICtrlCreateButton("OK", 120, 65, 60, 20)
										GUISetState(@SW_SHOW)
										While 1
											Switch GUIGetMsg()
												Case $gui_event_close
													GUIDelete($hguicalendar)
													Guidelete($hGUI)
													Return 0
												Case $idbtn
													Global $dateprolonge = GUICtrlRead(($iddate))
													GUIDelete($hguicalendar)
													ExitLoop
											EndSwitch
										WEnd
										$joursd = StringSplit($dateprolonge, "/")
										Global $jourd = $joursd[1]
										Global $moisd = $joursd[2]
										Global $anneed = $joursd[3]
										$dated = $anneed & "-" & $moisd & "-" & $jourd & " 00:00:00"
										$dated = _dateadd("d", 1, $dated)
										$ivalue = _ad_setaccountexpire($idrh1, $dated)
										Sleep(500)
									EndIf
								Else
									If _ispressed("10", $hdll) Then
										MsgBox(0, "no expiration date !", "[" & $idrh1 & "] : retirer la date d'expiration, n'expire plus ...", 7)
										$date_expire = "01/01/1970 00:00:00"
										$ivalue = _ad_setaccountexpire($idrh1, $date_expire)
									Else
										;$date_expire = _dateadd("d", -1, $date_expire)
										$hguicalendar = GUICreate("compte " & $idrh1, 300, 100, 350, 120)
										Local $iddate = GUICtrlCreateDate($date_expire, 30, 30, 100, 20, $dts_shortdateformat)
										Local $idbtn = GUICtrlCreateButton("OK", 120, 65, 60, 20)
										GUISetState(@SW_SHOW)
										While 1
											Switch GUIGetMsg()
												Case $gui_event_close
													GUIDelete($hguicalendar)
													Guidelete($hGUI)
													Return 0
												Case $idbtn
													Global $dateprolonge = GUICtrlRead(($iddate))
													GUIDelete($hguicalendar)
													ExitLoop
											EndSwitch
										WEnd
										$joursd = StringSplit($dateprolonge, "/")
										Global $jourd = $joursd[1]
										Global $moisd = $joursd[2]
										Global $anneed = $joursd[3]
										$dated = $anneed & "-" & $moisd & "-" & $jourd & " 00:00:00"
										$dated = _dateadd("d", 1, $dated)
										$ivalue = _ad_setaccountexpire($idrh1, $dated)
										Sleep(500)
									EndIf
								EndIf
							Case $idcheckboxmove = 1 AND StringLen($idrh1) <> 0
								$defautdc = $defautdcinit
								Global $ousourcedir = _ad_getobjectattribute($idrh1, "distinguishedName")
								$ousourcedir = StringSplit($ousourcedir, ",")
								$ousourcedir = $ousourcedir[3]
								$ousourcedir = StringTrimLeft($ousourcedir, 3)
								If $isdct = 1 Then
									$ou = InputBox("default OU ?", "ex: MILY, MITE, BPLY, GAUB ... " & @CRLF & @CRLF & "   " & $idrh1 & " , OU actuelle:  [ " & $ousourcedir & " ]" & @CRLF & @CRLF & "?? : scan all OUs...", "")
									If StringLen($ou) < 2 OR StringLen($ou) > 4 AND NOT ($ou = "virtuos" OR $ou = "??") Then
										Guidelete($hGUI)
										Return 0
									EndIf
								Else
									$ou = "??"
								EndIf
								If $ou = "??" Then
									$allou = 1
									treeview_affiche()
									treeselect()
								EndIf
								If $allou = 1 Then
									$ou = ""
									$allou = 0
									$sou2 = $sout
									$ou = $sout
								Else
									$defautdc = StringReplace($defautdc, ".", ",DC=")
									$defautdc = ",DC=" & $defautdc
									$sou2 = "OU=Utilisateurs,OU=" & $ou & ",OU=" & StringMid($ou, 1, 2) & $defautdc
									$idrh1 = StringUpper($idrh1)
								EndIf
								$sobject = _ad_samaccountnametofqdn($idrh1)
								Global $ivalue = _ad_moveobject($sou2, $sobject)
								If $ivalue = 1 Then
									MsgBox(64, "Active Directory ", $idrh1 & ",  déplacé de  '" & StringUpper($ousourcedir) & "  vers   '" & StringUpper($ou) & "'")
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & " déplacé vers " & $ou
									_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF &  $idrh1 &  " déplacé de  '" & StringUpper($ousourcedir) &  "'  vers  '" & StringUpper($ou) & "'" & @CRLF);, 1)
									ToolTip("synchronising AD", 5, 5)
									Sleep(3000)
									ToolTip("", 5, 5)
								ElseIf @error = 1 Then
									MsgBox(64, "Active Directory ", "OU cible '" & $ou & "' n'existe pas !")
								ElseIf @error = 2 Then
									MsgBox(64, "Active Directory ", "User Srce '" & $idrh1 & "' n'existe pas !")
								ElseIf @error = 3 Then
									MsgBox(64, "Active Directory ", "User Srce '" & $idrh1 & "' déjà dans la meme OU '" & $ou & "'   :)")
								ElseIf @error >= 4 OR @error = -2147352567 Then
									MsgBox(64, "Active Directory ", "Droits refusés pour déplacer user Srce " & $idrh1)
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & "non deplacé vers " & $ou & ",  accès refusé !"
									_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF &  $idrh1 & "non deplacé vers " & $ou & ",  accès refusé !" & @CRLF);, 1)
								Else
									MsgBox(0, "warning !", "User Idrh Srce '" & $idrh1 & "existe déjà dans l'AD !", 7)
									Guidelete($hGUI)
									Return 0
								EndIf
							Case $idcheckboxdrive = 1 AND $userexist = 1
								$defautdc = $defautdcinit
								Global $drivesidrh1 = ""
								Global $drivesidrh1 = _ad_getobjectproperties($idrh1, "LaPoste00-NetworkDrive")
								_arraydelete($drivesidrh1, 0)
								$t = 6
								Global $listmanualdrives = ""
								While $t = 6
									$drivesidrh1 = _arraytostring($drivesidrh1, "|")
									Global $presencemesscan = 0
									If StringRegExp($drivesidrh1, "Mes scan") Then
										$presencemesscan = 1
									EndIf
									$drivesidrh1 = StringReplace($drivesidrh1, "LaPoste00-NetworkDrive", "")
									$drivesidrh1 = StringStripCR($drivesidrh1)
									If StringInStr($drivesidrh1, ":\\") Then
										$drivesidrh1 = StringReplace($drivesidrh1, ":\\", ";\\")
									EndIf
									$drivesidrh1 = StringSplit($drivesidrh1, "|")
									_arraydelete($drivesidrh1, 0)
									_arraydelete($drivesidrh1, 0)
									$chaine2 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
									Global $chaine5 = ""
									For $j = 0 To UBound($drivesidrh1) - 1
										For $k = 1 To 26
											$chaine = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
											$lettre = StringUpper (StringMid($chaine, $k, 1) )
											; $test = StringRegExp($drivesidrh1[$j], $lettre & ";\\", 3) ;old
											$test = StringRegExp( StringUpper (StringLeft( $drivesidrh1[$j],1 )), $lettre , 3)  ;rectif
											If UBound($test) >= 1 Then
												$chaine2 = StringReplace($chaine2, $lettre, "")
												$chaine5 = $chaine5 & $lettre
											 Else
											EndIf
										Next
									Next
									$chaine3 = $chaine2
									$chaine2 = StringRegExpReplace($chaine2, "", ".")
									$chaine2 = StringMid($chaine2, 2)
									$taillechaine = StringLen($chaine2)
									$chaine2 = StringTrimRight($chaine2, 1)
									$taille = StringLen($chaine2)
									$chaine1 = StringLeft($chaine2, Round($taille / 2) - 1)
									$chaine2 = StringMid($chaine2, Round($taille / 2), $taille)
									Global $presencems = ""
									If $presencemesscan = 0 Then
										$presencems = " , 'Mes scan'"
									Else
										$presencems = ""
									EndIf
									$driveismanual = InputBox("Entrez le chemin en entier !", "ex: T;\\429940SN001.429940S.IRGP\DOCUMENT " & @CRLF & @CRLF & "Lettres de lecteurs dispo:" & @CRLF & $chaine1 & @CRLF & $chaine2 & $presencems, "")
									If $driveismanual = "" or StringLen($driveismanual)<4 Then
										MsgBox(0, "Warning !", "drive ! chemin non saisi / incomplet , abandon ...", 7)
										Guidelete($hGUI)
										Return 0
									 EndIf

									 $test=StringUpper ( StringLeft($driveismanual,1) )
									 $driveismanual=StringTrimLeft($driveismanual,1)
									 $driveismanual=$test & $driveismanual

									 if StringIsAlpha ($test) = 0 Then
										MsgBox(0, "Warning !", "drive, la lettre saisie est manquante ! " & @crlf & " " & $driveismanual & @crlf & @CRLF & "  Abandon ...", 7)
										Guidelete($hGUI)
										Return 0
									 EndIf

									$manual2 = $driveismanual
									If StringRegExp($manual2, "Mes scan;") AND $presencemesscan = 1 Then
										MsgBox(0, "Warning !", "'Mes scan' déjà utilisé! abandon..." & @CRLF & "ajout: " & $driveismanual & @CRLF & "lettres lecteurs dispo: " & $chaine3, 7)
										Guidelete($hGUI)
										Return 0
									EndIf
									For $k = 1 To StringLen($chaine5)
										$chaine4 = StringMid($chaine5, $k, 1)
										$chaine4 = $chaine4 & ";"
										If StringInStr(StringUpper(StringLeft($manual2, 2)), $chaine4) Then
											MsgBox(0, "Warning !", "Drive lettre lecteur déjà utilisé ! abandon..." & @CRLF & $driveismanual & @CRLF & "lettres lecteurs dispo: " & $chaine3, 7)
											Guidelete($hGUI)
											Return 0
										EndIf
									Next
									$drivesidrh1 = _arraytostring($drivesidrh1, "|")
									If $drivesidrh1 = "" Then
										$t = 7
										$drivesidrh1 = $driveismanual
									Else
										$drivesidrh1 = $drivesidrh1 & "|" & $driveismanual
									EndIf
									If StringInStr($drivesidrh1, "|") Then
										$drivesidrh1 = StringSplit($drivesidrh1, "|")
									Else
									EndIf
									$listmanualdrives = $listmanualdrives & "|" & $driveismanual
									_arraydisplay($drivesidrh1, "LaPoste00-NetworkDrive modifié pour: [" & $idrh1 & "]")
									If $t = 7 Then
										ExitLoop
									Else
										$t = MsgBox(256 + 32 + 4, "[+] Drive ?", "Souhaitez-vous rajouter un autre lecteur ?")
										If $t = 6 Then
										ElseIf $t = 7 Then
											_arraydelete($drivesidrh1, 0)
											ExitLoop
										EndIf
									EndIf
								WEnd
								$listmanualdrives = StringTrimLeft($listmanualdrives, 1)
								Global $ivalue = _ad_modifyattribute($idrh1, "LaPoste00-NetworkDrive", $drivesidrh1, 2)
								If $ivalue = 1 Then
								ElseIf @error = 1 Then
									MsgBox(0, "info !", "[user Srce n'existe pas]", 7)
								Else
									MsgBox(0, "info !", "[Return error code " & @error & "] from Active Directory", 7)
								EndIf
								ToolTip("synchronising AD", 5, 5)
								Sleep(1000)
								ToolTip("", 5, 5)
							Case $idcheckboxgroupremove = 1 AND StringLen($idrh1) <> 0 AND $userexist = 1
								$defautdc = $defautdcinit
								Global $groupsidrh1 = _ad_getusergroups($idrh1)
								If @error > 0 Then
									MsgBox(0, "Info !", "user " & $idrh1 & " n'a aucun groupe ...", 7)
									Guidelete($hGUI)
									Return 0
								EndIf
								_arraysort($groupsidrh1, 0, 1)
								Global $aselected[1]
								Global $hgui2 = GUICreate("Selection groupes [-] pour '" & $idrh1 & "' , puis valider avec le bouton [ Apply ] !  ", 1000, 500)
								Global $hlistbox = _guictrllistbox_create($hgui2, "String upon creation", 10, 30, 980, 470, BitOR($ws_border, $lbs_hasstrings, $ws_vscroll, $lbs_multiplesel))
								_guictrllistbox_beginupdate($hlistbox)
								_guictrllistbox_resetcontent($hlistbox)
								_guictrllistbox_initstorage($hlistbox, 100, 4096)
								_guictrllistbox_endupdate($hlistbox)
								If IsArray($groupsidrh1) = 1 Then
									For $z = 1 To $groupsidrh1[0]
										$groupsidrh2 = StringSplit($groupsidrh1[$z], ",")
										$groupsidrh3 = $groupsidrh2[1]
										$groupsidrh3 = StringTrimLeft($groupsidrh3, 3)
										_guictrllistbox_addstring($hlistbox, $groupsidrh3)
									Next
								EndIf
								If IsArray($groupsidrh1) = 0 Then
									$idcheckboxgroup = 2
									GUIDelete($hgui2)
									MsgBox(0, "Info !", "abandon, pas de Groupes pour: " & $idrh1, 7)
									Guidelete($hGUI)
									Return 0
								EndIf
								Global $hbutton2 = GUICtrlCreateButton("[ Apply ]", 20, 3, 80, 25)
								GUISetState()
								While 1
									Switch GUIGetMsg()
										Case $hbutton2
											$aselected = _guictrllistbox_getselitems($hlistbox)
											If $aselected[0] = 1 Then
												$sitem = " item"
											Else
												$sitem = " items"
											EndIf
											$sitems = ""
											For $i = 1 To $aselected[0]
												$sitems &= _guictrllistbox_gettext($hlistbox, $aselected[$i]) & @CRLF
											Next
											If $aselected[0] <> 0 Then
												$sitems = StringSplit($sitems, @CRLF)
												For $z = 1 To $sitems[0]
													If StringLen($sitems[$z]) <> 0 Then
														Global $ivalue = _ad_removeuserfromgroup($sitems[$z], $idrh1)
														If $ivalue = 1 AND $userexist = 1 Then
															$actiongroups = $actiongroups & $sitems[$z] & " [-] OK ; "
														ElseIf @error = 2 Then
															$actiongroups = $actiongroups & $sitems[$z] & " [ OU cible  ?  ] ; "
														ElseIf @error = 3 AND $test = 1 Then
															$actiongroups = $actiongroups & $sitems[$z] & " [ pas membre  ! ] ; "
														ElseIf @error = 3 AND $test = 0 Then
															$actiongroups = $actiongroups & $sitems[$z] & " [ déjà membre  ! ] ; "
														ElseIf @error >= 4 OR @error = -2147352567 Then
															$actiongroups = $actiongroups & $sitems[$z] & " [ accès refusé ! ] ; "
														EndIf
													EndIf
												Next
												$actiongroups = " Groupes [-]: " & $actiongroups
												GUIDelete($hgui2)
												ToolTip("Synchronizing with AD...", 5, 5)
												Sleep(3000)
												ToolTip("", 5, 5)
												ExitLoop
											Else
												MsgBox(4160, "warning !", "aucune selection... réessayez ou refermez la fenetre [x]")
											EndIf
										Case $gui_event_close
											GUIDelete($hgui2)
											Guidelete($hGUI)
											Return 0
											ExitLoop
									EndSwitch
								WEnd
								GUIDelete($hgui2)
							Case $idcheckboxdesactivesrce = 1 AND StringLen($idrh1) <> 0 AND $userexist = 1
								$defautdc = $defautdcinit
								If _ispressed("10", $hdll) Then
									$t = MsgBox(4, "Supprimer le compte ?", "Voulez-vous supprimer le compte [ " & $idrh1 & " ] ?")
									If $t = 6 Then
										Global $ivalue = _ad_deleteobject($idrh1, _ad_getobjectclass($idrh1))
										If $ivalue = 1 Then
											MsgBox(64, "Info !", "User Srce '" & $idrh1 & "'  supprimé !")
											Guidelete($hGUI)
											Return 0
										ElseIf @error = 1 Then
											MsgBox(64, "warning !", "User Srce '" & $idrh1 & "' n'existe pas !")
										Else
											MsgBox(64, "warning !", "Return code '" & @error & "' from AD !")
										EndIf
									Else
										MsgBox(0, "Abandon !", "Abandon suppression user Srce :  " & $idrh1, 7)
									EndIf
								Else
									Global $ivalue = _ad_disableobject($idrh1)
									If $ivalue = 1 Then
										MsgBox(64, "Info !", "User Srce '" & $idrh1 & "'  désactivé !")
										Guidelete($hGUI)
										Return 0
									ElseIf @error = 1 Then
										MsgBox(64, "warning !", "User Srce '" & $idrh1 & "' n'existe pas !")
									Else
										MsgBox(64, "warning !", "Return code '" & @error & "' from AD !")
									EndIf
								EndIf
							Case $idcheckboxdescription = 1 AND StringLen($idrh1) <> 0 AND $userexist = 1
								$defautdc = $defautdcinit
								$descriptsrce = _ad_getobjectattribute($idrh1, "description")
								$descriptsrce = InputBox("Description " & $idrh1, " ", $descriptsrce)
								If @error = 1 Then
									MsgBox(0, "Info !", " ...aborting !" & @CRLF & @CRLF & $descriptsrce)
									Guidelete($hGUI)
									Return 0
								EndIf
								If $descriptsrce <> "" Then
									_ad_modifyattribute($idrh1, "description", $descriptsrce, 2)
									MsgBox(0, "Info !", "Source " & $idrh1 & ", Description :" & @CRLF & @CRLF & $descriptsrce)
									Guidelete($hGUI)
									Return 0
								Else
									_ad_modifyattribute($idrh1, "description", $descriptsrce, 2)
									MsgBox(0, "Info !", "Source " & $idrh1 & ", Description est maintenant vide !")
									Guidelete($hGUI)
									Return 0
								EndIf
								Guidelete($hGUI)
								Return 0
							Case $idcheckboxdriveremove = 1 AND StringLen($idrh1) <> 0 AND $userexist = 1
								$defautdc = $defautdcinit
								Global $drivesidrh1 = _ad_getobjectproperties($idrh1, "LaPoste00-NetworkDrive")
								_arraydelete($drivesidrh1, 0)
								If UBound($drivesidrh1) = 0 Then
									MsgBox(0, "Info !", "abandon, pas de Drives pour: " & $idrh1, 7)
									Guidelete($hGUI)
									Return 0
								EndIf
								$drivesidrh1 = _arraytostring($drivesidrh1, "|")
								$drivesidrh1 = StringReplace($drivesidrh1, "LaPoste00-NetworkDrive|", "")
								$drivesidrh1 = StringSplit($drivesidrh1, @CRLF)
								Global $aselected[1]
								Global $hgui2 = GUICreate("Selection drives [-] pour '" & $idrh1 & "' , puis validez avec bouton [ Apply ] !  ", 1000, 500)
								Global $hlistbox = _guictrllistbox_create($hgui2, "String upon creation", 10, 30, 980, 470, BitOR($ws_border, $lbs_hasstrings, $ws_vscroll, $lbs_multiplesel))
								_guictrllistbox_beginupdate($hlistbox)
								_guictrllistbox_resetcontent($hlistbox)
								_guictrllistbox_initstorage($hlistbox, 100, 4096)
								_guictrllistbox_endupdate($hlistbox)
								If IsArray($drivesidrh1) = 1 Then
									For $z = 1 To $drivesidrh1[0]
										If $drivesidrh1[$z] <> "" Then
											_guictrllistbox_addstring($hlistbox, $drivesidrh1[$z])
										EndIf
									Next
								EndIf
								If IsArray($drivesidrh1) = 0 Then
									GUIDelete($hgui2)
									MsgBox(0, "Info !", "abandon, pas de Drives pour: " & $idrh1, 7)
									Guidelete($hGUI)
									Return 0
								EndIf
								Global $hbutton2 = GUICtrlCreateButton("[ Apply ]", 20, 3, 80, 25)
								GUISetState()
								While 1
									Switch GUIGetMsg()
										Case $hbutton2
											$aselected = _guictrllistbox_getselitems($hlistbox)
											If $aselected[0] = 1 Then
												$sitem = " item"
											Else
												$sitem = " items"
											EndIf
											$sitems = ""
											For $i = 1 To $aselected[0]
												$sitems &= _guictrllistbox_gettext($hlistbox, $aselected[$i]) & @CRLF
											Next
											If $aselected[0] <> 0 Then
												$sitems = StringSplit($sitems, @CRLF)
												$drivesidrh1 = _arraytostring($drivesidrh1, @CRLF)
												For $z = 1 To $sitems[0]
													If StringLen($sitems[$z]) <> 0 Then
														$drivesidrh1 = StringReplace($drivesidrh1, $sitems[$z], "")
														$actiongroups = $actiongroups & $sitems[$z] & "|"
													EndIf
												Next
												$drivesidrh1 = StringSplit($drivesidrh1, @CRLF)
												If StringTrimRight($actiongroups, 1) = "|" Then
													$actiongroups = StringTrimRight($actiongroups, 1)
												EndIf
												Dim $drivesidrhsrce[0]
												For $z = 1 To $drivesidrh1[0]
													If StringLen($drivesidrh1[$z]) > 4 Then
														_arrayadd($drivesidrhsrce, $drivesidrh1[$z])
														ReDim $drivesidrhsrce[UBound($drivesidrhsrce)]
													Else
													EndIf
												Next
												$actiongroups = " [-] drives: " & $actiongroups
												Global $ivalue = _ad_modifyattribute($idrh1, "LaPoste00-NetworkDrive", $drivesidrhsrce, 2)
												If $ivalue = 1 Then
													GUIDelete($hgui2)
													ToolTip("synchronising AD", 5, 5)
													Sleep(3000)
													ToolTip("", 5, 5)
													ExitLoop
												ElseIf @error = 1 Then
												Else
												EndIf
											EndIf
										Case $gui_event_close
											GUIDelete($hgui2)
											Guidelete($hGUI)
											Return 0
											ExitLoop
									EndSwitch
								WEnd
							Case $idcheckboxscan = 1 AND StringLen($idrh1) <> 0 AND $userexist = 1
								$defautdc = $defautdcinit
								$descriptsrce = _ad_getobjectattribute($idrh1, "Comment")
								$descriptsrce = InputBox("Comment " & $idrh1, " ", $descriptsrce)
								If @error = 1 Then
									MsgBox(0, "Info !", " ...abandon Comment !" & @CRLF & @CRLF & $descriptsrce)
									Guidelete($hGUI)
									Return 0
								Else
									_ad_modifyattribute($idrh1, "Comment", $descriptsrce, 2)
								EndIf
							Case $idcheckboxdirective = 1 AND StringLen($idrh1) <> 0 AND $userexist = 1
								$defautdc = $defautdcinit
								directives()
								Guidelete($hGUI)
							    Return 0

							Case $idcheckboxdirectiveDSIBA =1 AND StringLen($idrh1) <> 0 AND $userexist = 1
								$defautdc = $defautdcinit
								directives_DSIBA()
								Guidelete($hGUI)
							    Return 0

							Case $idcheckboxlbpai = 1 AND StringLen($idrh1) <> 0 AND $userexist = 1
								$defautdc = $defautdcinit
								lbpai_illiade()
								Guidelete($hGUI)
							    Return 0
							Case $idgroup2 <> "" AND StringInStr($idgroup2, ";") AND StringLen($idrh1) <> 0 AND $userexist = 1
								$defautdc = $defautdcinit
								$idgroup2 = StringSplit($idgroup2, ";")
								If IsArray($idgroup2) Then
									$idgroup2result = ""
									$actiongroups = ""
									$actiongrp = 1
									If _ispressed("10", $hdll) Then
										$actiongrp = 0
									Else
										$actiongrp = 1
									EndIf
									For $z = 1 To UBound($idgroup2) - 1
										If $actiongrp = 0 Then
											Global $ivalue = _ad_removeuserfromgroup($idgroup2[$z], $idrh1)
										Else
											Global $ivalue = _ad_addusertogroup($idgroup2[$z], $idrh1)
										EndIf
										If NOT $idgroup2[$z] = "" Then
											Select
												Case $actiongrp = 0
													If $ivalue = 1 AND $actiongrp = 0 Then
														$idgroup2result = $idgroup2result & $idrh1 & " [-] OK,  '" & $idgroup2[$z] & "'" & " ;  [" & _ad_getobjectattribute($idgroup2[$z], "distinguishedName") & "]" & @CRLF
														$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  :" & $idrh1 & "   '" & $idgroup2[$z] & "'" & "[-] OK  ;  [" & _ad_getobjectattribute($idgroup2[$z], "distinguishedName") & "]" & @CRLF
												    _GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
													_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idrh1 & "   '" & $idgroup2[$z] & "'" & "[-] OK  ;  [" & _ad_getobjectattribute($idgroup2[$z], "distinguishedName") & "]" & @CRLF);, 1)
													ElseIf @error = 1 Then
														$idgroup2result = $idgroup2result & "Groupe '" & $idgroup2[$z] & "' n'existe pas !" & @CRLF
														$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  :	" & $idgroup2[$z] & "	groupe n'existe pas !"
												    _GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
													_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idgroup2[$z] & "	groupe n'existe pas !" & @CRLF);, 1)
													_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
													ElseIf @error = 2 Then
														MsgBox(64, "", "Idrh source '" & $idrh1 & "' inexistant !")
													ElseIf @error = 3 Then
														$idgroup2result = $idgroup2result & $idrh1 & " n'est deja pas membre de '" & $idgroup2[$z] & "'" & @CRLF
														$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "	:  " & $idrh1 & ",  n'est deja pas membre de :  " & $idgroup2[$z]
													_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idrh1 & ",  n'est deja pas membre de :  " & $idgroup2[$z] & @CRLF);, 1)
													_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
													Else
														$idgroup2result = $idgroup2result & $idrh1 & " [X] '" & $idgroup2[$z] & "' " & "Return code '" & @error & "' from Active Directory" & @CRLF
														$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "	:	code erreur:  " & @error & " !  (accès refusé...)"
												    _GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
													_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "	:	code erreur:  " & @error & " !  (accès refusé...)" & @CRLF);, 1)
													_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
													EndIf
												Case $actiongrp = 1
													If $ivalue = 1 AND $actiongrp = 1 Then
														$idgroup2result = $idgroup2result & $idrh1 & " [+] OK,   '" & $idgroup2[$z] & "'" & " ;  [" & _ad_getobjectattribute($idgroup2[$z], "distinguishedName") & "]" & @CRLF
														$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  :" & $idrh1 & "	  '" & $idgroup2[$z] & "'" & "[+] OK  ;  [" & _ad_getobjectattribute($idgroup2[$z], "distinguishedName") & "]" & @CRLF
												    _GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
													_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idrh1 & "	  '" & $idgroup2[$z] & "'" & "[+] OK  ;  [" & _ad_getobjectattribute($idgroup2[$z], "distinguishedName") & "]" & @CRLF);, 1)
													ElseIf @error = 1 Then
														$idgroup2result = $idgroup2result & "Group '" & $idgroup2[$z] & "' n'existe pas !" & @CRLF
														$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  :	" & $idgroup2[$z] & "	groupe n'existe pas !"
												    _GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
													_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idgroup2[$z] & "	groupe n'existe pas !" & @CRLF);, 1)
													_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
													ElseIf @error = 2 Then
														MsgBox(64, "", "Idrh source '" & $idrh1 & "' inexistant !")
													ElseIf @error = 3 Then
														$idgroup2result = $idgroup2result & $idrh1 & " est deja membre de :  '" & $idgroup2[$z] & "'" & @CRLF
														$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & ",  est deja membre de :  " & $idgroup2[$z]
													_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF &  $idrh1 & ",  est deja membre de :  " & $idgroup2[$z] & @CRLF);, 1)
													_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
													Else
														$idgroup2result = $idgroup2result & $idrh1 & " [X] '" & $idgroup2[$z] & "' " & "Return code '" & @error & "' from Active Directory" & @CRLF
														$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  :  code erreur:  " & @error & " !  (accès refusé...)"
												    _GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
													_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  :  code erreur:  " & @error & " !  (accès refusé...)" & @CRLF);, 1)
													_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
													EndIf
											EndSelect
										EndIf
									Next
									$actiongroups = $idgroup2result
									If $idgroup2result <> "" Then
										If $actiongrp = 0 Then
											$actiongroups = "Retirer Multiple Groupes de user Srce:    " & $idrh1 & @CRLF & "--------------------------" & @CRLF & $actiongroups
										Else
											$actiongroups = "Rajouter Multiple Groupes pour user Srce: " & $idrh1 & @CRLF & "--------------------------" & @CRLF & $actiongroups
										EndIf
									EndIf
								Else
									MsgBox(0, "Info !", "abandon, pas de Groupes à rajouter pour user Srce: " & $idrh1, 7)
								EndIf
						EndSelect

					#EndRegion Select Case

					 If $idcheckboxprolonge = 1 AND NOT StringInStr(StringLeft($idrh1, 1), "x") Then
						$t = MsgBox(4, "Expiration date ? ", "Le compte [" & $idrh1 & "] n'est pas un compte en 'x' , souhaitez-vous mettre une date d'expiration ?" & @CRLF & "     ex:  pour les CF -> oui !")
						If $t = 6 Then
							Global $date_expire = ""
							$date_expire = _ad_getobjectproperties($idrh1, "accountExpires")
							$date_expire = $date_expire[1][1]
							$date_expire = _dateadd("d", -1, $date_expire)
							If $date_expire = "0000/00/00 00:00:00" OR $date_expire = "1601/01/01 00:00:00" Then
								$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "warning !  , " & "Le compte '" & $idrh1 & "' , n'expire jamais..."
								_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "warning !  , " & "Le compte '" & $idrh1 & "' , n'expire jamais..." & @CRLF);, 1)
							EndIf
							;$date_expire = _dateadd("d", -1, $date_expire)
							$hguicalendar = GUICreate("compte " & $idrh1, 300, 100, 350, 120)
							Local $iddate = GUICtrlCreateDate($date_expire, 30, 30, 100, 20, $dts_shortdateformat)
							Local $idbtn = GUICtrlCreateButton("OK", 120, 65, 60, 20)
							GUISetState(@SW_SHOW)
							While 1
								Switch GUIGetMsg()
									Case $gui_event_close
										GUIDelete($hguicalendar)
										Guidelete($hGUI)
										Return 0
									Case $idbtn
										Global $dateprolonge = GUICtrlRead(($iddate))
										GUIDelete($hguicalendar)
										ExitLoop
								EndSwitch
							WEnd
							$joursd = StringSplit($dateprolonge, "/")
							Global $jourd = $joursd[1]
							Global $moisd = $joursd[2]
							Global $anneed = $joursd[3]
							$dated = $anneed & "-" & $moisd & "-" & $jourd & " 00:00:00"
							$dated = _dateadd("d", 1, $dated)
							$ivalue = _ad_setaccountexpire($idrh1, $dated)
							Sleep(500)
						Else
							MsgBox(0, "no expiration date !", "[" & $idrh1 & "] : pas de date d'expiration ...", 7)
							$date_expire = "01/01/1970 00:00:00"
							$ivalue = _ad_setaccountexpire($idrh1, $date_expire)
						EndIf
					EndIf
					 If StringLen($idrh1) <> 0 Then
						Global $drivesidrh1 = _ad_getobjectproperties($idrh1, "LaPoste00-NetworkDrive")
						_arraydelete($drivesidrh1, 0)
						Global $cpydrives1 = $drivesidrh1
						$drivesidrh1 = _arraytostring($drivesidrh1)
						$drivesidrh1 = StringReplace($drivesidrh1, "LaPoste00-NetworkDrive|", "")
						Global $commentidrh1 = _ad_getobjectproperties($idrh1, "Comment")
						_arraydelete($commentidrh1, 0)
						$commentidrh1 = _arraytostring($commentidrh1)
						$commentidrh1 = StringReplace($commentidrh1, "Comment|", "")
						Global $printersidrh1 = _ad_getobjectproperties($idrh1, "LaPoste00-Printer")
						_arraydelete($printersidrh1, 0)
						$printersidrh1 = _arraytostring($printersidrh1)
						$printersidrh1 = StringReplace($printersidrh1, "LaPoste00-Printer|", "")
						$idcheckboxairwatch = BitAND(GUICtrlRead($idcheckboxairwatch), $gui_checked)
						$idcheckboxbpe = BitAND(GUICtrlRead($idcheckboxbpe), $gui_checked)
						$mail = _ad_getobjectattribute($idrh1, "mail")
						Local $setaw = 0
						Local $bpnn = "BP"
						If @error Then
						Else
							_arraydelete($drivesidrh1, 0)
						EndIf
						#Region AW
							If $idcheckboxairwatch = 1 Then
								$mail = InputBox("setting e-Mail ...", " ", $mail)
								If @error = 1 Then
									MsgBox(0, "Info !", "Airwatch abandon !" & @CRLF & @CRLF & $mail)
									Guidelete($hGUI)
									Return 0
								EndIf
								_ad_modifyattribute($idrh1, "mail", $mail, 2)
								If $isdct = 0 Then
									MsgBox(0, "Airwatch DCT ?", "AW n'est valable que pour domaine DCT" & @CRLF & "  mail est tout de meme defini à: " & $mail & @CRLF & "  pour user Srce: " & $idrh1)
								Else
									Global $ousourcedir = _ad_getobjectattribute($idrh1, "distinguishedName")
									$ousourcedir = StringSplit($ousourcedir, ",")
									$ousourcedir = $ousourcedir[3]
									$ousourcedir = StringTrimLeft($ousourcedir, 3)
									If StringInStr($ousourcedir, "GAUB") = 1 Then
										$bpnn = InputBox("OU ? (BPLY, BPGR, etc...)", "Actuel user Srce OU= " & $ousourcedir & " (pour: " & $idrh1 & ")" & @CRLF & "par defaut BPNA, vous pouvez le modifier !", "BPNA")
									Else
										$bpnn = InputBox("OU ? (BPLY, BPGR, etc...)", "Actuel user Srce OU= " & $ousourcedir & " (pour: " & $idrh1 & ")", $ousourcedir)
									EndIf
									If @error = 1 Then
										MsgBox(0, "Info !", "Airwatch ...abandon !" & @CRLF & @CRLF & $mail)
										Guidelete($hGUI)
										Return 0
									EndIf
									If $bpe = 0 Then
										_ad_addusertogroup("RG-" & $bpnn & "_AW Commun Manager", $idrh1)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  rajout de @mail et aussi du groupe Airwatch, pour: " & $idrh1 & "  " & "RG-" & $bpnn & "_AW Commun Manager"
									_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  rajout @mail et du groupe Airwatch, pour: " & $idrh1 & "  " & "RG-" & $bpnn & "_AW Commun Manager" & @CRLF);, 1)
									EndIf
									If $bpe = 1 Then
										Dim $arrayaw[0]
										_arrayadd($arrayaw, "RG-" & $bpnn & "_AW BPE")
										_arrayadd($arrayaw, "RG-" & $bpnn & "_AW BPE Restreint")
										_arrayadd($arrayaw, "RG-" & $bpnn & "_AW COMEX")
										_arrayadd($arrayaw, "RG-" & $bpnn & "_AW COMEX GROUPE")
										_arrayadd($arrayaw, "RG-" & $bpnn & "_AW COMMUN Manager")
										_arrayadd($arrayaw, "RG-" & $bpnn & "_AW COMMUN Standard")
										_arrayadd($arrayaw, "RG-" & $bpnn & "_AW DASTDR")
										_arrayadd($arrayaw, "RG-" & $bpnn & "_AW EBR Formateurs")
										_arrayadd($arrayaw, "RG-" & $bpnn & "_AW MFB")
										_arrayadd($arrayaw, "RG-" & $bpnn & "_AW SENSIBLE")
										Global $listeairwatch = _arraytostring($arrayaw, "|")
										Local $hguiaw = GUICreate("Choix du groupe AW ?", 300, 50)
										Local $idcomboboxaw = GUICtrlCreateCombo("", 10, 10, 275, 20)
										GUICtrlSetData($idcomboboxaw, $listeairwatch, "RG-" & $bpnn & "_AW BPE")
										GUISetState(@SW_SHOW, $hguiaw)
										Local $scomboreadaw = ""
										Local $idmsg = GUIGetMsg()
										While ($idmsg <> $gui_event_close)
											$idmsg = GUIGetMsg()
											$scomboreadaw = GUICtrlRead($idcomboboxaw)
										WEnd
										$scomboreadaw = GUICtrlRead($idcomboboxaw)
										MsgBox(0, "info AW !", " Choix AW: " & @CRLF & @CRLF & "> " & $scomboreadaw)
										_ad_addusertogroup($scomboreadaw, $idrh1)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  rajout @mail et du groupe Airwatch, pour: " & $idrh1 & "  " & $scomboreadaw
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  rajout @mail et du groupe Airwatch, pour: " & $idrh1 & "  " & $scomboreadaw & @CRLF);, 1)
										GUIDelete($hguiaw)
									EndIf
								EndIf
							EndIf
						#EndRegion AW
						$idcheckboxgroup = BitAND(GUICtrlRead($idcheckboxgroup), $gui_checked)
						$idcheckboxrole = BitAND(GUICtrlRead($idcheckboxrole), $gui_checked)
						If $idcheckboxgroup = 1 AND $idcheckboxrole = 0 Then
						   $defautdc = $defautdcinit
							If $isdct = 1 Then
								$unite = InputBox("default OU ? 2 or 4 chr$", "ex: MILY, MITE, BPLY, GAUB, MI ... virtuos" & @CRLF & "?? : scan all OUs...", "MI")
								If StringLen($unite) < 2 OR StringLen($unite) > 4 AND NOT ($unite = "virtuos" OR $unite = "??") Then
								   Guidelete($hGUI)
							    Return 0
								;	ExitLoop
								EndIf
							Else
								$unite = "??"
							EndIf
							If $unite = "??" Then
								$allou = 1
								treeview_affiche()
								treeselect()
							Else
								$allou = 0
							EndIf
							$defautdc = StringReplace($defautdc, ".", ",DC=")
							$defautdc = ",DC=" & $defautdc
							If $allou = 1 Then
								$unite = ""
								$allou = 0
								Global $aous = _ad_getobjectsinou("" & $sout)
							Else
								Global $aous = _ad_getobjectsinou("OU=Groupes,OU=" & $unite & ",OU=" & StringMid($unite, 1, 2) & $defautdc)
							EndIf
							If $unite = "MI" Then
								Global $aous = _ad_getobjectsinou("OU=Groupes,OU=_Support Transverse,OU=MI" & $defautdc)
							EndIf
							If $unite = "virtuos" Then
								Global $aous = _ad_getobjectsinou("OU=FunctionalGroups,OU=_Flexible Workspace Commun" & $defautdc)
							EndIf
							$sout = ""
							$filez = FileOpen(@ScriptDir & "\Groups-AD_OU-" & $unite & ".txt", 2)
							If @error > 0 Then
								MsgBox(64, $adroot, "Pas de Groupes trouvés ! ")
								FileClose($filez)
								Guidelete($hGUI)
								Return 0
							Else
								$filterou = InputBox("filtre des  groupes ?", "vide = tous , ou mots clés: _cbk, _dpi ... ", "")
								Global $aselected[1]
								Global $hgui2 = GUICreate("Selection groupes puis valider avec le bouton [ Apply ] !  filtre:(" & $filterou & ")", 1000, 500)
								Global $hlistbox = _guictrllistbox_create($hgui2, "String upon creation", 10, 30, 980, 470, BitOR($ws_border, $lbs_hasstrings, $ws_vscroll, $lbs_multiplesel))
								_guictrllistbox_beginupdate($hlistbox)
								_guictrllistbox_resetcontent($hlistbox)
								_guictrllistbox_initstorage($hlistbox, 100, 4096)
								_guictrllistbox_endupdate($hlistbox)
								If IsArray($aous) = 1 Then
									For $z = 1 To $aous[0] - 1
										FileWriteLine($filez, $aous[$z] & @CRLF)
										If StringLen($filterou) = 0 Then
											_guictrllistbox_addstring($hlistbox, $aous[$z])
										Else
											If NOT StringInStr($filterou, ";") Then
												If StringInStr($aous[$z], $filterou) Then
													_guictrllistbox_addstring($hlistbox, $aous[$z])
												EndIf
											Else
												Local $filterou2 = StringSplit($filterou, ";")
												For $w = 1 To $filterou2[0]
													If StringInStr($aous[$z], $filterou2[$w]) Then
														_guictrllistbox_addstring($hlistbox, $aous[$z])
													EndIf
												Next
											EndIf
										EndIf
									Next
								EndIf
								FileClose($filez)
								If _ispressed("10", $hdll) Then
								Else
									FileDelete(@ScriptDir & "\Groups-AD_OU-" & $unite & ".txt")
								EndIf
							EndIf
							If IsArray($aous) = 0 Then
								$idcheckboxgroup = 2
								GUIDelete($hgui2)
								MsgBox(0, "Info !", "abandon, pas de groupes trouvés dans l'OU: " & $unite, 7)
								FileClose($filez)
								Guidelete($hGUI)
								Return 0
							EndIf
							Global $hbutton = GUICtrlCreateButton("[ Apply ]", 20, 3, 80, 25)
							GUISetState()
							While 1
								Switch GUIGetMsg()
									Case $hbutton
										$aselected = _guictrllistbox_getselitems($hlistbox)
										If $aselected[0] = 1 Then
											$sitem = " item"
										Else
											$sitem = " items"
										EndIf
										$sitems = ""
										For $i = 1 To $aselected[0]
											$sitems &= _guictrllistbox_gettext($hlistbox, $aselected[$i]) & @CRLF
										Next
										If $aselected[0] <> 0 Then
											$sitems = StringSplit($sitems, @CRLF)
											For $z = 1 To $sitems[0]
												If StringLen($sitems[$z]) <> 0 Then
													$ivalue = _ad_addusertogroup($sitems[$z], $idrh1)
													If $ivalue = 1 Then
														$actiongroups = $actiongroups & $sitems[$z] & " [+] OK ; "
													ElseIf @error = 2 Then
														$actiongroups = $actiongroups & $sitems[$z] & " [  OU cible ?  ] ; "
													ElseIf @error = 3 Then
														$actiongroups = $actiongroups & $sitems[$z] & " [  deja membre ! ] ; "
													ElseIf @error = 3 Then
														$actiongroups = $actiongroups & $sitems[$z] & " [  pas membre !  ] ; "
													ElseIf @error >= 4 OR @error = -2147352567 Then
														$actiongroups = $actiongroups & $sitems[$z] & " [ accès refusé ! ] ; "
													EndIf
												EndIf
											Next
											$actiongroups = " Groupes [+]: " & $actiongroups
											GUIDelete($hgui2)
											ExitLoop
										Else
											MsgBox(4160, "warning !", "aucune selection... reessayez ou fermez la fenetre [x]")
										EndIf
									Case $gui_event_close
										$idcheckboxgroup = 2
										GUIDelete($hgui2)
										ExitLoop
								EndSwitch
							WEnd
						Else
						EndIf
						If $idcheckboxgroup = 1 AND $idcheckboxrole = 0 Then
							ToolTip("synchronizing with AD ...", 5, 5, "")
							Sleep(1800)
							ToolTip("", 5, 5, "")
						EndIf
						If $idcheckboxrole = 1 AND $idcheckboxgroup = 0 Then
						   $defautdc = $defautdcinit
							If $isdct = 0 Then
								MsgBox(0, "Warning !", "Roles valables que pour le domaine DCT !", 7)
								Guidelete($hGUI)
								Return 0
							EndIf
							$unite = InputBox("default OU ? ", "ex: MILY, MITE, BPLY, GAUB, ... " & @CRLF & "?? : scan all OUs...", "GAUB")
							If StringLen($unite) < 2 OR StringLen($unite) > 4 AND NOT ($unite = "virtuos" OR $unite = "??") Then
							   Guidelete($hGUI)
							    Return 0
								;ExitLoop
							EndIf
							If $unite = "??" Then
								$allou = 1
								treeview_affiche()
								treeselect()
							Else
								$allou = 0
							EndIf
							$defautdc = StringReplace($defautdc, ".", ",DC=")
							$defautdc = ",DC=" & $defautdc
							If $allou = 1 Then
								$unite = ""
								$allou = 0
								Global $aous = _ad_getobjectsinou("" & $sout)
							Else
								Global $aous = _ad_getobjectsinou("OU=_Roles Fonctionnels,OU=Groupes,OU=" & $unite & ",OU=" & StringMid($unite, 1, 2) & $defautdc)
							EndIf
							If $unite = "MI" Then
								Global $aous = _ad_getobjectsinou("OU=Groupes,OU=_Support Transverse,OU=MI" & $defautdc)
							EndIf
							If $unite = "virtuos" Then
								Global $aous = _ad_getobjectsinou("OU=FunctionalGroups,OU=_Flexible Workspace Commun" & $defautdc)
							EndIf
							$sout = ""
							$filez = FileOpen(@ScriptDir & "\Roles-AD_OU-" & $unite & ".txt", 2)
							If @error > 0 Then
								MsgBox(64, $adroot, "No Roles found ! ")
								FileClose($filez)
								Guidelete($hGUI)
								Return 0
							Else
								$filterou = InputBox("filtrer les Roles ?", "vide = tous , ou mot clé: Bancaire ... ", "")
								Global $aselected[1]
								Global $hgui2 = GUICreate("Selection Role puis valider avec le bouton [ Apply ] !  filtre:(" & $filterou & ")", 1000, 500)
								Global $hlistbox = _guictrllistbox_create($hgui2, "String upon creation", 10, 30, 980, 470, BitOR($ws_border, $lbs_hasstrings, $ws_vscroll, $lbs_multiplesel))
								_guictrllistbox_beginupdate($hlistbox)
								_guictrllistbox_resetcontent($hlistbox)
								_guictrllistbox_initstorage($hlistbox, 100, 4096)
								_guictrllistbox_endupdate($hlistbox)
								If IsArray($aous) = 1 Then
									For $z = 1 To $aous[0] - 1
										FileWriteLine($filez, $aous[$z] & @CRLF)
										If StringLen($filterou) = 0 Then
											_guictrllistbox_addstring($hlistbox, $aous[$z])
										Else
											If NOT StringInStr($filterou, ";") Then
												If StringInStr($aous[$z], $filterou) Then
													_guictrllistbox_addstring($hlistbox, $aous[$z])
												EndIf
											Else
												Local $filterou2 = StringSplit($filterou, ";")
												For $w = 1 To $filterou2[0]
													If StringInStr($aous[$z], $filterou2[$w]) Then
														_guictrllistbox_addstring($hlistbox, $aous[$z])
													EndIf
												Next
											EndIf
										EndIf
									Next
								EndIf
								FileClose($filez)
								If _ispressed("10", $hdll) Then
								Else
									FileDelete(@ScriptDir & "\Roles-AD_OU-" & $unite & ".txt")
								EndIf
							EndIf
							If IsArray($aous) = 0 Then
								$idcheckboxgroup = 2
								GUIDelete($hgui2)
								MsgBox(0, "Info !", "abandon, aucun Roles trouvé dans l'OU: " & $unite, 7)
								FileClose($filez)
								Guidelete($hGUI)
								Return 0
							EndIf
							Global $hbutton = GUICtrlCreateButton("[ Apply ]", 20, 3, 80, 25)
							GUISetState()
							While 1
								Switch GUIGetMsg()
									Case $hbutton
										$aselected = _guictrllistbox_getselitems($hlistbox)
										If $aselected[0] > 1 Then
											MsgBox(4160, "warning !", "multiples 'Roles' choisis ! choisissez un seul role ou refermer la fenetre [x]")
										Else
											If $aselected[0] = 1 Then
												$sitem = " item"
											Else
												$sitem = " items"
											EndIf
											If $aselected[0] = 1 Then
												$sitems = ""
												For $i = 1 To $aselected[0]
													$sitems &= _guictrllistbox_gettext($hlistbox, $aselected[$i]) & @CRLF
												Next
												If $aselected[0] <> 0 Then
													$sitems = StringSplit($sitems, @CRLF)
													For $z = 1 To $sitems[0]
														If StringLen($sitems[$z]) <> 0 Then
															$ivalue = _ad_addusertogroup($sitems[$z], $idrh1)
															If $ivalue = 1 Then
																$actiongroups = $actiongroups & $sitems[$z] & " [+] OK ; "
															ElseIf @error = 2 Then
																$actiongroups = $actiongroups & $sitems[$z] & " [  OU cible ? ] ; "
															ElseIf @error = 3 Then
																$actiongroups = $actiongroups & $sitems[$z] & " [  deja membre ! ] ; "
															ElseIf @error = 3 Then
																$actiongroups = $actiongroups & $sitems[$z] & " [  pas membre !  ] ; "
															ElseIf @error >= 4 OR @error = -2147352567 Then
																$actiongroups = $actiongroups & $sitems[$z] & " [ accès refusé ! ] ; "
															EndIf
														EndIf
													Next
													$actiongroups = "  Role [+]: " & $actiongroups
													$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  roles rajoutés à " & $idrh1 & " :" & @CRLF & $actiongroups
													_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  roles rajoutés à " & $idrh1 & " :" & @CRLF & $actiongroups & @CRLF);, 1)
													GUIDelete($hgui2)
													ExitLoop
												Else
													MsgBox(4160, "warning !", "aucune selection... reessayez ou refermez la fenetre [x]")
												EndIf
											Else
											EndIf
										EndIf
									Case $gui_event_close
										$idcheckboxgroup = 2
										GUIDelete($hgui2)
										ExitLoop
								EndSwitch
							WEnd
						Else
						EndIf
						If $idcheckboxgroup = 0 AND $idcheckboxrole = 1 Then
							ToolTip("synchronizing AD ...", 5, 5, "")
							Sleep(1800)
							ToolTip("", 5, 5, "")
						EndIf
						Global $groupsidrh1 = _ad_getusergroups($idrh1)
						If @error > 0 Then
						Else
							_arraysort($groupsidrh1, 0, 1)
							_arraydelete($groupsidrh1, 0)
							Global $groupidrh_final = ""
							Global $groupidrh1_add = $groupsidrh1
							For $k = 0 To UBound($groupsidrh1) - 1
								$groupidrh = $groupsidrh1[$k]
								$groupidrh = StringSplit($groupidrh, ",")
								$groupidrh_int = $groupidrh[1]
								$groupidrh_int = StringReplace($groupidrh_int, "CN=", "")
								$groupidrh_final = $groupidrh_final & $groupidrh_int & @CRLF
							Next
							$groupsidrh1 = _arraytostring($groupsidrh1)
						EndIf
						If $userexist = 1 Then
							$accountexpire = _ad_getobjectproperties($idrh1, "accountExpires")
							$accountexpire = $accountexpire[1][1]
							If $accountexpire = "0000/00/00 00:00:00" OR $accountexpire = "1601/01/01 00:00:00" Then
								$accountexpire = "Le compte n'expire jamais..."
								If StringInStr(StringLeft($idrh1, 1), "x") And $idcheckboxprolonge = 1 Then
									$t = MsgBox(4, "Warning ! ", "Le compte [" & $idrh1 & "], n'a pas de date d'expiration !" & @CRLF & @CRLF & @CRLF & "  (Oui): mettre une date d'expiration" & @CRLF & @CRLF & "  (Non): laisser tel quel")
									If $t = 6 Then
										$hguicalendar = GUICreate("compte " & $idrh1, 300, 100, 350, 120)
										Local $iddate = GUICtrlCreateDate("", 30, 30, 100, 20, $dts_shortdateformat)
										Local $idbtn = GUICtrlCreateButton("OK", 120, 65, 60, 20)
										GUISetState(@SW_SHOW)
										While 1
											Switch GUIGetMsg()
												Case $gui_event_close
													GUIDelete($hguicalendar)
													Guidelete($hGUI)
													Return 0
												Case $idbtn
													Global $dateprolonge = GUICtrlRead(($iddate))
													GUIDelete($hguicalendar)
													ExitLoop
											EndSwitch
										WEnd
										$joursd = StringSplit($dateprolonge, "/")
										Global $jourd = $joursd[1]
										Global $moisd = $joursd[2]
										Global $anneed = $joursd[3]
										$dated = $anneed & "-" & $moisd & "-" & $jourd & " 00:00:00"
										$dated = _dateadd("d", 1, $dated)
										$ivalue = _ad_setaccountexpire($idrh1, $dated)
										Sleep(500)
										$dated = _dateadd("d", -1, $dated)
										$accountexpire = $dated
									EndIf
								Else
								EndIf
							Else
								$accountexpire = _dateadd("d", -1, $accountexpire)
							EndIf
							$mail = _ad_getobjectattribute($idrh1, "mail")
							If $mail = "" Then
								$mail = " adresse mail non renseignée sur l'AD !"
							EndIf
							$checkbuttonpwdreset = BitAND(GUICtrlRead($idcheckboxpwdreset), $gui_checked)
							If $checkbuttonpwdreset = 1 Then
								$password = InputBox("password reset", "mot de passe par defaut: W!nd0ws10" & @CRLF & "Vous pouvez modifier le mdp par defaut", "W!nd0ws10")
								If @error = 1 Then
									MsgBox(0, "", "Abandon 'reset pwd' pour user Srce !", 3)
									Guidelete($hGUI)
									Return 0
								EndIf
								Local $achanger = 1
								$t = MsgBox(4, "reinit du mot de passe", "Le mot de passe reinit. doit-il etre changé" & @CRLF & @CRLF & "à la prochaine ouverture session (O/N) ?")
								If $t = 6 Then
									$achanger = 1
								ElseIf $t = 7 Then
									$achanger = 0
								EndIf
								$checkresetpwd = _ad_setpassword($idrh1, $password, $achanger)
								If $checkresetpwd = 1 Then
									If $achanger = 1 Then
										$pwdreset = @CRLF & @CRLF & "  Compte Srce: " & $idrh1 & @CRLF & "     -> reinit. mdp à:  " & $password & "   (Le mot de passe devra etre modifié à l'ouverture de session !)"
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $pwdreset
									_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $pwdreset & @CRLF);, 1)
									Else
										$pwdreset = @CRLF & @CRLF & "  Compte Srce: " & $idrh1 & @CRLF & "     -> reinit. mdp à:  " & $password & "   (Le mot de passe ne devra pas etre changé à l'ouverture de session !)"
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $pwdreset
									_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $pwdreset & @CRLF);, 1)
									EndIf
								Else
									MsgBox(0, "warning !", "Réinit. password : Acces refusé , pour user:  " & $idrh1)
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "" & "warning !, " & "Reset mot de passe windows, acces refusé pour:  " & $idrh1
									_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
									_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "warning !, " & "Reset mot de passe windows, acces refusé pour:  " & $idrh1 & @CRLF);, 1)
									_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
									Guidelete($hGUI)
									Return 0
								EndIf
							Else
								$pwdreset = ""
							EndIf
							If _ad_getobjectattribute($idrh1, "description") <> "" Then
								$description = "  Description:  " & _ad_getobjectattribute($idrh1, "description") & @CRLF
							Else
								$description = ""
							EndIf
							Global $proprietaire = ""
							Global $tmpowner = _ad_getobjectproperties($idrh1, "nTSecurityDescriptor")
							_arraydelete($tmpowner, 0)
							If IsArray($tmpowner) Then
								$tmpowner = $tmpowner[0][1]
								$tmpowner = StringSplit($tmpowner, ":")
								$tmpowner = $tmpowner[4]
								$tmpowner = StringReplace($tmpowner, ", Revision", "")
								$proprietaire = $tmpowner
								If StringInStr($proprietaire, "S-") Then
									$proprietaire = "SIDkey ! (compte AD supprimé ?)"
								EndIf
								$proprietaire = StringReplace($proprietaire, $racinead & "\", "")
								If StringInStr($proprietaire, "Admins") Then
									$proprietaire = "Administrateur du domaine"
								EndIf
								If StringLen($proprietaire) = 7 Then
									$proprietaire = $proprietaire & ""
								EndIf
							EndIf
							Global $lastchanged = _ad_getobjectproperties($idrh1, "whenChanged")
							If IsArray($lastchanged) Then
								_arraydelete($lastchanged, 0)
								If UBound($lastchanged) <> "" Then
									$lastchanged = $lastchanged[0][1]
								Else
									$lastchanged = "?"
								EndIf
							EndIf
							Global $createdsrce = _ad_getobjectproperties($idrh1, "whenCreated")
							If IsArray($createdsrce) Then
								_arraydelete($createdsrce, 0)
								If UBound($createdsrce) <> "" Then
									$createdsrce = $createdsrce[0][1]
								Else
									$createdsrce = "?"
								EndIf
							EndIf
							$pwdinfo = ""
							If $scomboreaddirectives <> "" Then
								If $dirapplied = 1 Then
									$outputidrh1 = "	   Directive appliquée:   " & $scomboreaddirectives & "	,  OU d'origine= " & $ousourcedir & @CRLF
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & $outputidrh1
									_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $outputidrh1 & @CRLF);, 1)
								Else
									$outputidrh1 = "	   Directive Non appliquée !  (rollback)  " & $scomboreaddirectives & "	,  OU d'origine= " & $ousourcedir & @CRLF
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & $outputidrh1
									_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $outputidrh1 & @CRLF);, 1)
								EndIf
							Else
								$outputidrh1 = ""
							 EndIf
							   if $isDCT=1 then ;domaine DCT=1 (Drives, comment, printers, groups)
							   $outputidrh1 = $outputidrh1 & "  Object:  " & _ad_getobjectattribute($idrh1, "distinguishedName") & "  | Expiration ?:  " & $accountexpire & @CRLF & "									Proprietaire:     " & $proprietaire & @CRLF & "									compte crée  le:  " & $createdsrce & @CRLF & "									mdp  changé  le:  " & $lastchanged & @CRLF & @CRLF & "  email :     " & $mail & $pwdreset & @CRLF & $pwdinfo & @CRLF & $description & @CRLF & "  Drives of user: " & $idrh1 & @CRLF & $drivesidrh1 & @CRLF & @CRLF & "  Comment 'Mes scan' of user: " & $idrh1 & @CRLF & $commentidrh1 & @CRLF & @CRLF & "  Printers for user: " & $idrh1 & @CRLF & $printersidrh1 & @CRLF & @CRLF & "  Groups for user: " & $idrh1 & @CRLF & $groupidrh_final
							   Else ; (Groups)
						       $outputidrh1 = $outputidrh1 & "  Object:  " & _ad_getobjectattribute($idrh1, "distinguishedName") & "  | Expiration ?:  " & $accountexpire & @CRLF & "									Proprietaire:     " & $proprietaire & @CRLF & "									compte crée  le:  " & $createdsrce & @CRLF & "									mdp  changé  le:  " & $lastchanged & @CRLF & @CRLF & "  email :     " & $mail & $pwdreset & @CRLF & $pwdinfo & @CRLF & $description & @CRLF & @CRLF & "  Groups for user: " & $idrh1 & @CRLF & $groupidrh_final
							   EndIf
							$stateidrh1 = _ad_isobjectdisabled($idrh1)
							If $stateidrh1 = 1 Then
								$stateidrh1 = $idrh1 & " , le compte est desactivé !"
							Else
								$stateidrh1 = $idrh1 & " , le compte est activé !   "
							EndIf
							$outputidrh1 = $outputidrh1 & @CRLF & "  status du compte:" & @CRLF & $stateidrh1
							$stateid1 = _ad_isobjectlocked($idrh1)
							$compteactive = _ad_isobjectdisabled($idrh1)
							If $stateid1 = 0 Then
								$stateidrh1 = $idrh1 & " , le compte n'est pas verouillé !"
							ElseIf $stateid1 <> 0 Then
								$stateidrh1 = $idrh1 & " , le compte est verouillé !"
								$t = MsgBox(4, "", "Souhaitez-vous déverouiller le compte source: " & $idrh1 & " ?")
								If $t = 6 Then
									$result = _ad_unlockobject($idrh1)
									If $result = 1 Then
										ToolTip("Compte " & $idrh1 & " déverouillé !", 5, 5, "")
										Sleep(2500)
										ToolTip("", 5, 5, "")
									EndIf
								ElseIf $t = 7 Then
								EndIf
							ElseIf $stateid1 = 1 AND $compteactive = 0 Then
								$stateidrh1 = $idrh1 & "  , compte vérouillé ou état inconnu !"
							EndIf
							$outputidrh1 = $outputidrh1 & @CRLF & @CRLF & "  état du compte:" & @CRLF & $stateidrh1
							$lastiter = $outputidrh1
							If $actiongroups <> "" Then
								$outputidrh1 = $outputidrh1 & @CRLF & @CRLF & "	Modifs: " & $actiongroups & @CRLF & @CRLF
								$lastiter = $lastiter & @CRLF & @CRLF & "	Modifs: " & $actiongroups & @CRLF & @CRLF
							EndIf
							If $actiongroupsar <> "" Then
								$actiongroupsar = " [-] groups: " & $actiongroupsar
								$outputidrh1 = $outputidrh1 & @CRLF & @CRLF & "	Modifs: " & $actiongroupsar & @CRLF & @CRLF
								$lastiter = $lastiter & @CRLF & @CRLF & "	Modifs: " & $actiongroupsar & @CRLF & @CRLF
							EndIf
							If $listmanualdrives <> "" Then
								$outputidrh1 = $outputidrh1 & @CRLF & @CRLF & "	Modifs: [+]  Drive:   " & $listmanualdrives & @CRLF & @CRLF
								$lastiter = $lastiter & @CRLF & @CRLF & "	Modifs: [+] Drive:   " & $listmanualdrives & @CRLF & @CRLF
							EndIf
							If $idcheckboxgroup <> 2 Then
								ClipPut($outputidrh1)
								$historik = $historik & @CRLF & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $outputidrh1
								_GUICtrlRichEdit_AppendText($aff, @CRLF & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $outputidrh1 & @CRLF & @CRLF & "			=====	=====	=====	=====	=====			" & @CRLF & @CRLF);, 1)
							EndIf
							$stateidrh1 = _ad_isobjectdisabled($idrh1)
							If $stateidrh1 = 1 Then
								$t = MsgBox(4, "", "Souhaitez-vous réactiver le compte source: " & $idrh1 & " ?")
								If $t = 6 Then
									$result = _ad_enableobject($idrh1)
									If $result = 1 Then
										ToolTip("Compte " & $idrh1 & " réactivé !", 5, 5, "")
										_GUICtrlRichEdit_SetCharColor($aff, 0x00ff0000)      ;   set color blue (BBGGRR format)
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Compte  [ " & $idrh1 & " ]  réactivé !");, 1)
										_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & " ,  compte réactivé !"
										Sleep(2500)
										ToolTip("", 5, 5, "")
									Else
										ToolTip("Compte " & $idrh1 & "  est resté désactivé  !", 5, 5, "Accès refusé !")
										_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Compte  [ " & $idrh1 & " ]  est resté désactivé !  =>  Accès refusé !");, 1)
										_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & " , compte resté désactivé ! => Accès refusé !"
										Sleep(2500)
										ToolTip("", 5, 5, "")
									EndIf
								ElseIf $t = 7 Then
								EndIf
							EndIf
							  $userexist2 = _ad_objectexists($idrh2)
						    	$checkbuttondrives = BitAND(GUICtrlRead($idcheckboxdrives), $gui_checked)
								$checkbuttongroups = BitAND(GUICtrlRead($idcheckboxgroups), $gui_checked)

							If StringUpper($idrh2) = "" OR StringLen($idrh2) = 0 Then
							   $checkbuttondrives=0
							   $idcheckscancpy=0
							   $checkbuttongroups=0
							   EndIf

							If  $userexist2 = 0 and StringUpper($idrh2) <> "" And $checkbuttondrives=1 Then ;StringUpper($idrh2) = "" OR StringLen($idrh2) = 0 OR
								;Sleep(2500)
								MsgBox(0,"","user Dest [ " & $idrh2 & " ] n'existe pas !",15)
								ToolTip("", 5, 5, "")
							Else
							;	$checkbuttondrives = BitAND(GUICtrlRead($idcheckboxdrives), $gui_checked)
							;	$checkbuttongroups = BitAND(GUICtrlRead($idcheckboxgroups), $gui_checked)
								If $checkbuttondrives = 1 And $isDCT=1 and $userexist2=1 Then
									$cpydrives1 = _arraytostring($cpydrives1, "|")
									$cpydrives1 = StringReplace($cpydrives1, $idrh1, $idrh2)
									$cpydrives1 = StringReplace($cpydrives1, "LaPoste00-NetworkDrive|", "")
									$cpydrives1 = StringSplit($cpydrives1, @CRLF)
									If IsArray($cpydrives1) AND $cpydrives1[1] <> "" Then
										For $i = UBound($cpydrives1) - 1 To 0 Step -1
											If $cpydrives1[$i] = "" Then
												_arraydelete($cpydrives1, $i)
											EndIf
											$drivewspace = ""
											$drivewspace = StringStripWS($cpydrives1[$i], $str_stripleading + $str_striptrailing)
											$drivewspace = StringStripCR($cpydrives1[$i])
											If StringInStr($drivewspace, " ") AND NOT StringInStr($drivewspace, "Mes scan") Then
												MsgBox(0, "Warning !", "  'White space' found in drive (from Idrh Srce): " & $idrh1 & @CRLF & $drivewspace & @CRLF & "  ..removed WS !", 15)
												_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
												_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  White space found in drive (from Idrh Srce): " & $idrh1 & @CRLF & $drivewspace & "  ..removed WS !" & @CRLF & @CRLF);, 1)
												_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
												$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  White space found in drive (from Idrh Srce): " & $idrh1 & @CRLF & $drivewspace & "  ..removed WS !" & @CRLF
												$drivewspace = StringSplit($drivewspace, " ")
												$drivewspace = $drivewspace[1]
											EndIf
											_arrayinsert($cpydrives1, $i, $drivewspace)
											_arraydelete($cpydrives1, $i + 1)
										Next
										_arraydelete($cpydrives1, 0)
									Else
										MsgBox(0, "Warning !", "No drives to copy from Srce:  " & $idrh1 & " !", 7)
										_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Aucun Drive à copier depuis user Srce:  " & $idrh1 & " !" & @CRLF & @CRLF);, 1)
										_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
									EndIf
									If UBound($cpydrives1) <> 0 AND $cpydrives1[1] <> "" Then
										_arraydisplay($cpydrives1, "cpy drives depuis user Srce '" & $idrh1 & "'   vers   '" & $idrh2 & "'")
									EndIf
									$cpydrives1bis = $cpydrives1
									$drivesidrh2 = _ad_getobjectproperties($idrh2, "LaPoste00-NetworkDrive")
									_arraydelete($drivesidrh2, 0)
									$drivesidrh2 = _arraytostring($drivesidrh2)
									$drivesidrh2 = StringReplace($drivesidrh2, "LaPoste00-NetworkDrive|", "")
									$drivesidrh2 = StringSplit($drivesidrh2, @CRLF)
									For $i = UBound($drivesidrh2) - 1 To 0 Step -1
										If $drivesidrh2[$i] = "" Then
											_arraydelete($drivesidrh2, $i)
										EndIf
									Next
									_arraydelete($drivesidrh2, 0)
									$letterused = ""
									For $i = UBound($drivesidrh2) - 1 To 0 Step -1
										If NOT StringInStr($drivesidrh2[$i], "Mes scan;") Then
											$letterused = $letterused & StringLeft($drivesidrh2[$i], 2)
										Else
											$letterused = $letterused & StringLeft($drivesidrh2[$i], 9)
										EndIf
									Next
									$letterused = StringSplit($letterused, ";")
									_arraydelete($letterused, $letterused[0])
									_arraydelete($letterused, 0)
									$doublons = ""
									For $i = UBound($cpydrives1) - 1 To 0 Step -1
										For $j = UBound($letterused) - 1 To 0 Step -1
											If StringInStr(StringLeft($cpydrives1[$i], 1), $letterused[$j]) AND NOT StringInStr($cpydrives1[$i], "Mes scan") Then
												$doublons = $doublons & @CRLF & $cpydrives1[$i]
												_arraydelete($cpydrives1bis, $i)
											ElseIf StringInStr($cpydrives1[$i], "Mes scan") AND StringInStr($letterused[$j], "Mes scan") Then
												$doublons = $doublons & @CRLF & $cpydrives1[$i]
												_arraydelete($cpydrives1bis, $i)
											ElseIf NOT StringInStr($cpydrives1bis, $cpydrives1[$i]) Then
											EndIf
										Next
									Next
									$cpydrives3 = $cpydrives1bis
									$cpydrives1bis = _arraytostring($cpydrives1bis, @CRLF)
									$drivesidrh1 = _arraytostring($drivesidrh1)
									$drivesidrh1 = StringReplace($drivesidrh1, "LaPoste00-NetworkDrive|", "")
									If UBound($cpydrives1) = 0 OR $cpydrives1[1] = "" Then
										MsgBox(0, "warning !", "no Drives pour user Srce: " & $idrh1 & " , Abandon...")
										_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "no Drives pour user Srce: " & $idrh1 & " , abandon..." & @CRLF & @CRLF);, 1)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "no Drives pour user Srce: " & $idrh1 & " , abandon..."
										_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
									ElseIf UBound($cpydrives3) = 0 Then
										MsgBox(0, "warning !", "Tous les lettres Drives sont déjà présents pour user Dest. " & $idrh2 & " !" & " , abandon..." & @CRLF & "déjà utilisés:" & @CRLF & $doublons)
										_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Tous les lettres Drives sont déjà présents pour user Dest. " & $idrh2 & " !" & " , abandon..." & @CRLF & "déjà utilisés:" & @CRLF & $doublons & @CRLF & @CRLF);, 1)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Tous les lettres Drives sont déjà présents pour user Dest. " & $idrh2 & " !" & " , abandon..." & @CRLF & "déjà utilisés:" & @CRLF & $doublons
										_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
									Else
										Global $ivalue = _ad_modifyattribute($idrh2, "LaPoste00-NetworkDrive", $cpydrives3, 3)
										If $ivalue = 1 Then
											If $doublons = "" Then
												MsgBox(0, "Info cpy drives", " [+] drives [OK] pour user Dest " & $idrh2 & "  recopiés depuis user Srce " & $idrh1 & @CRLF & @CRLF & $cpydrives1bis, 7)
												_GUICtrlRichEdit_SetCharColor($aff, 0x00ff0000)      ;   set color blue (BBGGRR format)
												_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & " [+] drives [OK] pour user Dest  " & $idrh2 & "  recopiés depuis user Srce  " & $idrh1 & @CRLF & @CRLF & $cpydrives1bis & @CRLF & @CRLF);, 1)
												_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
												$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  [+] drives [OK] pour user Dest  " & $idrh2 & "  recopiés depuis user Srce   " & $idrh1 & @CRLF & @CRLF & $cpydrives1bis
											Else
												MsgBox(0, "Warning cpy drives", " [+] drives [pas tous recopiés] pour user Dest " & $idrh2 & "  depuis user Srce " & $idrh1, 7)
												_GUICtrlRichEdit_SetCharColor($aff, 0x00ff0000)      ;   set color blue (BBGGRR format)
												_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & " Quelques [+] drives n'ont pas été récopiés vers user Dest " & $idrh2 & "  depuis user Srce " & $idrh1 & @CRLF & @CRLF);, 1)
												$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":   Quelques [+] drives n'ont pas été récopiés vers user Dest  " & $idrh2 & "  depuis user Srce  " & $idrh1 & @CRLF
												_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
												$drivesidrh2 = _arraytostring($drivesidrh2, @CRLF)
												$drivesidrh2bis = _ad_getobjectproperties($idrh2, "LaPoste00-NetworkDrive")
												_arraydelete($drivesidrh2bis, 0)
												$drivesidrh2bis = _arraytostring($drivesidrh2bis)
												$drivesidrh2bis = StringReplace($drivesidrh2bis, "LaPoste00-NetworkDrive|", "")
												$drivesidrh2bis = StringSplit($drivesidrh2bis, @CRLF)
												For $i = UBound($drivesidrh2bis) - 1 To 0 Step -1
													If $drivesidrh2bis[$i] = "" Then
														_arraydelete($drivesidrh2bis, $i)
													EndIf
												Next
												_arraydelete($drivesidrh2bis, 0)
												$drivesidrh2bis = _arraytostring($drivesidrh2bis, @CRLF)
												$result = "Warning !  :  " & "cpy drives depuis : " & $idrh1 & " vers " & $idrh2 & @CRLF & @CRLF & "Drives [+] :  :)" & @CRLF & $cpydrives1bis & @CRLF & @CRLF & "Lettres Drives déjà utilisés par user Dest  '" & $idrh2 & "'  ne peut pas etre écrasé..." & @CRLF & "-Vous devez rajouter les Drives non copiés manuellement !  :(" & $doublons & @CRLF & @CRLF & "'Liste Drives'  avant copie pour user Dest.  '" & $idrh2 & "' :" & @CRLF & $drivesidrh2 & @CRLF & @CRLF & "'Liste Drives' après copie pour user Dest.   '" & $idrh2 & "' :" & @CRLF & $drivesidrh2bis & @CRLF
												ClipPut($result)
												_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
												_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $result & @CRLF & @CRLF & "			=====	=====	=====	=====	=====			" & @CRLF & @CRLF);, 1)
												$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $result
											EndIf
										ElseIf @error = 1 Then
											MsgBox(0, "info cpy drives", "user Dest n'existe pas ! " & $idrh2)
											_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "user Dest n'existe pas ! " & $idrh2 & @CRLF & @CRLF);, 1)
											_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
										Else
											MsgBox(0, "info cpy drives", "[Return code " & @error & "] from Active Directory for user " & $idrh2 & @CRLF & "Acces refusé ...")
											_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "[Return code " & @error & "] from Active Directory for user " & $idrh2 & "  Acces refusé ..." & @CRLF & @CRLF);, 1)
											_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
										EndIf
									EndIf
								 Else
									if StringUpper($idrh2) <> "" and $checkbuttondrives=1 Then
									MsgBox(0,"Warning !","'cpy Drives' n'est valable que pour le domaine DCT !",15)
									EndIf
								EndIf
								If $idcheckscancpy = 1 And $isDCT=1 and $userexist2=1 Then
									$commentidrh2 = $commentidrh1
									If $commentidrh2 = "" Then
										MsgBox(0, "Warning !", "user 'Srce': " & $idrh1 & ", comment 'Mes scan' non defini !" & @CRLF & "Abandon copy Comment 'Me scan' depuis user Srce vers user Dest: " & $idrh2 & " !", 15)
										_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "user 'Srce': " & $idrh1 & ", comment 'Mes scan' non defini !" & @CRLF & "abandon copy Comment 'Me scan' depuis user Srce vers user Dest: " & $idrh2 & " !" & @CRLF & @CRLF);, 1)
										_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
									Else
										$commentidrh2 = StringReplace($commentidrh2, $idrh1, $idrh2)
										_ad_modifyattribute($idrh2, "Comment", "", 1)
										Global $ivalue = _ad_modifyattribute($idrh2, "Comment", $commentidrh2, 2)
										If $ivalue = 1 Then
											MsgBox(0, "info cpy Comment 'Mes scan'", " [+] comment [OK] pour user Dest " & $idrh2 & "  depuis user Srce " & $idrh1, 15)
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & " [+] comment [OK] pour user Dest " & $idrh2 & "  depuis user Srce " & $idrh1 & @CRLF & @CRLF);, 1)
										ElseIf @error = 1 Then
											MsgBox(0, "info cpy Comment 'Mes scan'", " user Dest n'existe pas : " & $idrh2, 15)
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "[user n'existe pas !] " & $idrh2 & @CRLF & @CRLF);, 1)
										Else
											MsgBox(0, "info cpy Comment 'Mes scan'", "[Return code " & @error & "]  from Active Directory for  " & $idrh2 & @CRLF & "Acces refusé !...", 15)
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "[Return code " & @error & "]  from Active Directory for  " & $idrh2 & @CRLF & "Acces refusé !..." & @CRLF & @CRLF);, 1)
										EndIf
									EndIf
								 Else
									if StringUpper($idrh2) <> "" and $idcheckscancpy=1 Then
									MsgBox(0,"Warning !","'cpy Mes Scan' n'est valable que pour le domaine DCT !",15)
									endif
								EndIf
								If $checkbuttongroups = 1 and $userexist2=1 Then ;cpy Groups
									_arraydisplay($groupidrh1_add, "  Liste des groupes pour [+] vers user Dest " & $idrh2)
									$idrh2 = _ad_samaccountnametofqdn($idrh2)
									$mail = _ad_getobjectattribute($idrh2, "mail")
									For $k = 0 To UBound($groupidrh1_add) - 1
										$ivalue = _ad_addusertogroup($groupidrh1_add[$k], $idrh2)
										If $ivalue = 1 Then
											ToolTip("User '" & $idrh2 & "' [+] groupe OK :  " & $groupidrh1_add[$k] & "", 5, 5, "")
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "User '" & $idrh2 & "' [+] groupe OK :  " & $groupidrh1_add[$k] & "");, 1)
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "User '" & $idrh2 & "' assigné au groupe metier  '" & $groupidrh1_add[$k] & "'"
										ElseIf @error = 1 Then
											MsgBox(64, "Active Directory", "Groupe '" & $groupidrh1_add[$k] & "' n'existe pas !")
											_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Active Directory", "Le Groupe '" & $groupidrh1_add[$k] & "' n'existe pas !");, 1)
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":  Le Groupe '" & $groupidrh1_add[$k] & "' n'existe pas !"
											_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
										ElseIf @error = 2 Then
											MsgBox(64, "Active Directory", "User Dest '" & $idrh2 & "' n'existe pas !")
											_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "User Dest '" & $idrh2 & "' n'existe pas !");, 1)
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":   User Dest '" & $idrh2 & "' n'existe pas !"
											_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
										ElseIf @error = 3 Then
											ToolTip("User '" & $idrh2 & "' déjà membre de '" & $groupidrh1_add[$k] & "'", 5, 5, "")
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "User '" & $idrh2 & "' déjà membre de '" & $groupidrh1_add[$k] & "'");, 1)
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":  User '" & $idrh2 & "' déjà membre de '" & $groupidrh1_add[$k] & "'"
											_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
										Else
											MsgBox(64, "Active Directory", "Return code '" & @error & "' from Active Directory for adding group " & $groupidrh1_add[$k] & " to " & $idrh2 & @CRLF & "Acces refusé ! ...")
											_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
											_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Return code '" & @error & "' from Active Directory for adding group " & $groupidrh1_add[$k] & " pour user Dest " & $idrh2 & " :" & " Acces refusé ! ...");, 1)
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":  " & $groupidrh1_add[$k] & "  =>  " & $idrh2 & ", acces refusé !"
											_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
										EndIf
										If StringInStr($groupidrh1_add[$k], "_AW Commun Manager") AND $mail = "" Then
											_ad_removeuserfromgroup($groupidrh1_add[$k], $idrh2)
											$historik = $historik & @CRLF & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":  " & $groupidrh1_add[$k] & "  =>  " & $idrh2 & ", retiré email Description !" & @CRLF
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $groupidrh1_add[$k] & "  =>  " & $idrh2 & ", retiré email Description !"  & @CRLF);, 1)
										EndIf
										If StringInStr($groupidrh1_add[$k], "_AW BPE") AND $mail = "" Then
											_ad_removeuserfromgroup($groupidrh1_add[$k], $idrh2)
											$historik = $historik & @CRLF & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":  " & $groupidrh1_add[$k] & "  =>  " & $idrh2 & ", retiré email Description !" & @CRLF
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $groupidrh1_add[$k] & "  =>  " & $idrh2 & ", retiré email Description !" & @CRLF);, 1)
										EndIf
										If StringInStr($groupidrh1_add[$k], "_AW BPE Restreint") AND $mail = "" Then
											_ad_removeuserfromgroup($groupidrh1_add[$k], $idrh2)
											$historik = $historik & @CRLF & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":  " & $groupidrh1_add[$k] & "  =>  " & $idrh2 & ", retiré email Description !" & @CRLF
										_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $groupidrh1_add[$k] & "  =>  " & $idrh2 & ", retiré email Description !" & @CRLF);, 1)
										EndIf
									Next
								 Else
								EndIf
							EndIf
							ToolTip("", 5, 5, "")

							Guidelete($hGUI)
							Return 0
						Else
							MsgBox(0, "Warning !", "user Srce  '" & $idrh1 & "'  n'existe pas !", 5)
							_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
							_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "user Srce:  '" & $idrh1 & "'  n'existe pas !");, 1)
							$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":  user Srce:  " & $idrh1 & ", n'existe pas !"
						    _GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
							Guidelete($hGUI)
							Return 0
						EndIf
					EndIf

					 Guidelete($hGUI)
					 Return 0

			    Case $gui_event_close
					;_terminate()

					Guidelete($hGUI)
					Return 0

			EndSwitch

		WEnd

EndFunc

#Region Localgroups

	Func readlocalgroup2()
	   Local $aPos = WinGetPos($hguiP)
_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Local Groups;" & @CRLF)
		 _Addusers() ;decompresse l'exe de Microsoft Addusers.exe au meme endroit que AD Tools puis le cache
		;FileInstall("addusers.exe", ".\") ; peut provoquer des faux positifs sur des exe embbeded .. plutot utiliser une routine de compactage binaire integré au code source comme une fonction..
		FileSetAttrib(@ScriptDir & "\addusers.exe", "+h")
		Global $cpname
		Global $readcombo
		Global $listmembers
		Global $userexist
		Global $hguilg = GUICreate("AD Tools [Add / Remove Members]-[Local Group]", 669, 433, $aPos[0]+38, $aPos[1]+28 )
		GUISetBkColor(16777215)
		Global $group1 = GUICtrlCreateGroup("", 8, 8, 324, 185)
		GUICtrlSetFont(-1, 10, 800, 0, "Arial")
		$labelcpname = GUICtrlCreateLabel("Computer:", 16, 32, 81, 23)
		GUICtrlSetFont(-1, 12, 800, 0, "Arial")
		Global $inputcpname = GUICtrlCreateInput(@ComputerName, 16, 56, 179, 30)
		GUICtrlSetFont(-1, 14, 800, 0, "Arial")
		$labellocalgroups = GUICtrlCreateLabel("Local Groups", 16, 101, 107, 20)
		GUICtrlSetFont(-1, 12, 800, 0, "Arial")
		Global $combogroups = GUICtrlCreateCombo("", 16, 157, 298, 25)
		GUICtrlSetData(-1, "")
		$buttonreadgroups = GUICtrlCreateButton("read local Groups", 200, 56, 120, 25)
		$buttonreadmembers = GUICtrlCreateButton("read Members", 16, 125, 130, 25)
		$iconreadlocalgroups = GUICtrlCreateIcon("", -1, 300, 60, 16, 16)
		$iconreadmembers = GUICtrlCreateIcon("", -1, 155, 130, 16, 16)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		$group2 = GUICtrlCreateGroup("User", 8, 210, 324, 199)
		GUICtrlSetFont(-1, 10, 800, 0, "Arial")
		$labelusersrce = GUICtrlCreateLabel("User Srce", 16, 259, 86, 23)
		$labeluserdest = GUICtrlCreateLabel("User Dest", 16, 339, 86, 23)
		Global $inputcomputer = GUICtrlCreateInput("", 88, 251, 89, 30)
		GUICtrlSetFont(-1, 14, 800, 0, "Arial")
		Global $inputcomputer2 = GUICtrlCreateInput("", 88, 331, 89, 30)
		GUICtrlSetFont(-1, 14, 800, 0, "Arial")
		$buttonadd = GUICtrlCreateButton("Add User", 224, 236, 100, 25)
		GUICtrlSetTip($buttonadd, "Add 'user Srce' to Local Group")
		$buttonremove = GUICtrlCreateButton("Remove user", 224, 273, 100, 25)
		GUICtrlSetTip($buttonremove, "Remove 'user Srce' from Local Group")
		$iconadduser = GUICtrlCreateIcon("", -1, 295, 240, 16, 16)
		$iconremoveuser = GUICtrlCreateIcon("", -1, 295, 280, 16, 16)
		$buttoncopy = GUICtrlCreateButton("Copy/Remove user", 224, 323, 100, 25)
		GUICtrlSetTip($buttoncopy, "Copy/Remove groups from 'user Srce' to 'user Dest'")
		$buttonremoveall = GUICtrlCreateButton("Remove All LG", 224, 360, 100, 25)
		GUICtrlSetTip($buttonremoveall, "Remove all Local Groups from 'user Dest.'")
		$buttonlistall = GUICtrlCreateButton("List user Srce", 83, 220, 100, 25)
		GUICtrlSetTip($buttonlistall, "List of all Local Groups from user 'Srce'")
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		Global $listmemberslg = GUICtrlCreateList("", 339, 31, 325, 383)
		GUICtrlSetData(-1, "")
		$labelmembersof = GUICtrlCreateLabel("Members of Group", 445, 10, 121, 20)
		GUICtrlSetFont(-1, 10, 800, 0, "Arial")
		GUICtrlSetColor(-1, 255)
		GUISetState(@SW_SHOW, $hguilg)
		Local $actiongl = ""
		While 1
			$msg = GUIGetMsg()
			Select
				Case $msg = $gui_event_close
					GUIDelete($hguilg)
					FileDelete(@TempDir & "\add_" & @ComputerName & ".txt")
					FileDelete(@TempDir & "\*.txt")
					FileDelete(@ScriptDir & "\addusers.exe")
					Return 0
				Case $msg = $buttonreadgroups
					$cpname = GUICtrlRead($inputcpname)
					If $cpname = "" Then
						MsgBox(16, "Error", "Missing Computer Nr. ")
					Else
						GUICtrlSetData($listmemberslg, "")
						GUICtrlSetData($listmemberslg, "working...")
						$agroups = _localaccounts_getgrouplist($cpname)
						GUICtrlSetData($listmemberslg, "")
					EndIf
				Case $msg = $buttonadd
					If GUICtrlRead($inputcomputer) <> "" AND GUICtrlRead($combogroups) <> "" Then
						GUICtrlSetData($listmemberslg, "")
						GUICtrlSetData($listmemberslg, "working...")
						RunWait(@ComSpec & " /c " & "net localgroup " & GUICtrlRead($combogroups) & " " & GUICtrlRead($inputcomputer) & " /add", @WorkingDir, @SW_HIDE)
						GUICtrlSetData($listmemberslg, "")
						checkuserexist()
						$u = ""
						If $userexist = 1 Then
							$u = " (user Srce found in AD)"
						Else
							$u = " (user Srce not found in AD)"
						EndIf
						GUICtrlSetData($listmemberslg, "" & StringUpper(GUICtrlRead($inputcomputer)) & " [+] " & GUICtrlRead($combogroups) & $u)
						listedesgroupeslocaux()
						If String($listelgroupatraiter) <> "" Then
							$listatraiter = StringSplit($listelgroupatraiter, ";")
							If IsArray($listatraiter) Then
								For $z = 1 To $listatraiter[0]
									GUICtrlSetData($listmemberslg, $listatraiter[$z] & @CRLF);, 1)
								Next
							EndIf
						EndIf
					Else
						GUICtrlSetData($listmemberslg, "")
						GUICtrlSetData($listmemberslg, "missing Local Group or User Srce")
					EndIf
				Case $msg = $buttonremove
					If GUICtrlRead($inputcomputer) <> "" AND GUICtrlRead($combogroups) <> "" Then
						GUICtrlSetData($listmemberslg, "")
						GUICtrlSetData($listmemberslg, "working...")
						RunWait(@ComSpec & " /c " & "net localgroup " & GUICtrlRead($combogroups) & " " & GUICtrlRead($inputcomputer) & " /DELETE", @WorkingDir, @SW_HIDE)
						GUICtrlSetData($listmemberslg, "")
						checkuserexist()
						$u = ""
						If $userexist = 1 Then
							$u = " (user Srce found in AD)"
						Else
							$u = " (user Srce not found in AD)"
						EndIf
						GUICtrlSetData($listmemberslg, "" & StringUpper(GUICtrlRead($inputcomputer)) & " [-] " & GUICtrlRead($combogroups) & $u)
						listedesgroupeslocaux()
						If String($listelgroupatraiter) <> "" Then
							$listatraiter = StringSplit($listelgroupatraiter, ";")
							If IsArray($listatraiter) Then
								For $z = 1 To $listatraiter[0]
									GUICtrlSetData($listmemberslg, $listatraiter[$z] & @CRLF);, 1)
								Next
							EndIf
						EndIf
					Else
						GUICtrlSetData($listmemberslg, "")
						GUICtrlSetData($listmemberslg, "missing Local Group or User Srce")
					EndIf
				Case $msg = $buttonreadmembers
					Local $readcombo = GUICtrlRead($combogroups)
					If $readcombo = "" Then
						MsgBox(16, "Error", "no Group Selected ! " & @CRLF & "press button [read local Groups]")
					Else
						Global $cpname = GUICtrlRead($inputcpname)
						$ipid = Run(@ComSpec & " /c " & "net localgroup " & GUICtrlRead($combogroups) & "", @WorkingDir, @SW_HIDE, 2)
						ProcessWaitClose($ipid)
						$ausers = StdoutRead($ipid)
						$ausers = StringReplace($ausers, "Nom alias       Administrateurs", "")
						$ausers = StringReplace($ausers, "Commentaire     ", "")
						$ausers = StringReplace($ausers, "Membres", "")
						$ausers = StringReplace($ausers, "-------------------------------------------------------------------------------", "")
						$ausers = StringReplace($ausers, "La commande s'est termin‚e correctement.", "")
						$ausers = StringReplace($ausers, "Nom alias       Duplicateurs", "")
						$ausers = StringReplace($ausers, "Prend en charge la r‚plication des fichiers dans le domaine", "")
						$ausers = StringReplace($ausers, "Nom alias       IIS_IUSRS", "")
						$ausers = StringReplace($ausers, "Groupe int‚gr‚ utilis‚ par les services Internet (IIS).", "")
						$ausers = StringReplace($ausers, "Nom alias       Invit‚s", "")
						$ausers = StringReplace($ausers, "Les  du groupe Invit‚s disposent par d‚faut du mˆme accŠs que les  du groupe Utilisateurs, … l'exception du compte Invit‚ qui dispose d'autorisations restreintes", "")
						$ausers = StringReplace($ausers, "Nom alias       Utilisateurs", "")
						$ausers = StringReplace($ausers, "Les utilisateurs ne peuvent pas effectuer de modifications accidentelles ou intentionnelles … l'‚chelle du systŠmeÿ; par ailleurs, ils peuvent ex‚cuter la plupart des applications.", "")
						$ausers = StringReplace($ausers, "Nom alias       __vmware__", "")
						$ausers = StringReplace($ausers, "VMware User Group", "")
						ClipPut($ausers)
						$ausers = StringSplit($ausers, @LF)
						Dim $ausers2[0]
						For $i = 1 To $ausers[0]
							If StringLen($ausers[$i]) > 2 Then
								_arrayadd($ausers2, $ausers[$i])
								ReDim $ausers2
							EndIf
						Next
						GUICtrlSetData($listmemberslg, "")
						GUICtrlSetData($listmemberslg, _arraytostring($ausers2))
					EndIf
				Case $msg = $buttoncopy
					Select
						Case GUICtrlRead($inputcomputer) = "" AND GUICtrlRead($inputcomputer2) = ""
							GUICtrlSetData($listmemberslg, "")
							GUICtrlSetData($listmemberslg, "missing 2 Users")
						Case GUICtrlRead($inputcomputer) = GUICtrlRead($inputcomputer2)
							GUICtrlSetData($listmemberslg, "")
							GUICtrlSetData($listmemberslg, "Same Users...")
						Case GUICtrlRead($inputcomputer) = "" OR GUICtrlRead($inputcomputer2) = ""
							GUICtrlSetData($listmemberslg, "")
							GUICtrlSetData($listmemberslg, "missing 'User Srce' or 'User Dest'")
						Case GUICtrlRead($inputcomputer) <> GUICtrlRead($inputcomputer2)
							$t = MsgBox(4, "Copy/Remove", "(Yes): to Copy user 'S' to 'D'" & @CRLF & "(No): to Remove user 'D' from 'S'")
							If $t = 6 Then
								$actiongl = 1
							ElseIf $t = 7 Then
								$actiongl = 0
							EndIf
							GUICtrlSetData($listmemberslg, "")
							GUICtrlSetData($listmemberslg, "working...")
							listlocalgroup()
							listedesgroupeslocaux()
							GUICtrlSetData($listmemberslg, "")
							If String($listelgroupatraiter) <> "" Then
								GUICtrlSetData($listmemberslg, $listelgroupatraiter, 1)
								$copygrps = StringSplit($listelgroupatraiter, ";")
								For $z = 1 To $copygrps[0]
									If $actiongl = 1 Then
										RunWait(@ComSpec & " /c " & "net localgroup " & Chr(34) & $copygrps[$z] & Chr(34) & " " & GUICtrlRead($inputcomputer2) & " /ADD", @WorkingDir, @SW_HIDE)
										GUICtrlSetData($listmemberslg, "'" & $copygrps[$z] & "   [    adding    ]  " & GUICtrlRead($inputcomputer2) & "  from:  " & GUICtrlRead($inputcomputer) & @CRLF);, 1)
									Else
										RunWait(@ComSpec & " /c " & "net localgroup " & Chr(34) & $copygrps[$z] & Chr(34) & " " & GUICtrlRead($inputcomputer2) & " /DELETE", @WorkingDir, @SW_HIDE)
										GUICtrlSetData($listmemberslg, "'" & $copygrps[$z] & "   [   deleting   ]  " & GUICtrlRead($inputcomputer2) & "  from:  " & GUICtrlRead($inputcomputer) & @CRLF);, 1)
									EndIf
								Next
								checkuserexist()
								$u = ""
								If $userexist = 1 Then
									$u = " (user Srce found in AD)"
								Else
									$u = " (user Srce not found in AD)"
								EndIf
								checkuserexistd()
								$v = ""
								If $userexist = 1 Then
									$v = " (user Dest. found in AD)"
								Else
									$v = " (user Dest. not found in AD)"
								EndIf
								GUICtrlSetData($listmemberslg, $u & @CRLF & $v, 1)
								listedesgroupeslocaux()
								If String($listelgroupatraiter) <> "" Then
									$listatraiter = StringSplit($listelgroupatraiter, ";")
									If IsArray($listatraiter) Then
										For $z = 1 To $listatraiter[0]
											GUICtrlSetData($listmemberslg, $listatraiter[$z] & @CRLF);, 1)
										Next
									EndIf
								EndIf
							Else
								If $actiongl = 1 Then
									GUICtrlSetData($listmemberslg, GUICtrlRead($inputcomputer) & ", Source is not member of any Local Group..." & ", nothing to copy 'S' to 'D'  ..." & GUICtrlRead($inputcomputer2), 1)
								Else
									GUICtrlSetData($listmemberslg, GUICtrlRead($inputcomputer) & ", Source is not member of any Local Group..." & ", nothing to remove 'S' to 'D'  ..." & GUICtrlRead($inputcomputer2), 1)
								EndIf
							EndIf
					EndSelect
				Case $msg = $buttonremoveall
					If GUICtrlRead($inputcomputer2) <> "" AND GUICtrlRead($combogroups) <> "" Then
						GUICtrlSetData($listmemberslg, "")
						GUICtrlSetData($listmemberslg, "working...")
						listlocalgroup()
						listedesgroupeslocaux()
						GUICtrlSetData($listmemberslg, "")
						If String($listelgroupatraiter) <> "" Then
							$copygrps = StringSplit($listelgroupatraiter, ";")
							For $z = 1 To $copygrps[0]
								RunWait(@ComSpec & " /c " & "net localgroup " & Chr(34) & $copygrps[$z] & Chr(34) & " " & GUICtrlRead($inputcomputer2) & " /DELETE", @WorkingDir, @SW_HIDE)
								checkuserexistd()
								$u = ""
								If $userexist = 1 Then
									$u = " (user Dest. found in AD)"
								Else
									$u = " (user Dest. not found in AD)"
								EndIf
								GUICtrlSetData($listmemberslg, "'" & $copygrps[$z] & "   [   deleting   ]  " & GUICtrlRead($inputcomputer2) & @CRLF);, 1)
								GUICtrlSetData($listmemberslg, $u, 1)
							Next
						Else
							GUICtrlSetData($listmemberslg, GUICtrlRead($inputcomputer2) & ", Dest. is not member of any Local Group..." & ", nothing to remove...", 1)
						EndIf
					Else
						GUICtrlSetData($listmemberslg, "")
						GUICtrlSetData($listmemberslg, "button [read local group] not pressed ! or missing User 'Dest'")
					EndIf
				Case $msg = $buttonlistall
					If GUICtrlRead($inputcomputer) <> "" AND GUICtrlRead($combogroups) <> "" Then
						GUICtrlSetData($listmemberslg, "")
						GUICtrlSetData($listmemberslg, "working...")
						listlocalgroup()
						listedesgroupeslocaux()
						GUICtrlSetData($listmemberslg, "")
						If String($listelgroupatraiter) <> "" Then
							$listatraiter = StringSplit($listelgroupatraiter, ";")
							If IsArray($listatraiter) Then
								For $z = 1 To $listatraiter[0]
									GUICtrlSetData($listmemberslg, $listatraiter[$z] & @CRLF);, 1)
								Next
							EndIf
						Else
							GUICtrlSetData($listmemberslg, "")
							checkuserexist()
							$u = ""
							If $userexist = 1 Then
								$u = " (user found in AD)"
							Else
								$u = " (user not found in AD)"
							EndIf
							GUICtrlSetData($listmemberslg, "No Local Group found for User 'Srce'" & $u)
						EndIf
					Else
						GUICtrlSetData($listmemberslg, "")
						GUICtrlSetData($listmemberslg, "button [read local group] not pressed ! or missing User 'Srce'")
					EndIf
			EndSelect
		WEnd
	EndFunc

	Func checkuserexist()
		If _ad_objectexists(GUICtrlRead($inputcomputer)) Then
			$userexist = 1
		Else
			$userexist = 0
		EndIf
		Return $userexist
	EndFunc

	Func checkuserexistd()
		If _ad_objectexists(GUICtrlRead($inputcomputer2)) Then
			$userexist = 1
		Else
			$userexist = 0
		EndIf
		Return $userexist
	EndFunc

	Func listedesgroupeslocaux()
		listlocalgroup()
		$computerlistgroups = FileOpen(@TempDir & "\GL_" & @ComputerName & ".txt", 0)
		$linelg = FileRead($computerlistgroups)
		Global $aretarray = StringSplit($linelg, @CRLF)
		FileClose($computerlistgroups)
		Global $listelgroupatraiter = ""
		If IsArray($aretarray) Then
			For $z = 1 To $aretarray[0]
				$linetmp = $aretarray[$z]
				If StringInStr($linetmp, GUICtrlRead($inputcomputer)) Then
					$linetmp2 = StringSplit($linetmp, ",")
					If IsArray($linetmp2) Then
						$groupdetect = $linetmp2[1]
						$listelgroupatraiter = $listelgroupatraiter & $groupdetect & ";"
					EndIf
				EndIf
			Next
			Return $listelgroupatraiter
		EndIf
	EndFunc

	Func _localaccounts_getgrouplist($scomputername = $cpname)
		Global $afilter = ["group"], $aresult[1], $ogroup
		Global $ocomputer = ObjGet("WinNT://" & $cpname)
		If NOT IsObj($ocomputer) Then Return SetError(1, 0, 0)
		$ocomputer.filter = $afilter
		For $ogroup In $ocomputer
			ReDim $aresult[UBound($aresult) + 1]
			$aresult[UBound($aresult) - 1] = $ogroup.name
		Next
		$aresult[0] = UBound($aresult) - 1
		If IsArray($aresult) AND $aresult[0] <> "" Then
			_arraydelete($aresult, 0)
			$firstlg = $aresult[0]
			GUICtrlSetData($combogroups, _arraytostring($aresult), $firstlg)
			GUICtrlSetBkColor($combogroups, 16776960)
			Return $aresult
		Else
			MsgBox(0, "Warning !", "No access !?" & @CRLF & "Local Group not reachable from computer [" & @ComputerName & "]  to  [" & $cpname & "]")
		EndIf
	EndFunc

	Func _localaccounts_getuserlist($scomputername = $cpname)
		Global $afilter = ["user"], $aresult[1], $ouser
		Global $ocomputer = ObjGet("WinNT://" & $cpname & "\" & $readcombo)
		$ocomputer.filter = $afilter
		For $ouser In $ocomputer
			ReDim $aresult[UBound($aresult) + 1]
			$aresult[UBound($aresult) - 1] = $ouser.name
		Next
		$aresult[0] = UBound($aresult) - 1
		_arraydelete($aresult, 0)
		GUICtrlSetData($listmemberslg, _arraytostring($aresult))
		Return $aresult
	EndFunc

	Func readlocalgroup()
	    _Addusers()
		; FileInstall("addusers.exe", ".\") ;exe embedded provoque faux positif ?
		; FileSetAttrib("addusers.exe", "+h")
		Global $hguilg = GUICreate("Managing Locals Groups [" & @ComputerName & "]", 596, 550)
		$labeltitre = GUICtrlCreateLabel("Managing users and locals groups  (by: EAPI69/NR)", 24, 0, 549, 20, 1)
		GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
		$labelgroup = GUICtrlCreateLabel("List all locals groups with users memberof from..." & @CRLF & " [ " & @ComputerName & " ]", 8, 32, 282, 41, 8388608, 512)
		$buttonlistservergroup = GUICtrlCreateButton("List server groups", 386, 32, 100, 41)
		$labellistusergrp = GUICtrlCreateLabel("List groups wich belong to a user and then list them", 8, 88, 282, 41, 8388608, 512)
		$listusergrp = GUICtrlCreateInput("", 312, 90, 65, 21, -1, 512)
		GUICtrlSetLimit($listusergrp, 7)
		$buttonlistonly = GUICtrlCreateButton("List only", 392, 88, 89, 41)
		$buttonaddlg = GUICtrlCreateButton("Add Local Group", 496, 88, 89, 41)
		$usersbelgroup2 = GUICtrlCreateLabel("List groups wich belong to user and then Delete user from is own locals groups", 8, 144, 282, 41, 8388608, 512)
		$listuserdelete = GUICtrlCreateInput("", 312, 146, 65, 21, -1, 512)
		GUICtrlSetLimit($listuserdelete, 7)
		$buttonviewlist = GUICtrlCreateButton("View list", 392, 144, 89, 41)
		$buttondeleteuser = GUICtrlCreateButton("Delete user", 496, 144, 89, 41)
		$usersbelcopysd = GUICtrlCreateLabel("Copy from user 'S' all locals groups to user 'D'", 8, 200, 282, 57, 8388608, 512)
		$usersbeld = GUICtrlCreateLabel("D", 296, 204, 12, 17)
		GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
		$labels = GUICtrlCreateLabel("S", 296, 236, 12, 17)
		GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
		Global $userd = GUICtrlCreateInput("", 312, 202, 65, 21, -1, 512)
		GUICtrlSetLimit($userd, 7)
		Global $users = GUICtrlCreateInput("", 312, 234, 65, 21, -1, 512)
		GUICtrlSetLimit($users, 7)
		$buttonlistusers = GUICtrlCreateButton("View list 'S'", 392, 200, 89, 57)
		$buttoncopys2d = GUICtrlCreateButton("Copy 'S' to 'D'", 496, 200, 89, 57)
		$exitlg = GUICtrlCreateButton("[ Exit ]", 305, 272, 280, 49)
		GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
		$buttonfolder = GUICtrlCreateButton("[ Folder ACL ]", 8, 272, 280, 49)
		GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
		Global $affiche = GUICtrlCreateEdit("", 8, 338, 577, 208, -1, 512)
		GUISetState(@SW_SHOW, $hguilg)
		$listgroup = 0
		While 1
			$msg = GUIGetMsg()
			Select
				Case $msg = $buttonfolder
					$sfolder = FileSelectFolder("Select a folder to scan", "")
					If @error Then
					Else
						$sfolder &= ""
						$cmd = "icacls.exe " & Chr(34) & $sfolder
						$foo = Run(@ComSpec & " /c " & $cmd, "", @SW_HIDE, $stdout_child)
						ProcessWaitClose($foo)
						While 1
							$line = StdoutRead($foo)
							If @error Then
								ExitLoop
							EndIf
							$line = StringStripWS($line, 4)
							$line = StringReplace($line, "‚", "é")
							$line = StringReplace($line, "š", "è")
							$line = StringReplace($line, "Š", "è")
							$line = StringReplace($line, $sfolder, "")
							$line = StringTrimLeft($line, 1)
							If StringInStr($line, "1") Then
								$line = StringSplit($line, "1")
								$line = $line[1]
							EndIf
							If NOT StringInStr($sfolder, "{") OR NOT StringInStr($sfolder, ":") Then
								If $listgroup = 1 Then
									GUICtrlSetData($affiche, "")
									$listgroup = 0
								EndIf
								GUICtrlSetData($affiche, $sfolder & " => " & $line & @CRLF);, 1)
							Else
								MsgBox(0, "Not with system folder !", $sfolder)
							EndIf
						WEnd
					EndIf
				Case $msg = $exitlg
					GUIDelete($hguilg)
					Return 0
				Case $msg = $buttonlistservergroup
					$listgroup = 1
					GUICtrlSetData($affiche, "")
					listlocalgroup()
					afficheruser(@TempDir & "\GL_" & @ComputerName & ".txt")
				Case $msg = $buttonlistonly
					If StringLen(GUICtrlRead($listusergrp)) = 7 Then
						GUICtrlSetData($affiche, "")
						listlocalgroup()
						compare2(GUICtrlRead($listusergrp))
						afficheruser(@TempDir & "\G_" & GUICtrlRead($listusergrp) & ".txt")
					Else
						MsgBox(0, "", "Idrh empty ?")
					EndIf
				Case $msg = $buttonviewlist
					If StringLen(GUICtrlRead($listuserdelete)) = 7 Then
						GUICtrlSetData($affiche, "")
						listlocalgroup()
						traitetl(GUICtrlRead($listuserdelete))
						afficheruser(@TempDir & "\T_" & GUICtrlRead($listuserdelete) & ".txt")
						$ealg = @TempDir & "\T_" & GUICtrlRead($listuserdelete) & ".txt"
						GUICtrlSetState($buttondeleteuser, 64)
					Else
						MsgBox(0, "", "Idrh empty ?")
					EndIf
				Case $msg = $buttondeleteuser
					If StringLen(GUICtrlRead($listuserdelete)) = 7 Then
						deleteuserlc(@TempDir & "\T_" & GUICtrlRead($listuserdelete) & ".txt")
						GUICtrlSetData($affiche, "")
						GUICtrlSetState($buttondeleteuser, 128)
					Else
						MsgBox(0, "", "Idrh empty ?")
					EndIf
				Case $msg = $buttonlistusers
					If StringLen(GUICtrlRead($userd)) = 7 AND StringLen(GUICtrlRead($users)) = 7 Then
						GUICtrlSetData($affiche, "")
						listlocalgroup()
						compare1(GUICtrlRead($users), GUICtrlRead($userd))
						afficheruser(@TempDir & "\C_" & GUICtrlRead($userd) & ".txt")
						$ealg = @TempDir & "\C_" & GUICtrlRead($userd) & ".txt"
						GUICtrlSetState($buttoncopys2d, 64)
					Else
						MsgBox(0, "", "Idrh 'S' or 'D' empty ?")
					EndIf
				Case $msg = $buttoncopys2d
					If StringLen(GUICtrlRead($userd)) = 7 AND StringLen(GUICtrlRead($users)) = 7 Then
						create(@TempDir & "\C_" & GUICtrlRead($userd) & ".txt")
						GUICtrlSetData($affiche, "")
						GUICtrlSetState($buttoncopys2d, 128)
						_ad_open()
						$exist = ""
						If _ad_objectexists(GUICtrlRead($userd)) Then
							$exist = "'D'  [ " & GUICtrlRead($userd) & "  exist ? => YES ]"
						Else
							$exist = "'D'  [ " & GUICtrlRead($userd) & "  exist ? => NO  ]"
						EndIf
						GUICtrlSetData($affiche, "" & $exist & @CRLF);, 1)
						If _ad_objectexists(GUICtrlRead($users)) Then
							$exist = "'S'  [ " & GUICtrlRead($users) & "  exist ? => YES ]"
						Else
							$exist = "'S'  [ " & GUICtrlRead($users) & "  exist ? => NO  ]"
						EndIf
						GUICtrlSetData($affiche, "" & $exist & @CRLF);, 1)
					Else
						MsgBox(0, "", "Idrh 'D' or 'S' empty ?")
					EndIf
				Case $msg = $buttonaddlg
					$act = 1
					Local $hdll = DllOpen("user32.dll")
					If StringLen(GUICtrlRead($listusergrp)) <> 7 Then
						MsgBox(0, "", "Idrh empty ?")
					Else
						$user = GUICtrlRead($listusergrp)
						Local $ifileexists = FileExists(@TempDir & "\GL_" & @ComputerName & ".txt")
						If $ifileexists = 1 Then
							$logfile = @TempDir & "\GL_" & @ComputerName & ".txt"
							$string = FileRead($logfile)
							If $string = "" Then
								MsgBox(0, "Info !", @TempDir & "\GL_" & @ComputerName & ".txt , is an empty file !")
							Else
								Dim $aarray
								_filereadtoarray(@TempDir & "\GL_" & @ComputerName & ".txt", $aarray)
								$hguigl = GUICreate("List of [" & @ComputerName & "] Local groups", 500, 150)
								$buttok = GUICtrlCreateButton(" [OK]=Add group  // [Shift]+[OK]=Remove group", 50, 60, 400, 40)
								$line = ""
								$line1 = ""
								$combogl = GUICtrlCreateCombo("", 10, 10, 480, 25)
								For $j = 2 To UBound($aarray) - 1
									$line1 = $aarray[$j]
									$line1 = StringSplit($line1, ",")
									$line1 = $line1[1]
									$line = $line & $line1 & "|"
									If $j = 2 Then
										$line0 = $line1
									EndIf
								Next
								GUICtrlSetData($combogl, $line, $line0)
								GUISetState(@SW_SHOW)
								While 1
									Switch GUIGetMsg()
										Case $buttok
											If _ispressed("10", $hdll) Then
												$act = 0
											Else
												$act = 1
											EndIf
											_ad_open()
											Global $scomboreadgl = GUICtrlRead($combogl)
											If $scomboreadgl = "" Then
												MsgBox(0, "", "please Select a local group")
											Else
												$exist = ""
												If _ad_objectexists($user) Then
													$exist = "  [ " & $user & "  exist ? => YES ]"
												Else
													$exist = "  [ " & $user & "  exist ? => NO  ]"
												EndIf
												If $act = 1 Then
													RunWait(@ComSpec & " /c " & "net localgroup " & Chr(34) & $scomboreadgl & Chr(34) & " " & $user & " /ADD", @WorkingDir, @SW_HIDE)
													GUICtrlSetData($affiche, "'" & $scomboreadgl & "'  [  adding   ]  " & $user & "" & $exist & @CRLF);, 1)
												Else
													RunWait(@ComSpec & " /c " & "net localgroup " & Chr(34) & $scomboreadgl & Chr(34) & " " & $user & " /DELETE", @WorkingDir, @SW_HIDE)
													GUICtrlSetData($affiche, "'" & $scomboreadgl & "'  [  removing ]  " & $user & "" & $exist & @CRLF);, 1)
												EndIf
												GUIDelete($hguigl)
												ExitLoop
											EndIf
										Case $gui_event_close
											GUIDelete($hguigl)
											FileDelete(@scriptdir & "\Addusers.exe")
											ExitLoop
									EndSwitch
								WEnd
							EndIf
						Else
							MsgBox(0, "Info !", "Clic button [List Servers groups] at first to use this function !")
						EndIf
					EndIf
			EndSelect
		WEnd
	EndFunc

	Func create($cu)
		RunWait(@ComSpec & " /c " & "addusers /c " & $cu, @WorkingDir, @SW_HIDE)
	EndFunc

	Func listlocalgroup()
		RunWait(@ComSpec & " /c " & "addusers \\" & @ComputerName & " /d " & @TempDir & "\add_" & @ComputerName & ".txt", @WorkingDir, @SW_HIDE)
		FileSetAttrib(@TempDir & "\add_" & @ComputerName & ".txt", "+h")
		$computerlgadd = FileOpen(@TempDir & "\add_" & @ComputerName & ".txt", 0)
		$computerlistgroups = FileOpen(@TempDir & "\GL_" & @ComputerName & ".txt", 2)
		$islocalgroup = 0
		While 1
			$readlinelg = FileReadLine($computerlgadd)
			If @error = -1 Then ExitLoop
			$readlinelg = StringStripCR($readlinelg)
			If $islocalgroup = 0 Then
				If $readlinelg = "[Local]" Then $islocalgroup = 1
			EndIf
			If ($islocalgroup = 1) AND ($readlinelg <> "") Then
				$readlinelg = StringReplace($readlinelg, "COURRIER\", "")
				FileWriteLine($computerlistgroups, $readlinelg)
			EndIf
		WEnd
		FileClose($computerlgadd)
		FileClose($computerlistgroups)
		FileDelete(@TempDir & "\add_" & @ComputerName & ".txt")
	EndFunc

	Func afficheruser($cu)
		$computerlgadd = FileOpen($cu, 0)
		GUICtrlSetData($affiche, "")
		While 1
			$readlinelg = FileReadLine($computerlgadd)
			If @error = -1 Then ExitLoop
			$readlinelg = StringStripCR($readlinelg)
			GUICtrlSetData($affiche, $readlinelg & @CRLF);, 1)
		WEnd
		FileClose($computerlgadd)
	EndFunc

	Func traitetl($tu)
		$tu = StringUpper($tu)
		$computerlgadd = FileOpen(@TempDir & "\GL_" & @ComputerName & ".txt", 0)
		$computerlistgroups = FileOpen(@TempDir & "\T_" & $tu & ".txt", 2)
		While 1
			$readlinelg = FileReadLine($computerlgadd)
			If @error = -1 Then ExitLoop
			$readlinelg = StringStripCR($readlinelg)
			If StringInStr(StringUpper($readlinelg), $tu) Then
				$readlinelg2 = StringSplit($readlinelg, ",")
				FileWriteLine($computerlistgroups, $readlinelg2[1] & "," & $tu)
			EndIf
		WEnd
		FileClose($computerlgadd)
		FileClose($computerlistgroups)
	EndFunc

	Func compare1($cuser1, $cuser2)
		$cuser1 = StringUpper($cuser1)
		$cuser2 = StringUpper($cuser2)
		$computerlgadd = FileOpen(@TempDir & "\GL_" & @ComputerName & ".txt", 0)
		$computerlistgroups = FileOpen(@TempDir & "\C_" & $cuser2 & ".txt", 2)
		FileWriteLine($computerlistgroups, "[Local]")
		While 1
			$readlinelg = FileReadLine($computerlgadd)
			If @error = -1 Then ExitLoop
			$readlinelg = StringStripCR($readlinelg)
			If StringInStr(StringUpper($readlinelg), $cuser1) Then
				$readlinelg2 = StringSplit($readlinelg, ",")
				FileWriteLine($computerlistgroups, $readlinelg2[1] & "," & $readlinelg2[2] & "," & $cuser2)
			EndIf
		WEnd
		FileClose($computerlgadd)
		FileClose($computerlistgroups)
	EndFunc

	Func compare2($tu)
		$tu = StringUpper($tu)
		$computerlgadd = FileOpen(@TempDir & "\GL_" & @ComputerName & ".txt", 0)
		$computerlistgroups = FileOpen(@TempDir & "\G_" & $tu & ".txt", 2)
		FileWriteLine($computerlistgroups, $tu & " is member of:")
		While 1
			$readlinelg = FileReadLine($computerlgadd)
			If @error = -1 Then ExitLoop
			$readlinelg = StringStripCR($readlinelg)
			If StringInStr(StringUpper($readlinelg), $tu) Then
				$readlinelg2 = StringSplit($readlinelg, ",")
				FileWriteLine($computerlistgroups, $readlinelg2[1])
			EndIf
		WEnd
		FileClose($computerlgadd)
		FileClose($computerlistgroups)
	EndFunc

	Func ecrireaffiche($var1r)
		$fichierea = FileOpen($ealg, 2)
		FileWrite($fichierea, _8a($affiche))
		FileClose($fichierea)
	EndFunc

	Func deleteuserlc($cu)
		$computerlgadd = FileOpen($cu, 0)
		While 1
			$readlinelg = FileReadLine($computerlgadd)
			If @error = -1 Then ExitLoop
			$1t = StringSplit($readlinelg, ",")
			RunWait(@ComSpec & " /c " & "net localgroup " & $1t[1] & " " & $1t[2] & " /DELETE", @WorkingDir, @SW_HIDE)
		WEnd
		FileClose($computerlgadd)
	EndFunc

#EndRegion Localgroups

Func bitlocker()
	_ad_open()
	  if $domainname="" Then
		   MsgBox(0,"Info !","no Active Directory found ! unable to scan...",15)
		   ToolTip("",5,5,"")
		   Return 0
		EndIf

	treeview_affiche()
	treeselect()
	Global $aous = _ad_getobjectsinou($sout, "(objectClass=computer)")
	$filez = FileOpen(@ScriptDir & "\computers-AD_OU" & ".txt", 2)
	If @error > 0 Then
		MsgBox(64, $adroot, "No computers found ! ")
		_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
		_GUICtrlRichEdit_AppendText($aff, @CRLF & "Bitlocker Error ! No computers found in this OU: " & $sout);, 1)
		ToolTip("", 5, 5)
		Return 0
	Else
		$filterou = ""
		$text = "La POSTE/DSEM/EAPI69/NR"
		Global $aselected[1]
		Global $hgui2 = GUICreate("Select computer for Bitlocker key  and [ Apply ] !  (" & $text & ")", 1000, 500)
		Global $hlistbox = _guictrllistbox_create($hgui2, "String upon creation", 10, 30, 980, 470, BitOR($ws_border, $lbs_hasstrings, $ws_vscroll, $lbs_multiplesel))
		_guictrllistbox_beginupdate($hlistbox)
		_guictrllistbox_resetcontent($hlistbox)
		_guictrllistbox_initstorage($hlistbox, 100, 4096)
		_guictrllistbox_endupdate($hlistbox)
		If IsArray($aous) = 1 Then
			For $z = 1 To $aous[0] - 1
				FileWriteLine($filez, $aous[$z] & @CRLF)
				If StringLen($filterou) = 0 Then
					_guictrllistbox_addstring($hlistbox, $aous[$z])
				Else
					If NOT StringInStr($filterou, ";") Then
						If StringInStr($aous[$z], $filterou) Then
							_guictrllistbox_addstring($hlistbox, $aous[$z])
						EndIf
					Else
						Local $filterou2 = StringSplit($filterou, ";")
						For $w = 1 To $filterou2[0]
							If StringInStr($aous[$z], $filterou2[$w]) Then
								_guictrllistbox_addstring($hlistbox, $aous[$z])
							EndIf
						Next
					EndIf
				EndIf
			Next
		EndIf
		FileClose($filez)
		FileDelete(@ScriptDir & "\computers-AD_OU" & ".txt")
		If IsArray($aous) = 0 Then
			GUIDelete($hgui2)
			ToolTip("", 5, 5)
			MsgBox(0, "Info !", "aborting, no computers found !", 7)
			_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "Bitlocker Error ! No computers found in this OU: " & $sout);, 1)
			Return 0
		EndIf
		Global $hbutton = GUICtrlCreateButton("[ Apply ]", 20, 3, 80, 25)
		GUISetState()
		While 1
			Switch GUIGetMsg()
				Case $hbutton
					$aselected = _guictrllistbox_getselitems($hlistbox)
					If $aselected[0] = 1 Then
						$sitem = " item"
					Else
						$sitem = " items"
					EndIf
					$sitems = ""
					For $i = 1 To $aselected[0]
						$sitems &= _guictrllistbox_gettext($hlistbox, $aselected[$i]) & @CRLF
					Next
					If $aselected[0] <> 0 Then
						$sitems = StringSplit($sitems, @CRLF)
						GUIDelete($hgui2)
						ToolTip("", 5, 5)
						For $z = 1 To $sitems[0]
							If StringLen($sitems[$z]) <> 0 Then
								$sad_ou = $sitems[$z]
								$sad_ou2 = _ad_samaccountnametofqdn($sad_ou)
								ToolTip("Bitlocker for:  [" & $sad_ou2 & "]" & @CRLF & "   " & $z & "/" & $sitems[0], 5, 5)
								$aresult = _ad_getobjectsinou($sad_ou2, "(objectcategory=msFVE-RecoveryInformation)", 2, "distinguishedname")
								If @error <> 0 Then
									_GUICtrlRichEdit_AppendText($aff, @CRLF & "Bitlocker Recovery:  " & "@error: " & @error & ", @extended: " & @extended & "  No rights or Bitlocker not found for: " & $sad_ou & " !  Aborting...");, 1)
								Else
									_GUICtrlRichEdit_AppendText($aff, @CRLF & $sad_ou & @CRLF & _arraytostring($aresult);, 1)
									$aresult = _ad_getobjectproperties($aresult[1])
									If @error <> 0 Then
										_GUICtrlRichEdit_AppendText($aff, "Bitlocker Recovery:  " & "@error: " & @error & ", @extended: " & @extended & "  No rights or Bitlocker not found !   " & "  Aborting...");, 1)
									Else
										_GUICtrlRichEdit_AppendText($aff, @CRLF & $sad_ou & @CRLF & _arraytostring($aresult);, 1)
									EndIf
								EndIf
							EndIf
						Next
						ToolTip("", 5, 5)
						ExitLoop
					Else
						MsgBox(4160, "warning !", "aucune selection... réessayez ou fermer la fenetre... [x]")
					EndIf
				Case $gui_event_close
					GUIDelete($hgui2)
					ExitLoop
			EndSwitch
		WEnd
		GUIDelete($hgui2)
	EndIf
EndFunc

Func computergroup()
	$defautdc = $defautdcinit
	Dim $sitems, $sitemss
	Global $sidkeymatch = ""
	Global $multipleidrhbis = ""
	Global $displayname = ""
	Global $actualcomputername = ""

	  if $domainname="" Then
		   MsgBox(0,"Info !","no Active Directory found ! unable to scan...",15)
		   ToolTip("",5,5,"")
		   Return 0
		EndIf

$result=""

	$t = MsgBox(4, "Choix (O/N)", "[Oui]: ajoute un groupe => ordinateur" & @CRLF & @CRLF & "[Non]: retire un groupe => ordinateur")
	If $t = 6 Then
		$test = 0
	ElseIf $t = 7 Then
		$test = 1
	EndIf
	_ad_open()

	treeview_affiche()
	treeselect()

	Global $aous = _ad_getobjectsinou($sout, "(objectClass=computer)")
	$filez = FileOpen(@ScriptDir & "\computers-AD_OU" & ".txt", 2)
	If @error > 0 Then
		MsgBox(64, $adroot, "No computers found ! ")
		 if $isDCT=1 Then
			_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "DCT: SCCM group in PITR ! No computers found in this OU: " & $sout);, 1)
		 Else
			_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
			_GUICtrlRichEdit_AppendText($aff, @CRLF & " No computers found in this OU: " & $sout);, 1)
		 EndIf
		ToolTip("", 5, 5)
		Return 0
	 EndIf

	$filterou = ""
	if $test=0 Then
	$text = "  computer(s):  action is,  [+]   groups"
    Else
	$text = "  computer(s):  action is,  [-]   groups"
    EndIf

	Global $aselected[1]
		 If $isdct = 1 Then
	  Global $hgui2 = GUICreate("Select computer(s) for SCCM/PITR groups, and [ Apply ] !  (" & $text & ")", 1000, 500)
	    Else
	  Global $hgui2 = GUICreate("Select computer(s) for groups, and [ Apply ] !  (" & $text & ")", 1000, 500)
	    EndIf

	Global $hlistbox = _guictrllistbox_create($hgui2, "String upon creation", 10, 30, 980, 470, BitOR($ws_border, $lbs_hasstrings, $ws_vscroll, $lbs_multiplesel))
	_guictrllistbox_beginupdate($hlistbox)
	_guictrllistbox_resetcontent($hlistbox)
	_guictrllistbox_initstorage($hlistbox, 100, 4096)
	_guictrllistbox_endupdate($hlistbox)
	If IsArray($aous) = 1 Then
		For $z = 1 To $aous[0] - 1
			FileWriteLine($filez, $aous[$z] & @CRLF)
			If StringLen($filterou) = 0 Then
				_guictrllistbox_addstring($hlistbox, $aous[$z])
			Else
				If NOT StringInStr($filterou, ";") Then
					If StringInStr($aous[$z], $filterou) Then
						_guictrllistbox_addstring($hlistbox, $aous[$z])
					EndIf
				Else
					Local $filterou2 = StringSplit($filterou, ";")
					For $w = 1 To $filterou2[0]
						If StringInStr($aous[$z], $filterou2[$w]) Then
							_guictrllistbox_addstring($hlistbox, $aous[$z])
						EndIf
					Next
				EndIf
			EndIf
		Next
	EndIf
	FileClose($filez)
	FileDelete(@ScriptDir & "\computers-AD_OU" & ".txt")
	If IsArray($aous) = 0 Then
		GUIDelete($hgui2)
		ToolTip("", 5, 5)
		MsgBox(0, "Info !", "aborting, no computers found !", 7)
				  if $isDCT=1 Then
				  _GUICtrlRichEdit_AppendText($aff, @CRLF & "DCT: SCCM group in PITR ! No computers found in this OU: " & $sout & @crlf );, 1)
				  Else
				  _GUICtrlRichEdit_AppendText($aff, @CRLF & " No computers found in this OU: " & $sout& @CRLF );, 1)
				  EndIf
		Return 0
	EndIf
	Global $hbutton = GUICtrlCreateButton("[ Apply ]", 20, 3, 80, 25)
	GUISetState()
	While 1
		Switch GUIGetMsg()
			Case $hbutton
				$aselected = _guictrllistbox_getselitems($hlistbox)
				If $aselected[0] = 1 Then
					$sitem = " item"
				Else
					$sitem = " items"
				EndIf
				$sitems = ""
				For $i = 1 To $aselected[0]
					$sitems &= _guictrllistbox_gettext($hlistbox, $aselected[$i]) & @CRLF
				Next
				If $aselected[0] <> 0 Then
					$sitems = StringSplit($sitems, @CRLF)
					GUIDelete($hgui2)
					ToolTip("", 5, 5)
					ExitLoop
				Else
					MsgBox(4160, "warning !", "aucune selection... réessayez ou fermer la fenetre... [x]")
				EndIf
			Case $gui_event_close
				GUIDelete($hgui2)
				ExitLoop
		EndSwitch
	WEnd
	If IsArray($sitems) = 0 Then
		MsgBox(4160, "warning !", "no computer(s) selected !... abort !")
		_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  Pas d' Ordinateur(s) selectionnés...  Abandon !" & @CRLF);, 1)
		Return 0
	 EndIf

  ;$multipleidrh=$sItems

		for $i=1 to $sItems[0]
$computers=$sItems[$i]

$groupsComputer = _AD_GetUserGroups($computers)
If @error > 0 Then
Else
    _ArraySort($groupsComputer, 0, 1)
   _ArrayDelete($groupsComputer,0)
Global $groupIDRH_final=""
Global $groupidrh1_add=$groupsComputer

for $k=0 to UBound($groupsComputer)-1
$groupIDRH=$groupsComputer[$k]
$groupIDRH=stringsplit($groupIDRH,",")
$groupIDRH_int=$groupIDRH[1]
$groupIDRH_int=StringReplace($groupIDRH_int,"CN=","")
$groupIDRH_final=$groupIDRH_final & $groupIDRH_int & " ; " ;@crlf
next
   $groupsComputer = $groupIDRH_final
   EndIf

;if $groupsComputer<>"" then
if $computers<>"" Then
$result=$result & "  "  & $computers  & "  =>  groupes:  " & $groupsComputer &  @CRLF
endif
;
next ;$i

_GUICtrlRichEdit_AppendText($aff,  @crlf & "  Time: " & @hour & ":" & @min & ":" & @sec & "=>   Initial computer groups: " & @crlf & $result  & "  Time: " & @hour & ":" & @min & ":" & @sec & "   Initial computer groups   <=" & @crlf);,1)
$historik=$historik & @crlf & "  Time: " & @hour & ":" & @min & ":" & @sec & "=>   Initial computer groups: " & @crlf & $result  & "  Time: " & @hour & ":" & @min & ":" & @sec & "   Initial computer groups	<=" & @crlf


	If $isdct = 1 Then
		$unite = InputBox("default OU ? 2 or 4 chr$", "ex: MILY, MITE, BPLY, GAUB, MI ... virtuos" & @CRLF & "?? : scan all OUs...", "PITR")
		If StringLen($unite) < 2 OR StringLen($unite) > 4 AND NOT ($unite = "virtuos" OR $unite = "??") Then
			ToolTip("", 5, 5)
			Return 0
		EndIf
	Else
		$unite = "??"
		ToolTip("", 5, 5)
	EndIf
	If $unite = "??" Then
		$allouu = 1
		treeview_affiche()
		treeselect()
	Else
		$allouu = 0
	EndIf
	$defautdc = StringReplace($defautdc, ".", ",DC=")
	$defautdc = ",DC=" & $defautdc
	ToolTip("", 5, 5)

	If $allouu = 1 Then
		$unite = ""
		$allouu = 0
		Global $souu = $sout
	Else
		Global $souu = "OU=Groupes,OU=" & $unite & ",OU=" & StringMid($unite, 1, 2) & $defautdc
	 EndIf

	If $unite = "MI" Then
		Global $souu = "OU=Groupes,OU=_Support Transverse,OU=MI" & $defautdc
	 EndIf

	If $unite = "virtuos" Then
		Global $souu = "OU=FunctionalGroups,OU=_Flexible Workspace Commun" & $defautdc
	 EndIf

	$filezz = FileOpen(@ScriptDir & "\Groups-AD_OU-" & $unite & ".txt", 2)
	Global $aouus = _ad_getobjectsinou($souu, "(objectClass=group)")
	If @error > 0 Then
		MsgBox(64, $adroot, "No Groups found ! " & @CRLF & $souu)
		ToolTip("", 5, 5)
		Return 0
	 EndIf

 if $isDCT=1 Then
$filterOUu=InputBox("filtering groups ? [;]","empty = all groups , or type: W9D; etc... ","W9D;")
Else
$filterOUu=InputBox("filtering groups ? [;]","empty = all groups , or type keyword1;keyword2 ... ","")
   EndIf

Global $aSelected[1]
if $test=1 then
Global $hGUI2 = GUICreate("Select groups to remove  and [ Apply ] !  (" & $filterOUu & ")", 1000, 500)
Else
Global $hGUI2 = GUICreate("Select groups for adding and [ Apply ] !  (" & $filterOUu & ")", 1000, 500)
EndIf

	Global $hlistboxx = _guictrllistbox_create($hgui2, "String upon creation", 10, 30, 980, 470, BitOR($ws_border, $lbs_hasstrings, $ws_vscroll, $lbs_multiplesel))
	_guictrllistbox_beginupdate($hlistboxx)
	_guictrllistbox_resetcontent($hlistboxx)
	_guictrllistbox_initstorage($hlistboxx, 100, 4096)
	_guictrllistbox_endupdate($hlistboxx)
	If IsArray($aouus) = 1 Then
		For $z = 1 To $aouus[0] - 1
			FileWriteLine($filezz, $aouus[$z] & @CRLF)
			If StringLen($filterouu) = 0 Then
				_guictrllistbox_addstring($hlistboxx, $aouus[$z])
			Else
				If NOT StringInStr($filterouu, ";") Then
					If StringInStr($aouus[$z], $filterouu) Then
						_guictrllistbox_addstring($hlistboxx, $aouus[$z])
					EndIf
				Else
					Local $filterou2 = StringSplit($filterouu, ";")
					For $w = 1 To $filterou2[0]
						If StringInStr($aouus[$z], $filterou2[$w]) Then
							_guictrllistbox_addstring($hlistboxx, $aouus[$z])
						EndIf
					Next
				EndIf
			EndIf
		Next
	 EndIf

	FileClose($filezz)
	FileDelete(@ScriptDir & "\Groups-AD_OU-" & $unite & ".txt")

	If _ispressed("10", $hdll) Then
	Else
		FileDelete(@ScriptDir & "\Groups-AD_OU-" & $unite & ".txt")
	 EndIf

	If IsArray($aouus) = 0 Then
		$idcheckboxgroup = 2
		GUIDelete($hgui2)
		ToolTip("", 5, 5)
		MsgBox(0, "Info !", "aborting, no groups found in: " & $unite, 7)
		Return 0
	 EndIf

	Global $hbutton = GUICtrlCreateButton("[ Apply ]", 20, 3, 80, 25)
	GUISetState()
	While 1
		Switch GUIGetMsg()
		Case $hbutton
				$aselected = _guictrllistbox_getselitems($hlistboxx)
				If $aselected[0] = 1 Then
					$sitem = " item"
				Else
					$sitem = " items"
				 EndIf
				$sitemss = ""
				For $i = 1 To $aselected[0]
					$sitemss &= _guictrllistbox_gettext($hlistboxx, $aselected[$i]) & @CRLF
				Next
				If $aselected[0] <> 0 Then
					$sitemss = StringSplit($sitemss, @CRLF)
					If $test = 0 Then
						$result = "		[+] group(s) to selected computer(s) ...	" & @CRLF
					Else
						$result = "		[-] group(s) to selected computer(s) ...	" & @CRLF
					EndIf
					If IsArray($sitemss) = 0 Then
						MsgBox(4160, "warning !", "no group(s) selected !... abort !")
						Return 0
					EndIf
					 ;sort de la boucle et termine boucle for $i (computers[$i])
					ExitLoop
				Else
					MsgBox(4160, "warning !", "aucune selection... réessayez ou fermer la fenetre... [x]")
				EndIf
			Case $gui_event_close
				$idcheckboxgroup = 2
				GUIDelete($hgui2)
				ExitLoop
		EndSwitch
	 WEnd

	 ; boucle for $i (computers[$i]=$sitems[$i] )
	 GUIDelete($hgui2)
	 if IsArray($sitemss) then ;=> on a selectionné des groupes on continue sinon on passe à la fin => abandon !
					For $i = 1 To $sitems[0]
						$computers = $sitems[$i]
						   if $computers<>"" then
							  $result=$result & "> " & $computers & @CRLF
						   EndIf
						$userexist = _ad_objectexists($computers)
						If $userexist = 1 Then
							For $z = 1 To $sitemss[0] ;( $sitemss[$z] = liste des groupes selectionnés )
								If StringLen($sitemss[$z]) <> 0 Then
									$multipleidrh = $sitems
									If $test = 0 Then
										$ivalue = _ad_addusertogroup($sitemss[$z], $computers)
										If $ivalue = 1 Then
											$txt = " 	: [OK]"
										Else
											$txt = " 	: [KO], est deja membre ou pb. de droits à l'ajout du groupe."
										EndIf
										$result = $result &  "[+] groupe: " & $sitemss[$z] & "  =>  computer:  " & $computers & $txt & @crlf
									Else
										$ivalue = _ad_removeuserfromgroup($sitemss[$z], $computers)
										If $ivalue = 1 Then
											$txt = " 	: [OK]"
										Else
											$txt = " 	: [KO], n'est deja pas membre ou pb. de droits pour retirer le groupe."
										EndIf
										$result = $result &  "[-] groupe: " & $sitemss[$z] & "  =>  computer:  " & $computers & $txt & @crlf
									EndIf
								Else
									If $computers <> "" AND $userexist = 0 Then
										$result = $result & $computers & " pas trouvé sur l'AD ! "
									EndIf
								EndIf
							Next
						EndIf
						$groupscomputer = _ad_getusergroups($computers)
						If @error > 0 Then
						Else
							_arraysort($groupscomputer, 0, 1)
							_arraydelete($groupscomputer, 0)
							Global $groupidrh_final = ""
							Global $groupidrh1_add = $groupscomputer
							For $k = 0 To UBound($groupscomputer) - 1
								$groupidrh = $groupscomputer[$k]
								$groupidrh = StringSplit($groupidrh, ",")
								$groupidrh_int = $groupidrh[1]
								$groupidrh_int = StringReplace($groupidrh_int, "CN=", "")
								$groupidrh_final = $groupidrh_final & $groupidrh_int & " ; "
							Next
							$groupscomputer = $groupidrh_final
						EndIf
						If $groupscomputer <> "" Then
							$result = $result &  "  " & $computers & @CRLF & "   => Liste finale:  " & $groupscomputer & @CRLF
						EndIf
					 Next

					ToolTip("", 5, 5)
					GUIDelete($hgui2)
					ClipPut($result)
					_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $result & @CRLF & "			=====	=====	=====	=====	=====			" & @CRLF);, 1)
					$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $result
				 Else ; Abandon car aucun groupe n'a été selectionné...
					MsgBox(0,"Information !","  Pas de Groupe(s) selectionné(s) pour Ordinateur(s) !" & @crlf & @crlf & " Abandon !")
			_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
	 		_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  Pas de groupe(s) selectionné(s) => Ordinateur(s)...  Abandon !" & @CRLF);, 1)
			;$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $result
	 EndIf

	 ; fin boucle for $i (computers[$i])

	ToolTip("", 5, 5)
	;GUIDelete($hgui2)
	Return 0
EndFunc

#Region Directives Metiers
;/////////////////////////////////////////////////////  Directives Metiers domaine DCT uniquement  /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Func directives() ;routine principale des Directives Metiers
	$defautdc = $defautdcinit
    $numdirmetier=""

	#Region IHM Directives
		If $isdct = 0 Then
			MsgBox(0, "warning !", "'Directives Metier'  valable que pour le domaine DCT !", 7)
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "'Directives Metiers' valable que pour le domaine: DCT.ADT.Local " & @CRLF);, 1)
			Return 0
		 EndIf

		Global $ousourcedir = _ad_getobjectattribute($idrh1, "distinguishedName")
		$ousourcedir = StringSplit($ousourcedir, ",")
		$ousourcedir = $ousourcedir[3]
		$ousourcedir = StringTrimLeft($ousourcedir, 3)
		$hguidir = GUICreate("Selection 'Directive Metier'  puis fermer fenetre [x]...", 650, 80)
	#EndRegion

	#Region Liste des Directives
		If $directives = "" Then
			liste_directives()
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & " Premier lancement des 'Directives Metiers' !  (matrice crée)" & @CRLF);, 1)
			_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
		EndIf
	#EndRegion Liste des Directives

	$idcombobox = GUICtrlCreateCombo("", 20, 20, 600, 250)
	GUICtrlSetData(-1, $directives, "_SM_Sites Tertiaires RLP")
	GUISetState(@SW_SHOW, $hguidir)
	While 1
		Switch GUIGetMsg()
			Case $gui_event_close
				Global $scomboreaddirectives = GUICtrlRead($idcombobox)
				If StringLen($scomboreaddirectives) > 4 Then
				Else
					Global $dirapplied = 0
					MsgBox(0, "Warning !", "Directive Metier n'existe pas dans cette OU !", 7)
					_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & " Directive Metier n'existe pas... abandon !" & @CRLF);, 1)
					GUIDelete($hguidir)
					ExitLoop
				EndIf
				GUIDelete($hguidir)
				If $scomboreaddirectives = "[ Zero Directives ]" Then
					$t = MsgBox(4, "Zero Directives ?", "Souhaitez-vous retirer les Directives pour : " & $idrh1 & " ...?")
				Else
					$t = MsgBox(4, "Appliquer Directive ?", "Souhaitez-vous appliquer la Directive ? " & @CRLF & $scomboreaddirectives)
					 $numdirmetier=StringLeft($scomboreaddirectives,3) ; exemple: 733
					; MsgBox(0,"",$numdirmetier)
					; Exit
				EndIf
				If $t = 6 Then
					SplashTextOn("", "Valide la Directive:  " & $scomboreaddirectives & " !", 1100, 100, -1, -1, 1, -1, 13, 600)
					Sleep(1200)
					SplashOff()
					$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  Valider une 'Directive'"
					_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  Valider une 'Directive': " & $scomboreaddirectives & @CRLF);, 1)
					Global $dirapplied = 0

			#Region traitement des Directives
						If NOT StringInStr($scomboreaddirectives, ";") Then
							If StringInStr($scomboreaddirectives, "[EAID]") Then
								$direaid = _stringexplode($scomboreaddirectives, "]", -2)
								$dirid = $direaid[1]
								$scomboreaddirectives = $dirid
							 EndIf

							#Region anciennes Directives
								If $scomboreaddirectives = "[ Zero Directives ]" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user Srce " & $idrh1 & " n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$scomboreaddirectives = " [ Zero Directives ] = Retirer la Directive !"
									$dirapplied = 1
								EndIf
								If $scomboreaddirectives = "_SM_Sites Tertiaires RLP" OR $scomboreaddirectives = "_SM_Site_Tertiaire_Enseigne" Then ; _SM_Sites Tertiaires RLP
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes..", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "_SM_Sites Tertiaires RLP", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Sites Tertiaires RLP", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User Srce '" & $idrh1 & "'  assigné  au groupe metier '" & "RG-" & $ou & "_SM_Sites Tertiaires RLP" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe '" & "RG-" & $ou & "_SM_Sites Tertiaires RLP" & "' n'existe pas dans cette OU")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User Srce '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' déjà membre du groupe metier '" & "RG-" & $ou & "_SM_Sites Tertiaires RLP" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_GUICHET (VirtuOS)" Then ;Chargé de CLI
								   ;SG-GAUB_ACCES_EVS; USR_BP_GUICHET_GENE;USR_BP_GUICHET_ESPACE_CO
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE ... " & @CRLF & "Dir-" & $numdirmetier & "   =>  " & "[ GAUB ]" & @CRLF & "_SM_GUICHET" & @crlf & " on ne met plus: RG-GAUB_SM_GUICHET...", "GAUB")
									move_user()
									;Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_GUICHET", $idrh1)
									Global $ivalue = _ad_addusertogroup("SG-" & $ou & "_ACCES_EVS" , $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "SG-" & $ou & "_ACCES_EVS" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "SG-" & $ou & "_ACCES_EVS" & "' n'existe pas !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "SG-" & $ou & "_ACCES_EVS" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
										_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)
										_ad_addusertogroup("USR_BP_GUICHET_ESPACE_CO", $idrh1)
									;	_ad_addusertogroup("RG-PITR_VIR_COM1_IE11_NOREDIRECTFOLDER", $idrh1)
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
							    If $scomboreaddirectives = "_Vendeur Formateur LPM" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE ... " & @CRLF & "_Vendeur Formateur LPM" & @crlf & " Vendeur Formateur LPM... idem Guichetier en plus:" & @crlf & "RG-PITR_VIR_COM1_IE11_NOREDIRECTFOLDER", "GAUB")
									move_user()
									;Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_GUICHET", $idrh1)
									Global $ivalue = _ad_addusertogroup("SG-" & $ou & "_ACCES_EVS" , $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "SG-" & $ou & "_ACCES_EVS" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "SG-" & $ou & "_ACCES_EVS" & "' n'existe pas !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "SG-" & $ou & "_ACCES_EVS" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
										_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)
										_ad_addusertogroup("USR_BP_GUICHET_ESPACE_CO", $idrh1)
										_ad_addusertogroup("RG-PITR_VIR_COM1_IE11_NOREDIRECTFOLDER", $idrh1)
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								 EndIf

;COBA inhibés avec #cs/#ce...
#cs
							    If $scomboreaddirectives = "19 _SM_Conseil Bancaire" Then
								   ;RG-[EAID]_SM_Conseil Bancaire; SG-GAUB_ACCES_EVS; USR_BP_GUICHET_GENE
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									_arraysort($groupsidrh1, 0, 1)
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE ... " & @CRLF & "_SM_Conseil Bancaire", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Conseil Bancaire", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "RG-" & $ou & "_SM_Conseil Bancaire" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe '" & "RG-" & $ou & "_SM_Conseil Bancaire" & "' n'existe pas !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "RG-" & $ou & "_SM_Conseil Bancaire" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
										_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Conseil Bancaire W10" Then ;SM_Conseil Bancaire (w10)
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									_arraysort($groupsidrh1, 0, 1)
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE ... " & @CRLF & "_SM_Conseil Bancaire W10 (sur BPnn)", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "SM_Conseil Bancaire", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "RG-" & $ou & "SM_Conseil Bancaire" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe '" & "RG-" & $ou & "SM_Conseil Bancaire" & "' n'existe pas !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "RG-" & $ou & "SM_Conseil Bancaire" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										; _ad_addusertogroup("RG-PITR_CM_COBA_Agent", $idrh1)
										_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
										_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
#ce

							    If $scomboreaddirectives = "_SM_Conseil Bancaire" Then ; $t = 6 == win 10 (_SM_DET)   //    $t = 7 == win 7  (_SM_Conseil Bancaire)
								   ;RG-[EAID]_SM_Conseil Bancaire; SG-GAUB_ACCES_EVS; USR_BP_GUICHET_GENE
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									_arraysort($groupsidrh1, 0, 1)
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE ... " & @CRLF & "_SM_Conseil Bancaire (sur BPnn)", $ousourcedir)
									move_user()
									$t=7 ; on inhibe le choix w7/w10 ci-dessous !
									;$t = MsgBox(4, "choix W10/W7", "_SM_Conseil Bancaire" & @crlf & @CRLF & "  poste W10: (Oui) => (_SM_DET)?" & @CRLF & "  poste W7: (Non) => (_SM_Conseil Bancaire)")
									If $t = 6 Then 		;_DET
										Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_DET", $idrh1) ;_DET
									ElseIf $t = 7 Then  ;_Conseil bancaire + role COBA
										Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Conseil Bancaire", $idrh1) ;_Conseil bancaire + role COBA
									EndIf
									If $ivalue = 1 Then
										$dirapplied = 1
										If $t = 6 Then
											MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "RG-" & $ou & "_SM_DET" & "'")
										EndIf
										If $t = 7 Then
											MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "RG-" & $ou & "_SM_Conseil Bancaire" & "'")
										EndIf
									ElseIf @error = 1 Then
										; " n'existe pas !"
										If $t = 6 Then
											MsgBox(64, "", "Groupe '" & "RG-" & $ou & "_SM_DET" & "' n'existe pas !")
										EndIf
										If $t = 7 Then
											MsgBox(64, "", "Groupe '" & "RG-" & $ou & "_SM_Conseil Bancaire" & "' n'existe pas !")
										EndIf
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										; " User $idrh1 est déjà membre du groupe "
										If $t = 6 Then
											MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "RG-" & $ou & "_SM_DET" & "'")
										EndIf
										If $t = 7 Then
											MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "RG-" & $ou & "_SM_Conseil Bancaire" & "'")
										 EndIf

										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										 If $t = 6 Then ;win 10 (_SM_det)
											_ad_addusertogroup("RG-PITR_CM_COBA_Agent", $idrh1)
										 EndIf
										 If $t = 7 Then ;win 7 (_SM_Conseil Bancaire)
										_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
										_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)
										_ad_addusertogroup("RG-PITR_CM_COBA_Agent", $idrh1) ;rectif vu avec Estelle...
										EndIf
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Siege_Enseigne" Then ; n'existe plus renommé en RG-[EAID]_SM_Siege Banque (ne sera pas pris en compte)
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "_SM_Siege_Enseigne", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Siege_Enseigne", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "RG-" & $ou & "_SM_Siege_Enseigne" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe '" & "RG-" & $ou & "_SM_Siege_Enseigne" & "' n'existe pas !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "RG-" & $ou & "_SM_Siege_Enseigne" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Siege Banque" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "_SM_Siege Banque", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Siege Banque", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "RG-" & $ou & "_SM_Siege Banque" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe '" & "RG-" & $ou & "_SM_Siege Banque" & "' n'existe pas !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "RG-" & $ou & "_SM_Siege Banque" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Siege Banque_LBPGP" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "_SM_Siege Banque_LBPGP", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Siege Banque_LBPGP", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "RG-" & $ou & "_SM_Siege Banque_LBPGP" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe '" & "RG-" & $ou & "_SM_Siege Banque_LBPGP" & "' n'existe pas !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "RG-" & $ou & "_SM_Siege Banque_LBPGP" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Siege Banque_LBPAI" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "_SM_Siege Banque_LBPAI", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Siege Banque_LBPAI", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "RG-" & $ou & "_SM_Siege Banque_LBPAI" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe '" & "RG-" & $ou & "_SM_Siege Banque_LBPAI" & "' n'existe pas !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "RG-" & $ou & "_SM_Siege Banque_LBPAI" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Siege Banque_FondEcranFixe" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "_SM_Siege Banque_FondEcranFixe", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Siege Banque_FondEcranFixe", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "RG-" & $ou & "_SM_Siege Banque_FondEcranFixe" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe '" & "RG-" & $ou & "_SM_Siege Banque_FondEcranFixe" & "' n'existe pas !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "RG-" & $ou & "_SM_Siege Banque_FondEcranFixe" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
										_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_DISFE" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, MITO ... " & @CRLF & "_SM_DISFE", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_DISFE", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "RG-" & $ou & "_SM_DISFE" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe '" & "RG-" & $ou & "_SM_DISFE" & "' n'existe pas !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "RG-" & $ou & "_SM_DISFE" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Siege Banque_LBPF" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "_SM_Siege Banque_LBPF", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Siege Banque_LBPF", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "RG-" & $ou & "_SM_Siege Banque_LBPF" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe '" & "RG-" & $ou & "_SM_Siege Banque_LBPF" & "' n'existe pas !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "RG-" & $ou & "_SM_Siege Banque_LBPF" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_DET" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "_SM_DET", "GAUB")
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_DET", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "RG-" & $ou & "_SM_DET" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe '" & "RG-" & $ou & "_SM_DET" & "' n'existe pas !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "RG-" & $ou & "_SM_DET" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										_ad_addusertogroup("RG-PITR_CM_DIR-BP_DS_REC_RE_RC", $idrh1)
										$t = MsgBox(4, "Groupes GAUB DET ?", "Rajout des 2 groupes DET accès Lecteurs <regate> ?" & @CRLF & "( RG-GAUB_[regate]_GESTION_GROUPE ; SG-GAUB_[regate]_PRIVE )" & @CRLF & @CRLF & "pour le Role: DET...")
										If $t = 6 Then
											Global $regategaub = ""
											$regategaub = InputBox("Regate (6 chiffres) ?", "ex: 694100" & @CRLF & "RG-GAUB_[regate]_GESTION_GROUPE" & @CRLF & "SG-GAUB_[regate]_PRIVE" & @CRLF & "Oui => (REC/RE,DS et CECI)" & @CRLF & "Non => (CCPRO,GCPRO,GDCPRO,GESPRO,RCPART)", "")
											_ad_addusertogroup("RG-GAUB_" & $regategaub & "_GESTION_GROUPE", $idrh1)
											_ad_addusertogroup("SG-GAUB_" & $regategaub & "_PRIVE", $idrh1)
										Else
										EndIf
										$t = MsgBox(4, "DET/REC, ou CCPRO VirtuOS Guichet ?", "Rajout des 2 groupes virtuOS Guichet ?" & @CRLF & "( USR_BP_CAISSE_GENE ; USR_BP_GUICHET_GENE )" & @CRLF & @CRLF & "pour le Role: DET...")
										If $t = 6 Then
											_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
											_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)
										Else
										EndIf
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Centres Financiers et Centres Nationaux" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "Dir-" & $numdirmetier & "   =>  " & "[ CFxx ]" & @CRLF & "SM_Centres Financiers et Centres Nationaux", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Centres Financiers et Centres Nationaux", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "RG-" & $ou & "_SM_Centres Financiers et Centres Nationaux" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe '" & "RG-" & $ou & "_SM_Centres Financiers et Centres Nationaux" & "' n'existe pas !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "RG-" & $ou & "_SM_Centres Financiers et Centres Nationaux" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Utilisateurs_Postes_SIA" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "SM_Utilisateurs_Postes_SIA", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Utilisateurs_Postes_SIA", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "RG-" & $ou & "_SM_Utilisateurs_Postes_SIA" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe '" & "RG-" & $ou & "_SM_Utilisateurs_Postes_SIA" & "' n'existe pas !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "RG-" & $ou & "_SM_Utilisateurs_Postes_SIA" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_RF_INDET" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "SM_RF_INDET", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_RF_INDET", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "RG-" & $ou & "_SM_RF_INDET" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe '" & "RG-" & $ou & "_SM_RF_INDET" & "' n'existe pas !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "RG-" & $ou & "_SM_RF_INDET" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Management Commercial Unique" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "SM_Management Commercial Unique", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Management Commercial Unique", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier  '" & "RG-" & $ou & "_SM_Management Commercial Unique" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe '" & "RG-" & $ou & "_SM_Management Commercial Unique" & "' n'existe pas !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' est déjà membre du groupe '" & "RG-" & $ou & "_SM_Management Commercial Unique" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive non appliquée ...")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Initiale : ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
							#EndRegion anciennes Directives

						 Else

							#Region nouvelles Directives

								#Region mask
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user Srce " & $idrh1 & " n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " groupes Obsoletes Directive : " & $actiongroups
									If StringInStr($scomboreaddirectives, ";") Then
										$dira = StringSplit($scomboreaddirectives, ";")
										$direaid = _stringexplode($dira[1], "]", -2)
										$dirid = $direaid[1]
										$dirpitr = $dira[2]
									 EndIf

									 Global $ouDirmetier=""
									 if StringInStr($numdirmetier,"790") Then ;GAUB
										$ouDirmetier="DIR-" & $numdirmetier & "   =>  " & "[ GAUB ]"

									 ElseIf StringInStr($numdirmetier,"733") Then
										$ouDirmetier="DIR-" & $numdirmetier & "   =>  " & "[ GAUB ]"

									 ElseIf StringInStr($numdirmetier,"789") Then
										$ouDirmetier="DIR-" & $numdirmetier & "   =>  " & "[ BPXX ]"

									 ElseIf StringInStr($numdirmetier,"731") Then
										$ouDirmetier="DIR-" & $numdirmetier & "   =>  " & "[ BPXX ]"

									 ElseIf StringInStr($numdirmetier,"774") Then
									    $ouDirmetier="DIR-" & $numdirmetier & "   =>  " & "[ BPXX ]"

									 ElseIf StringInStr($numdirmetier,"788") Then
										$ouDirmetier="DIR-" & $numdirmetier & "   =>  " & "[ GAUB ]"

									 ElseIf StringInStr($numdirmetier,"732") Then
										$ouDirmetier="DIR-" & $numdirmetier & "   =>  " & "[ GAUB ]"

									 ElseIf StringInStr($numdirmetier,"319") Then
										$ouDirmetier="DIR-" & $numdirmetier & "   =>  " & "[ BPXX ]"

									 Elseif StringInStr($dirid ,"_SM_Centres Financiers et Centres Nationaux") Then
									    $ouDirmetier="DIR-" & $numdirmetier & "   =>  " & "[ CFXX ]"

									 Else
										$ouDirmetier=""
								     EndIf ;cas normal

									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & $ouDirmetier & @crlf & @CRLF & $dirid, $ousourcedir)

									move_user()
								#EndRegion mask

								Global $ivalue = _ad_addusertogroup("RG-" & $ou & $dirid, $idrh1)
								If $ivalue = 1 Then
									$dirapplied = 1
									MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe Metier '" & "RG-" & $ou & $dirid & "'")
								ElseIf @error = 1 Then
									MsgBox(64, "", "Groupe '" & "RG-" & $ou & $dirid & "' n'existe pas dans cette OU !")  ; ex: RG-[EAID]_SM_Siege_Enseigne n'existe pas/plus.. => rollback
								ElseIf @error = 2 Then
									MsgBox(64, "", "User Srce '" & $idrh1 & "' n'existe pas !")
								ElseIf @error = 3 Then
									MsgBox(64, "", "User Srce '" & $idrh1 & "' déjà membre du groupe Metier '" & "RG-" & $ou & $dirid & "'")
									$dirapplied = 1
								Else
									MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
								 EndIf

								If $dirapplied = 0 Then
								    $defautdc = $defautdcinit
									$defautdc = StringReplace($defautdc, ".", ",DC=")
									$defautdc = ",DC=" & $defautdc
									$sobject  = _ad_samaccountnametofqdn($idrh1)
									$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
									Global $ivalue = _ad_moveobject($souinitial, $sobject)

									If $ivalue = 1 Then
										MsgBox(64, "Active Directory ", $idrh1 & "' replacé dans son OU d'origine '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Non Appliquée... (rollback) !")
										ToolTip("synchronising AD...", 5, 5)
										Sleep(1800)
										ToolTip("", 5, 5)
									 Else
										MsgBox(64, "Active Directory ", $idrh1 & "' Vous devez replacer manuellement dans son OU d'origine '" & $ousourcedir & "'" & "  user: " & $idrh1 & @CRLF & @CRLF & "Directive Non Appliquée... (rollback) !")
									 EndIf


									$restoredirectives = StringTrimRight($restoredirectives, 1)
									$groupsidrh1 = StringSplit($restoredirectives, "|")
									If IsArray($groupsidrh1) = 1 Then
										For $z = 1 To $groupsidrh1[0]
											_ad_addusertogroup($groupsidrh1[$z], $idrh1)
										Next
									EndIf
									$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback avec la Directive Metier Initiale :  ")
								EndIf

							    If $dirapplied = 1 Then
									_ad_addusertogroup($dirpitr, $idrh1)

									#Region cas particuliers

									Select
										    Case StringInStr($numdirmetier,"790") And StringInStr($dirpitr, "RG-PITR_CM_DIR-BP_DS_REC_RE_RC") AND StringInStr($dirid, "_SM_DET" )  ;Dir-790
									   ;RG-GAUB_SM_DET; RG-PITR_CM_DIR-BP_DS_REC_RE_RC; SG-GAUB_ACCES_EVS; USR_BP_GUICHET_GENE;USR_BP_GUICHET_ESPACE_CO; RG-GAUB_XXXXXX; SG-GAUB_XXXXXX
												Global $regategaub = ""
												$regategaub = InputBox("Regate (6 chiffres) ?", "ex: 694100" & @CRLF & "RG-GAUB_[regate]_GESTION_GROUPE" & @CRLF & "SG-GAUB_[regate]_PRIVE" & @CRLF & "Oui => (REC/RE,DS et CECI)" & @CRLF & "Non => (CCPRO,GCPRO,GDCPRO,GESPRO,RCPART)", "")
												_ad_addusertogroup("RG-GAUB_" & $regategaub & "_GESTION_GROUPE", $idrh1) ;RG-gaub_xxxxxx
												_ad_addusertogroup("SG-GAUB_" & $regategaub & "_PRIVE", $idrh1) ; SG-gaub_xxxxxx
												_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
												_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)
												_ad_addusertogroup("USR_BP_GUICHET_ESPACE_CO", $idrh1)

										    Case StringInStr($numdirmetier,"733") And StringInStr($dirpitr, "RG-PITR_CM_DIR-BP_DS_REC_RE_RC") AND StringInStr($dirid, "_SM_DET" ) ;Dir-733
											   ;RG-GAUB_SM_DET; RG-PITR_CM_DIR-BP_DS_REC_RE_RC; SG-GAUB_ACCES_EVS; USR_BP_GUICHET_GENE;USR_BP_GUICHET_ESPACE_CO; SG-GAUB_XXXXXX
											   Global $regategaub = ""
												$regategaub = InputBox("Regate (6 chiffres) ?", "ex: 694100" & @CRLF & "RG-GAUB_[regate]_GESTION_GROUPE" & @CRLF & "SG-GAUB_[regate]_PRIVE" & @CRLF & "Oui => (REC/RE,DS et CECI)" & @CRLF & "Non => (CCPRO,GCPRO,GDCPRO,GESPRO,RCPART)", "")
											    ;_ad_addusertogroup("RG-GAUB_" & $regategaub & "_GESTION_GROUPE", $idrh1) ;RG-gaub_xxxxxx
												_ad_addusertogroup("SG-GAUB_" & $regategaub & "_PRIVE", $idrh1) ; SG-gaub_xxxxxx
												_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
												_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)
												_ad_addusertogroup("USR_BP_GUICHET_ESPACE_CO", $idrh1)


											Case StringInStr($numdirmetier,"789") And StringInStr($dirpitr,"RG-PITR_CM_DIR-BP_CSI_CSP") AND StringInStr($dirid, "_SM_DET") ;Dir-789
											   ;RG-GAUB_SM_DET;RG-PITR_CM_DIR-BP_CSI_CSP; RG-PITR_CM_COBA_Agent
											_ad_addusertogroup("RG-PITR_CM_COBA_Agent",$idrh1)

											Case StringInStr($numdirmetier,"731") And StringInStr($dirpitr, "RG-PITR_CM_DIR-BP_COCLI_COFI_Itinerant") AND StringInStr($dirid, "_SM_DET") ;Dir-731
											   ;RG-GAUB_SM_DET; RG-PITR_CM_DIR-BP_COCLI_COFI_Itinerant; RG-PITR_CM_COBA_Agent
										    _ad_addusertogroup("RG-PITR_CM_COBA_Agent", $idrh1)

										    Case StringInStr($numdirmetier,"774") And StringInStr($dirpitr, "RG-PITR_CM_DIR-TERTIAIRE_MAISON_HABITAT") AND StringInStr($dirid, "_SM_Sites Tertiaires RLP") ;Dir-774
											 ;RG-BPXX_SM_Sites Tertiaires RLP; RG-PITR_CM_DIR-TERTIAIRE_MAISON_HABITAT  ; => pas de groupe AD supplémentaire !  RG-BPLY_SM_Sites Tertiaires RLP

											Case StringInStr($numdirmetier,"788") And StringInStr($dirpitr, "RG-PITR_CM_DIR-BP_BPE_RCP") AND StringInStr($dirid, "_SM_DET") ;Dir-788
											   ;RG-GAUB_SM_DET; RG-PITR_CM_DIR-BP_BPE_RCP; SG-GAUB_ACCES_EVS; USR_BP_GUICHET_GENE
												_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
												_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)

											Case StringInStr($numdirmetier,"732") And StringInStr($dirpitr, "RG-PITR_CM_DIR-BP_CCPRO_GCPRO") AND StringInStr($dirid, "_SM_DET"); Dir-732
											   ;RG-GAUB_SM_DET; RG-PITR_CM_DIR-BP_CCPRO_GCPRO; SG-GAUB_ACCES_EVS; USR_BP_GUICHET_GENE
												_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
												_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)


										    Case StringInStr($dirid, "_SM_Conseil Bancaire") ;Dir-19
											 ;RG-[EAID]_SM_Conseil Bancaire; SG-GAUB_ACCES_EVS; USR_BP_GUICHET_GENE
												_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
												_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)
												_ad_addusertogroup("RG-PITR_CM_COBA_Agent", $idrh1)


									EndSelect

									#EndRegion Cas Particuliers

									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									affiche_remove_groupes()
								EndIf
							#EndRegion nouvelles Directives

			#EndRegion traitement des Directives

					 EndIf

						Global $groupsidrh1 = _ad_getusergroups($idrh1)
						If @error > 0 Then
						Else
							_arraysort($groupsidrh1, 0, 1)
							_arraydelete($groupsidrh1, 0)
							Global $groupidrh_final = ""
							Global $groupidrh1_add = $groupsidrh1
							For $k = 0 To UBound($groupsidrh1) - 1
								$groupidrh = $groupsidrh1[$k]
								$groupidrh = StringSplit($groupidrh, ",")
								$groupidrh_int = $groupidrh[1]
								$groupidrh_int = StringReplace($groupidrh_int, "CN=", "")
								$groupidrh_final = $groupidrh_final & $groupidrh_int & @CRLF
							Next
							$groupsidrh1 = _arraytostring($groupsidrh1)
						EndIf

				  if $dirapplied=1 then
				  _GUICtrlRichEdit_SetCharColor($aff, 0x00ff0000)      ;   set color blue (BBGGRR format)
				  _GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Directive appliquée: " & $scomboreaddirectives & "   >" & $idrh1 & @crlf & $actiongroups & @CRLF );, 1)
				  _GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
				  _GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  >Groupes user Srce:  " & $idrh1 & @CRLF & $groupidrh_final & @CRLF);, 1)
				  Else
				  _GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
				  _GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Directive non appliquée: " & $scomboreaddirectives & "   >" & $idrh1 & @crlf & $actiongroups & @CRLF  & "	  Rollback de l'OU d'origine: " & $ousourcedir & @CRLF);, 1) ; $restoredirectives
				  _GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
				  _GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  >Groupes user Srce:  " & $idrh1 & @CRLF & $groupidrh_final & @CRLF);, 1)
				  endif

				  if $actiongroupsar<>"" Then
				  _GUICtrlRichEdit_SetCharColor($aff, 0x00ff0000)      ;   set color blue (BBGGRR format)
				  _GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  > Groupes retirés manuellement, après [Apply] de la Directive Metier [ " & $scomboreaddirectives & " ] , pour user Srce:  " & $idrh1 & @CRLF & $actiongroupsar & @CRLF);, 1)
				  _GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
				  EndIf

				ElseIf $t = 7 Then
					SplashTextOn("", "Abandon ! " & $scomboreaddirectives & " !", 1100, 100, -1, -1, 1, -1, 13, 600)
					Sleep(1000)
					SplashOff()
					$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  Annulation 'Directive Metier' " & $scomboreaddirectives
					_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
					_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  Abandon 'Directive Metier' " & $scomboreaddirectives & @CRLF);, 1)
					_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
					$scomboreaddirectives = ""
				EndIf
				GUIDelete($hguidir)
				ExitLoop
		EndSwitch
	WEnd
EndFunc

Func auto_remove() ;nettoyage ancienne Directive si detecté
	$defautdc = $defautdcinit

	Global $groupsidrh1 = _ad_getusergroups($idrh1)
	_arraysort($groupsidrh1, 0, 1)
	Global $actiongroups = ""

	If IsArray($groupsidrh1) = 1 Then
		For $z = 1 To $groupsidrh1[0]
			$groupsidrh2 = StringSplit($groupsidrh1[$z], ",")
			$groupsidrh3 = $groupsidrh2[1]
			$groupsidrh3 = StringTrimLeft($groupsidrh3, 3)
			If StringInStr($groupsidrh3, "SM_Sites Tertiaires RLP") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "SM_Conseil Bancaire") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "SM_Siege_Enseigne") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "SM_Siege Banque") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "SM_Siege Banque_LBPGP") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "SM_Siege Banque_FondEcranFixe") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "SM_DISFE") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "SM_Siege Banque_LBPF") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "SM_Siege Banque_LBPAI") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "Utilisateurs_Postes_SIA") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "SM_RF_INDET") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "SM_Management Commercial Unique") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "SM_DET") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "SM_Centres Financiers et Centres Nationaux") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "SM_GUICHET") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_COBA_Agent") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_VIR_COM1_IE11") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "_GESTION_GROUPE") AND StringInStr($groupsidrh3, "GAUB") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "_PRIVE") AND StringInStr($groupsidrh3, "GAUB") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_VIR_COM1_IE11_NOREDIRECTFOLDER") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "SG-GAUB_ACCES_EVS") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "USR_BP_GUICHET_GENE") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "USR_BP_CAISSE_GENE") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-BP_COCLI_COFI_Itinerant") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-BP_CCPRO_GCPRO") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-BP_DS_REC_RE_RC") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-STD-W10") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-DEV-W10") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-BEL-DEV-MEDIAS-W10") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-TERTIAIRE_DIRECTEURS_ASS_DIR") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-TERTIAIRE_FILIERE RH_PSST") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-TERTIAIRE_STANDARD_STD") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-TERTIAIRE_COMMUNICATION") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-TERTIAIRE_3631_EXP_CLIENT") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-TERTIAIRE_COMMUNICATION") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-TERTIAIRE_STANDARD_BANCAIRE") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-TERTIAIRE_FIDUCIAIRE_TRANSPORT") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-TERTIAIRE_EBR_SUPPORT-FORMATEUR") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-TERTIAIRE_EBR_INGE_FORMATEUR") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-TERTIAIRE_EBR_MOA_SI_FORMATION") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_IG") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-LBP_BPE_LBPIC") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-POLE_ASSURANCE_LBPA") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_DFI_DHA") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_DFI_DFIS") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_DFI_DDC") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_DPC") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-TERTIAIRE_EXPERT_FINANCE_COMPTA") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_SG") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-POLE_ASSURANCE_CNAH") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_DEDT_STD") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-TERTIAIRE_MAISON_HABITAT") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-TERTIAIRE_CONTROLE_GESTION") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-TERTIAIRE_GEOMARKETING") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_DJ") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_PPH_DBD") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_DRH") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_DIRCOM") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_DRG_DCP") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-TRANSACTIS") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_DRG_STD") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-LBP_BFI_STD") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-LBP_BFI_MOE") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-LBP_BFI_SDM") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-BP_BPE_RCP") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-BP_CSI_CSP") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-BP_DS_REC_RE_RC") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_DEDT_MOP") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-MANAGER_CF_Standard") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_DFI_DCG") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_DCG") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_DIDD") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_DO") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-SIEGE_BANQUE_DP") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-CF_DOVM") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-CF_Standard") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "RG-PITR_CM_DIR-CF_Rumba") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			EndIf
			If StringInStr($groupsidrh3, "USR_BP_GUICHET_ESPACE_CO") Then
				_ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			 EndIf
   ;VirtuOS 28-01-2022 (Vir_01 à Vir_015)
			If StringInStr($groupsidrh3,"RG-PITR_CM_CF_SCLI_Agent") then
			   _ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			 EndIf
			If StringInStr($groupsidrh3,"RG-PITR_CM_CF_Credit-Immo_Agent") then
			   _ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			 EndIf
		   If StringInStr($groupsidrh3,"RG-PITR_CM_CF_SQA_Agent") then
			   _ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			 EndIf
			If StringInStr($groupsidrh3,"RG-PITR_CM_CF_Succession_Agent") then
			   _ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			 EndIf
			If StringInStr($groupsidrh3,"RG-PITR_CM_CF_DRCR_Agent") then
			   _ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			 EndIf
			If StringInStr($groupsidrh3,"RG-PITR_CM_CF_SDEV_Agent") then
			   _ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			 EndIf
			If StringInStr($groupsidrh3,"RG-PITR_CM_CF_CNMR-3639_Agent") then
			   _ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			 EndIf
			If StringInStr($groupsidrh3,"RG-PITR_CM_CF_CNBEL_Agent") then
			   _ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			 EndIf
			If StringInStr($groupsidrh3,"RG-PITR_CM_CF_Appui_Agent") then
			   _ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			 EndIf
			If StringInStr($groupsidrh3,"RG-PITR_CM_CF_CNMR-CI_Agent") then
			   _ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			 EndIf
			If StringInStr($groupsidrh3,"RG-PITR_CM_CF_VDC_Agent") then
			   _ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			 EndIf
			If StringInStr($groupsidrh3,"RG-PITR_CM_CF_DSCS_Agent") then
			   _ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			 EndIf
			If StringInStr($groupsidrh3,"RG-PITR_CM_CF_SBA_Agent") then
			   _ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			 EndIf

			;V15 RG-[EAID]_SM_Conseil Bancaire;RG-PITR_CM_COBA_Agent

   ;VirtuOS 28-01-2022 (Vir_01 à Vir_015)

   ;Vendeur Formateur LPM 09/02/2022 >
			If StringInStr($groupsidrh3,"RG-PITR_VIR_COM1_IE11_NOREDIRECTFOLDER") Then
			   _ad_removeuserfromgroup($groupsidrh3, $idrh1)
				$actiongroups = $actiongroups & $groupsidrh3 & " ; "
			 EndIf
   ;< Vendeur Formateur LPM 09/02/2022

		Next
	 EndIf
	Return $actiongroups
EndFunc

Func liste_directives() ;Liste globale des Directives sous forme Matricielle
	Dim $array[0]
	Dim $array2[0]

#Region Liste des Directives
		_arrayadd($array2, "01")
		_arrayadd($array2, "02")
		_arrayadd($array2, "03 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "04 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "05")
		_arrayadd($array2, "06")
		_arrayadd($array2, "07 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "08")
		_arrayadd($array2, "09 RG-[EAID]_SM_Conseil Bancaire")
		_arrayadd($array2, "10 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "11 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "12 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "13 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "14 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "15 RG-[EAID]_SM_Conseil Bancaire")
		_arrayadd($array2, "16 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "17")
		_arrayadd($array2, "18")
		_arrayadd($array2, "19 RG-[EAID]_SM_Conseil Bancaire")
		_arrayadd($array2, "20 RG-[EAID]_SM_Conseil Bancaire")
		_arrayadd($array2, "21 RG-[EAID]_SM_Conseil Bancaire")
		_arrayadd($array2, "22 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "23")
		_arrayadd($array2, "24 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "25 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "26 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "27 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "28 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "29 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "30 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "31 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "32 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "33 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "34 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "35 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "36 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "37 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "38 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "39 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "40 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "41 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "42")
		_arrayadd($array2, "43 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "44 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "45")
		_arrayadd($array2, "46")
		_arrayadd($array2, "47")
		_arrayadd($array2, "48")
		_arrayadd($array2, "49 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "50")
		_arrayadd($array2, "51")
		_arrayadd($array2, "52")
		_arrayadd($array2, "53")
		_arrayadd($array2, "54 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "55 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "56")
		_arrayadd($array2, "57")
		_arrayadd($array2, "58")
		_arrayadd($array2, "59")
		_arrayadd($array2, "60")
		_arrayadd($array2, "61")
		_arrayadd($array2, "62")
		_arrayadd($array2, "63")
		_arrayadd($array2, "64")
		_arrayadd($array2, "65")
		_arrayadd($array2, "66")
		_arrayadd($array2, "67")
		_arrayadd($array2, "68 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "69")
		_arrayadd($array2, "70")
		_arrayadd($array2, "71")
		_arrayadd($array2, "72")
		_arrayadd($array2, "73")
		_arrayadd($array2, "74")
		_arrayadd($array2, "75")
		_arrayadd($array2, "76")
		_arrayadd($array2, "77")
		_arrayadd($array2, "78 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "79 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "80")
		_arrayadd($array2, "81")
		_arrayadd($array2, "82")
		_arrayadd($array2, "83")
		_arrayadd($array2, "84 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "85")
		_arrayadd($array2, "86")
		_arrayadd($array2, "87 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "88 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "89 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "90 RG-[EAID]_SM_Siege Banque_FondEcranFixe")
		_arrayadd($array2, "91 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "92")
		_arrayadd($array2, "93")
		_arrayadd($array2, "94")
		_arrayadd($array2, "95")
		_arrayadd($array2, "96")
		_arrayadd($array2, "97")
		_arrayadd($array2, "98")
		_arrayadd($array2, "99 RG-[EAID]_SM_Siege Banque_FondEcranFixe")
		_arrayadd($array2, "100")
		_arrayadd($array2, "101")
		_arrayadd($array2, "102 RG-[EAID]_SM_Siege Banque_FondEcranFixe")
		_arrayadd($array2, "103")
		_arrayadd($array2, "104 RG-[EAID]_SM_Siege Banque_FondEcranFixe")
		_arrayadd($array2, "105 RG-[EAID]_SM_Siege Banque_FondEcranFixe")
		_arrayadd($array2, "106")
		_arrayadd($array2, "107")
		_arrayadd($array2, "108 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "109")
		_arrayadd($array2, "110")
		_arrayadd($array2, "111 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "112 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "113")
		_arrayadd($array2, "114")
		_arrayadd($array2, "115 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "116 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "117 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "118")
		_arrayadd($array2, "119")
		_arrayadd($array2, "120 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "121")
		_arrayadd($array2, "122")
		_arrayadd($array2, "123")
		_arrayadd($array2, "124")
		_arrayadd($array2, "125")
		_arrayadd($array2, "126")
		_arrayadd($array2, "127")
		_arrayadd($array2, "128 RG-[EAID]_SM_DISFE")
		_arrayadd($array2, "129 RG-[EAID]_SM_DISFE")
		_arrayadd($array2, "130")
		_arrayadd($array2, "131 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "132 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "133")
		_arrayadd($array2, "134")
		_arrayadd($array2, "135")
		_arrayadd($array2, "136")
		_arrayadd($array2, "137")
		_arrayadd($array2, "138")
		_arrayadd($array2, "139")
		_arrayadd($array2, "140")
		_arrayadd($array2, "141")
		_arrayadd($array2, "142")
		_arrayadd($array2, "143")
		_arrayadd($array2, "144")
		_arrayadd($array2, "145 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "146")
		_arrayadd($array2, "147")
		_arrayadd($array2, "148")
		_arrayadd($array2, "149")
		_arrayadd($array2, "150")
		_arrayadd($array2, "151")
		_arrayadd($array2, "152 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "153")
		_arrayadd($array2, "154")
		_arrayadd($array2, "155 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "156")
		_arrayadd($array2, "157 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "158")
		_arrayadd($array2, "159")
		_arrayadd($array2, "160")
		_arrayadd($array2, "161")
		_arrayadd($array2, "162")
		_arrayadd($array2, "163 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "164")
		_arrayadd($array2, "165")
		_arrayadd($array2, "166")
		_arrayadd($array2, "167")
		_arrayadd($array2, "168")
		_arrayadd($array2, "169")
		_arrayadd($array2, "170")
		_arrayadd($array2, "171")
		_arrayadd($array2, "172")
		_arrayadd($array2, "173")
		_arrayadd($array2, "174")
		_arrayadd($array2, "175")
		_arrayadd($array2, "176")
		_arrayadd($array2, "177")
		_arrayadd($array2, "178")
		_arrayadd($array2, "179 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "180 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "181 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "182")
		_arrayadd($array2, "183")
		_arrayadd($array2, "184 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "185 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "186 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "187 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "188")
		_arrayadd($array2, "189")
		_arrayadd($array2, "190")
		_arrayadd($array2, "191 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "192 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "193 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "194 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "195")
		_arrayadd($array2, "196")
		_arrayadd($array2, "197")
		_arrayadd($array2, "198")
		_arrayadd($array2, "199 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "200 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "201 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "202 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "203 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "204 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "205 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "206 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "207")
		_arrayadd($array2, "208")
		_arrayadd($array2, "209")
		_arrayadd($array2, "210")
		_arrayadd($array2, "211")
		_arrayadd($array2, "212")
		_arrayadd($array2, "213")
		_arrayadd($array2, "214")
		_arrayadd($array2, "215 RG-[EAID]_SM_DISFE")
		_arrayadd($array2, "216")
		_arrayadd($array2, "217")
		_arrayadd($array2, "218")
		_arrayadd($array2, "219")
		_arrayadd($array2, "220")
		_arrayadd($array2, "221")
		_arrayadd($array2, "222")
		_arrayadd($array2, "223 RG-[EAID]_SM_DISFE")
		_arrayadd($array2, "224")
		_arrayadd($array2, "225 RG-[EAID]_SM_DISFE")
		_arrayadd($array2, "226")
		_arrayadd($array2, "227")
		_arrayadd($array2, "228 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "229 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "230 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "231 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "232")
		_arrayadd($array2, "233")
		_arrayadd($array2, "234 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "235")
		_arrayadd($array2, "236 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "237")
		_arrayadd($array2, "238")
		_arrayadd($array2, "239 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "240 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "241")
		_arrayadd($array2, "242")
		_arrayadd($array2, "243")
		_arrayadd($array2, "244 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "245")
		_arrayadd($array2, "246")
		_arrayadd($array2, "247")
		_arrayadd($array2, "248")
		_arrayadd($array2, "249")
		_arrayadd($array2, "250")
		_arrayadd($array2, "251")
		_arrayadd($array2, "252 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "253 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "254")
		_arrayadd($array2, "255 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "256")
		_arrayadd($array2, "257 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "258")
		_arrayadd($array2, "259 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "260 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "261 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "262 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "263 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "264 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "265 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "266 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "267 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "268")
		_arrayadd($array2, "269")
		_arrayadd($array2, "270 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "271 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "272 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "273 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "274")
		_arrayadd($array2, "275")
		_arrayadd($array2, "276 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "277 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "278")
		_arrayadd($array2, "279")
		_arrayadd($array2, "280")
		_arrayadd($array2, "281")
		_arrayadd($array2, "282")
		_arrayadd($array2, "283 RG-[EAID]_SM_DISFE")
		_arrayadd($array2, "284 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "285")
		_arrayadd($array2, "286")
		_arrayadd($array2, "287")
		_arrayadd($array2, "288")
		_arrayadd($array2, "289")
		_arrayadd($array2, "290")
		_arrayadd($array2, "291 RG-[EAID]_SM_DISFE")
		_arrayadd($array2, "292")
		_arrayadd($array2, "293")
		_arrayadd($array2, "294 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "295")
		_arrayadd($array2, "296")
		_arrayadd($array2, "297")
		_arrayadd($array2, "298")
		_arrayadd($array2, "299")
		_arrayadd($array2, "300 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "301 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "302")
		_arrayadd($array2, "303 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "304")
		_arrayadd($array2, "305")
		_arrayadd($array2, "306 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "307 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "308 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "309 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "310 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "311 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "312 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "313 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "314 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "315")
		_arrayadd($array2, "316 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "317 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "318 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "319 RG-[EAID]_SM_Siege Banque_FondEcranFixe")
		_arrayadd($array2, "320 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "321")
		_arrayadd($array2, "322 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "323")
		_arrayadd($array2, "324 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "325 RG-[EAID]_SM_DISFE")
		_arrayadd($array2, "326")
		_arrayadd($array2, "327")
		_arrayadd($array2, "328")
		_arrayadd($array2, "329")
		_arrayadd($array2, "330")
		_arrayadd($array2, "331 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "332 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "333 RG-[EAID]_SM_DET")
		_arrayadd($array2, "334 RG-[EAID]_SM_DET")
		_arrayadd($array2, "335 RG-[EAID]_SM_DET")
		_arrayadd($array2, "336")
		_arrayadd($array2, "337 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "338 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "339 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "340 RG-[EAID]_SM_Siege Banque_FondEcranFixe")
		_arrayadd($array2, "341 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "342 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "343 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "344 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "345 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "346 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "347 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "348 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "349 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "350")
		_arrayadd($array2, "351 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "352")
		_arrayadd($array2, "353")
		_arrayadd($array2, "354 RG-[EAID]_SM_Siege Banque_FondEcranFixe")
		_arrayadd($array2, "355 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "356 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "357")
		_arrayadd($array2, "358 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "359 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "360 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "361 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "362 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "363 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "364 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "365")
		_arrayadd($array2, "366 RG-[EAID]_SM_Siege Banque_FondEcranFixe")
		_arrayadd($array2, "367 RG-[EAID]_SM_Siege Banque_FondEcranFixe")
		_arrayadd($array2, "368 RG-[EAID]_SM_Siege Banque_FondEcranFixe")
		_arrayadd($array2, "369")
		_arrayadd($array2, "370")
		_arrayadd($array2, "371")
		_arrayadd($array2, "372")
		_arrayadd($array2, "373 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "374")
		_arrayadd($array2, "375")
		_arrayadd($array2, "376")
		_arrayadd($array2, "377")
		_arrayadd($array2, "378")
		_arrayadd($array2, "379 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "380")
		_arrayadd($array2, "381")
		_arrayadd($array2, "382")
		_arrayadd($array2, "383")
		_arrayadd($array2, "384")
		_arrayadd($array2, "385")
		_arrayadd($array2, "386")
		_arrayadd($array2, "387")
		_arrayadd($array2, "388 RG-[EAID]_SM_DISFE")
		_arrayadd($array2, "389")
		_arrayadd($array2, "390")
		_arrayadd($array2, "391")
		_arrayadd($array2, "392 RG-[EAID]_SM_DISFE")
		_arrayadd($array2, "393 RG-[EAID]_SM_DISFE")
		_arrayadd($array2, "394 RG-[EAID]_SM_DISFE")
		_arrayadd($array2, "395 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "396 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "397 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "398 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "399")
		_arrayadd($array2, "400")
		_arrayadd($array2, "401")
		_arrayadd($array2, "402")
		_arrayadd($array2, "403 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "404 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "405 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "406 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "407 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "408 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "409 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "410 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "411 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "412 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "413 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "414 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "415 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "416 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "417 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "418 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "419 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "420 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "421 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "422 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "423 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "424 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "425 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "426 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "427 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "428 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "429 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "430 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "431 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "432 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "433 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "434 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "435 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "436 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "437 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "438 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "439 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "440 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "441 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "442 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "443 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "444 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "445 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "446 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "447 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "448 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "449 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "450 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "451 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "452 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "453 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "454 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "455 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "456 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "457 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "458 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "459 RG-[EAID]_SM_Siege Banque_FondEcranFixe")
		_arrayadd($array2, "460 RG-[EAID]_SM_Siege Banque_FondEcranFixe")
		_arrayadd($array2, "461 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "462 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "463 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "464 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "465 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "466 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "467 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "468 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "469 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "470 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "471 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "472 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "473 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "474 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "475 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "476 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "477 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "478 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "479 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "480 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "481 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "482 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "483 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "484 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "485 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "486 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "487 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "488 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "489 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "490 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "491 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "492 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "493")
		_arrayadd($array2, "494 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "495 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "496 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "497 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "498 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "499 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "500 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "501 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "502 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "503 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "504 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "505 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "506 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "507 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "508")
		_arrayadd($array2, "509 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "510 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "511 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "512 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "513 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "514")
		_arrayadd($array2, "515")
		_arrayadd($array2, "516 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "517 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "518 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "519")
		_arrayadd($array2, "520")
		_arrayadd($array2, "521 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "522 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "523")
		_arrayadd($array2, "524 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "525")
		_arrayadd($array2, "526 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "527")
		_arrayadd($array2, "528 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "529 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "530")
		_arrayadd($array2, "531")
		_arrayadd($array2, "532")
		_arrayadd($array2, "533 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "534 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "535")
		_arrayadd($array2, "536")
		_arrayadd($array2, "537 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "538 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "539 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "540 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "541")
		_arrayadd($array2, "542 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "543 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "544 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "545 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "546")
		_arrayadd($array2, "547")
		_arrayadd($array2, "548 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "549 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "550 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "551 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "552 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "553 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "554 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "555 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "556 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "557 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "558 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "559")
		_arrayadd($array2, "560")
		_arrayadd($array2, "561")
		_arrayadd($array2, "562")
		_arrayadd($array2, "563")
		_arrayadd($array2, "564")
		_arrayadd($array2, "565")
		_arrayadd($array2, "566")
		_arrayadd($array2, "567 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "568 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "569 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "570 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "571 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "572 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "573 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "574 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "575 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "576 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "577 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "578 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "579 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "580 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "581 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "582 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "583")
		_arrayadd($array2, "584")
		_arrayadd($array2, "585 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "586 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "587 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "588 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "589 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "590 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "591 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "592 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "593 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "594 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "595 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "596 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "597 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "598 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "599 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "600 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "601 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "602 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "603 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "604 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "605 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "606 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "607 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "608 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "609 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "610 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "611 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "612 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "613 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "614 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "615 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "616 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "617 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "618 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "619 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "620 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "621 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "622")
		_arrayadd($array2, "623")
		_arrayadd($array2, "624")
		_arrayadd($array2, "625 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "626")
		_arrayadd($array2, "627 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "628 RG-[EAID]_SM_DET")
		_arrayadd($array2, "629 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "630 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "631 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "632 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "633 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "634 RG-[EAID]_SM_DISFE")
		_arrayadd($array2, "635 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "636 RG-[EAID]_SM_Sites Tertiaires RLP")
		_arrayadd($array2, "637 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "638 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "639 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "640 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "641 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "642 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "643 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "644 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "645 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "646 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "647 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "648 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "649 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "650 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "651 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "652 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "653 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "654 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "655 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "656 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "657 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "658 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "659 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "660 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "661 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "662 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "663 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "664 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "665 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "666 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "667 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "668 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "669 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "670 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "671 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "672 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "673 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "674 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "675 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "676 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "677 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "678 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "679 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "680 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "681 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "682 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "683 RG-[EAID]_SM_Siege Banque_LBPGP")
		_arrayadd($array2, "684 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "685 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "686 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "687 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "688 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "689 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "690 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "691 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "692 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "693 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "694 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "695 RG-[EAID]_SM_Siege Banque_FondEcranFixe")
		_arrayadd($array2, "696 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "697 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "698 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "699 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "700 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "701 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "702 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "703 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "704 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "705 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "706 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "707 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "708 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "709 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "710 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "711 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "712 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "713 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "714 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "715 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "716 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "717 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "718 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "719 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "720 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "721 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "722 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "723 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "724 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "725 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "726 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "727 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "728 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "729 RG-[EAID]_SM_Siege Banque_LBPPF")
		_arrayadd($array2, "730 RG-[EAID]_SM_Centres Financiers et Centres Nationaux")
		_arrayadd($array2, "731 RG-[EAID]_SM_DET;RG-PITR_CM_DIR-BP_COCLI_COFI_Itinerant")
		_arrayadd($array2, "732 RG-[EAID]_SM_DET;RG-PITR_CM_DIR-BP_CCPRO_GCPRO")
		_arrayadd($array2, "733 RG-[EAID]_SM_DET;RG-PITR_CM_DIR-BP_DS_REC_RE_RC")
		_arrayadd($array2, "734")
		_arrayadd($array2, "735")
		_arrayadd($array2, "736 RG-[EAID]_SM_DISFE;RG-PITR_CM_DIR-STD-W10")
		_arrayadd($array2, "737 RG-[EAID]_SM_DISFE;RG-PITR_CM_DIR-DEV-W10")
		_arrayadd($array2, "738 RG-[EAID]_SM_DISFE;RG-PITR_CM_DIR-BEL-DEV-MEDIAS-W10")
		_arrayadd($array2, "739 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "740 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "741 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "742 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "743 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "744 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "745 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "746 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "747 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "748 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "749 RG-[EAID]_SM_Siege Banque")
		_arrayadd($array2, "750 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_DIRECTEURS_ASS_DIR")
		_arrayadd($array2, "751 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_FILIERE RH_PSST")
		_arrayadd($array2, "752 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_STANDARD_STD")
		_arrayadd($array2, "753 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_COMMUNICATION")
		_arrayadd($array2, "754 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_3631_EXP_CLIENT")
		_arrayadd($array2, "755")
		_arrayadd($array2, "756 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_STANDARD_BANCAIRE")
		_arrayadd($array2, "757 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-TERTIAIRE_FIDUCIAIRE_TRANSPORT")
		_arrayadd($array2, "758 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_EBR_SUPPORT-FORMATEUR")
		_arrayadd($array2, "759")
		_arrayadd($array2, "760 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_EBR_INGE_FORMATEUR")
		_arrayadd($array2, "761 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_EBR_MOA_SI_FORMATION")
		_arrayadd($array2, "762")
		_arrayadd($array2, "763 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_IG")
		_arrayadd($array2, "764 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-LBP_BPE_LBPIC")
		_arrayadd($array2, "765 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-POLE_ASSURANCE_LBPA")
		_arrayadd($array2, "766 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_DFI_DHA")
		_arrayadd($array2, "767 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_DFI_DFIS")
		_arrayadd($array2, "768 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_DFI_DDC")
		_arrayadd($array2, "769 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_DPC")
		_arrayadd($array2, "770 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-TERTIAIRE_EXPERT_FINANCE_COMPTA")
		_arrayadd($array2, "771 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_SG")
		_arrayadd($array2, "772 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-POLE_ASSURANCE_CNAH")
		_arrayadd($array2, "773 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_DEDT_STD")
		_arrayadd($array2, "774 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_MAISON_HABITAT")
		_arrayadd($array2, "775 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_CONTROLE_GESTION")
		_arrayadd($array2, "776 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_GEOMARKETING")
		_arrayadd($array2, "777 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_DJ")
		_arrayadd($array2, "778 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_PPH_DBD")
		_arrayadd($array2, "779 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_DRH")
		_arrayadd($array2, "780 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_DIRCOM")
		_arrayadd($array2, "781 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_DRG_DCP")
		_arrayadd($array2, "782 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-TRANSACTIS")
		_arrayadd($array2, "783 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_DRG_STD")
		_arrayadd($array2, "784 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-LBP_BFI_STD")
		_arrayadd($array2, "785 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-LBP_BFI_MOE")
		_arrayadd($array2, "786 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-LBP_BFI_SDM")
		_arrayadd($array2, "787")
		_arrayadd($array2, "788 RG-[EAID]_SM_DET;RG-PITR_CM_DIR-BP_BPE_RCP")
		_arrayadd($array2, "789 RG-[EAID]_SM_DET;RG-PITR_CM_DIR-BP_CSI_CSP")
		_arrayadd($array2, "790 RG-[EAID]_SM_DET;RG-PITR_CM_DIR-BP_DS_REC_RE_RC")
		_arrayadd($array2, "791 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_DEDT_MOP")
		_arrayadd($array2, "792 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_DIR-MANAGER_CF_Standard")
		_arrayadd($array2, "793 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_DFI_DCG")
		_arrayadd($array2, "794 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_DCG")
		_arrayadd($array2, "795 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_DIDD")
		_arrayadd($array2, "796 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_DO")
		_arrayadd($array2, "797 RG-[EAID]_SM_Siege Banque;RG-PITR_CM_DIR-SIEGE_BANQUE_DP")
		_arrayadd($array2, "798 RG-[EAID]_SM_Siege Banque_LBPF")
		_arrayadd($array2, "799 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_DIR-CF_DOVM")
		_arrayadd($array2, "800 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_DIR-CF_Standard")
		_arrayadd($array2, "801 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_DIR-CF_Rumba")
		;VirtuOS 28-01-2022
		_arrayadd($array2, "V01 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_CF_SCLI_Agent")
		_arrayadd($array2, "V02 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_CF_Credit-Immo_Agent")
		_arrayadd($array2, "V03 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_CF_SQA_Agent")
		_arrayadd($array2, "V04 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_CF_Succession_Agent")
		_arrayadd($array2, "V05 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_CF_Succession_Agent")
		_arrayadd($array2, "V06 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_CF_DRCR_Agent")
		_arrayadd($array2, "V07 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_CF_SDEV_Agent")
		_arrayadd($array2, "V08 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_CF_CNMR-3639_Agent")
		_arrayadd($array2, "V09 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_CF_CNBEL_Agent")
		_arrayadd($array2, "V10 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_CF_Appui_Agent")
		_arrayadd($array2, "V11 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_CF_CNMR-CI_Agent")
		_arrayadd($array2, "V12 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_CF_VDC_Agent")
		_arrayadd($array2, "V13 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_CF_DSCS_Agent")
		_arrayadd($array2, "V14 RG-[EAID]_SM_Centres Financiers et Centres Nationaux;RG-PITR_CM_CF_SBA_Agent")
		_arrayadd($array2, "V15 RG-[EAID]_SM_Conseil Bancaire;RG-PITR_CM_COBA_Agent")
		;VirtuOS 29-03-2022
	#EndRegion Liste des Directives

	For $i = 0 To UBound($array2) - 1
		If StringLen($array2[$i]) <= 4 Then
		Else
			_arrayadd($array, $array2[$i])
		EndIf
	Next

#Region mise en forme $array2 => $array
		_arrayinsert($array, 0, "_SM_RF_INDET")
		_arrayinsert($array, 0, "_SM_Utilisateurs_Postes_SIA")
		_arrayinsert($array, 0, "_SM_Centres Financiers et Centres Nationaux")
		_arrayinsert($array, 0, "_SM_DISFE")
		_arrayinsert($array, 0, "_SM_Siege Banque_FondEcranFixe")
		_arrayinsert($array, 0, "_SM_Siege Banque_LBPAI")
		_arrayinsert($array, 0, "_SM_Siege Banque_LBPF")
		_arrayinsert($array, 0, "_SM_Siege Banque_LBPGP")
		_arrayinsert($array, 0, "_SM_Siege Banque")
		_arrayinsert($array, 0, "_SM_DET")
	;	_arrayinsert($array, 0, "_SM_Siege_Enseigne") n'existe plus renommé en RG-[EAID]_SM_Siege Banque
	;	_arrayinsert($array, 0, "_SM_Conseil Bancaire W10")
	;    _arrayinsert($array, 0, "19 _SM_Conseil Bancaire")
		_arrayinsert($array, 0, "_SM_GUICHET (VirtuOS)")
		_arrayinsert($array, 0, "_Vendeur Formateur LPM")
		_arrayinsert($array, 0, "_SM_Management Commercial Unique")
		_arrayinsert($array, 0, "_SM_Sites Tertiaires RLP")
		_arrayinsert($array, 0, "[ Zero Directives ]")
	#EndRegion mise en forme $array2 => $array

Global $directives = _arraytostring($array, "|")
EndFunc

Func directives_DSIBA() ;routine principale des Directives Metiers DSIBA
	$defautdc = $defautdcinit

	#Region IHM Directives
		If $isdct = 0 Then
			MsgBox(0, "warning !", "'Directives DSIBA' valable que pour le domaine: DCT.Adt.Local", 7)
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "'Directives DSIBA' valable que pour le domaine: DCT.Adt.Local " & @CRLF);, 1)
			Return 0
		 EndIf

		Global $ousourcedir = _ad_getobjectattribute($idrh1, "distinguishedName")
		$ousourcedir = StringSplit($ousourcedir, ",")
		$ousourcedir = $ousourcedir[3]
		$ousourcedir = StringTrimLeft($ousourcedir, 3)
		$hguidir = GUICreate("Selection Directive DSIBA et fermer la fenetre [x] ...", 650, 80)
	#EndRegion

	#Region Liste des Directives
		If $directivesDSIBA = "" Then

				Dim $arrayD[0]
#Region mise en forme $array2 => $array
_arrayadd($arrayD,"736 SUPPORT N2") ;  Dir-736	 Directive Standard
_arrayadd($arrayD,"736 DIRECTEUR") ; 	Dir-736  Directive Standard
_arrayadd($arrayD,"736 RESPONSABLE DEPARTEMENT/SERVICE") ;	Dir-736	Directive Standard
_arrayadd($arrayD,"736 CHEF DE PROJET STANDARD") ;  Dir-736	Directive Standard
_arrayadd($arrayD,"736 INFRA SECURITE"); Dir-736	Directive Standard
_arrayadd($arrayD,"736 INFRA SURVEILLANCE") ; Dir-736	Directive Standard
_arrayadd($arrayD,"736 DMR-STANDARD"); Dir-736	Directive Standard
_arrayadd($arrayD,"736 PRODUCTION STANDARD") ; Dir-736	Directive Standard
_arrayadd($arrayD,"736 EXPERTISE METIER GA") ; Dir-736	Directive Standard
_arrayadd($arrayD,"736 PILOTAGE") ; Dir-736	Directive Standard
_arrayadd($arrayD,"736 STANDARD") ; Dir-736	Directive Standard
_arrayadd($arrayD,"737 DEVELOPPEUR") ;	Dir-737	Dir Développeur
_arrayadd($arrayD,"738 DSIBR-BEL-DEV-MEDIAS") ;  Dir-738  Dir BEL DEV MEDIAS
_arrayadd($arrayD,"[ Zero Directives ]")
#EndRegion mise en forme $array2 => $array

Global $directivesDSIBA = _arraytostring($arrayD, "|")
_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Premier lancement Directives Metiers DSIBA !" & @CRLF );, 1)
_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
		EndIf
	#EndRegion Liste des Directives

	$idcombobox = GUICtrlCreateCombo("", 20, 20, 600, 250)
	GUICtrlSetData(-1, $directivesDSIBA, "736 SUPPORT N2")
	GUISetState(@SW_SHOW, $hguidir)
	While 1
		Switch GUIGetMsg()
			Case $gui_event_close
				Global $scomboreaddirectives = GUICtrlRead($idcombobox)
				If StringLen($scomboreaddirectives) > 4 Then
				Else
					Global $dirapplied = 0
					MsgBox(0, "Warning !", " Directive DSIBA inexistante dans cette OU... abandon !", 7)
					_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
					_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & " Directive DSIBA inexistante dans cette OU ... abandon !" & @CRLF);, 1)
					GUIDelete($hguidir)
					ExitLoop
				EndIf
				GUIDelete($hguidir)
				If $scomboreaddirectives = "[ Zero Directives ]" Then
					$t = MsgBox(4, "Retirer Directive DSIBA ?", "Retirer Directive DSIBA pour user Srce: " & $idrh1 & "  ?")
				Else
					$t = MsgBox(4, "Valider Directive DSIBA ?", "Valider la Directive DSIBA : " & @CRLF & $scomboreaddirectives)
				EndIf
				If $t = 6 Then
					SplashTextOn("", "Valdation de la Directive DSIBA:  " & $scomboreaddirectives & " !", 1100, 100, -1, -1, 1, -1, 13, 600)
					Sleep(1200)
					SplashOff()
					$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  Valider une 'Directive'"
					_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  Validation de la 'Directive DSIBA': " & $scomboreaddirectives & @CRLF);, 1)
					Global $dirapplied = 0


						#Region  Directives DSIBA

								If $scomboreaddirectives = "[ Zero Directives ]" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$scomboreaddirectives = " [ Zero Directives ] = Enleve toutes les Directives connues !"
									$dirapplied = 1
								EndIf

							    If StringInStr($scomboreaddirectives ,"736" ) Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "736 RG-[EAID]_SM_DISFE;RG-PITR_CM_DIR-STD-W10", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_DISFE", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										_ad_addusertogroup("RG-PITR_CM_DIR-STD-W10", $idrh1)
										MsgBox(64, "", "User Srce '" & $idrh1 & "' assigné au groupe metier '" & "RG-" & $ou & "_SM_DISFE" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe Metier '" & "RG-" & $ou & "_SM_DISFE" & "' n'existe pas dans cette OU : " & $ou & "  !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User Srce '" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User Srce '" & $idrh1 & "' déjà membre du groupe metier '" & "RG-" & $ou & "_SM_DISFE" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers OU initiale '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Metier Non Appliquée ... (rollback)")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf

							    If StringInStr($scomboreaddirectives ,"737" ) Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user Srce " & $idrh1 & " n'a pas d groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "737 RG-[EAID]_SM_DISFE;RG-PITR_CM_DIR-DEV-W10", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_DISFE", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										_ad_addusertogroup("RG-PITR_CM_DIR-DEV-W10", $idrh1)
										MsgBox(64, "", "User '" & $idrh1 & "' assigné au groupe metier '" & "RG-" & $ou & "_SM_DISFE" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe metier '" & "RG-" & $ou & "_SM_DISFE" & "' n'existe pas dans cette OU: " & $ou & " !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User Srce'" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User Srce'" & $idrh1 & "' déjà membre du groupe metier  '" & "RG-" & $ou & "_SM_DISFE" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
						 			    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers OU '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Non Appliquée... (rollback)")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								 EndIf

							    If StringInStr($scomboreaddirectives ,"738" ) Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "n'a pas de groupes...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Groupes Obsoletes Directive : " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "738 RG-[EAID]_SM_DISFE;RG-PITR_CM_DIR-BEL-DEV-MEDIAS-W10", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_DISFE", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										_ad_addusertogroup("RG-PITR_CM_DIR-BEL-DEV-MEDIAS-W10", $idrh1)
										MsgBox(64, "", "User Srce '" & $idrh1 & "' assigné au groupe metier '" & "RG-" & $ou & "_SM_DISFE" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Groupe metier '" & "RG-" & $ou & "_SM_DISFE" & "' n'existe pas dans cette OU : " & $ou & "  !")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User Srce'" & $idrh1 & "' n'existe pas !")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' déjà affecté au groupe metier '" & "RG-" & $ou & "_SM_DISFE" & "'")
										$dirapplied = 1
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
									EndIf
									If $dirapplied = 0 Then
									    $defautdc = $defautdcinit
										$defautdc = StringReplace($defautdc, ".", ",DC=")
										$defautdc = ",DC=" & $defautdc
										$sobject = _ad_samaccountnametofqdn($idrh1)
										$souinitial = "OU=Utilisateurs,OU=" & $ousourcedir & ",OU=" & StringUpper(StringMid($ousourcedir, 1, 2)) & $defautdc
										$ivalue = _ad_moveobject($souinitial, $sobject)
										If $ivalue = 1 Then
											MsgBox(64, "Active Directory ", $idrh1 & "' replacé vers  '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Non Appliquée... (rollback)")
											ToolTip("synchronising AD", 5, 5)
											Sleep(3000)
											ToolTip("", 5, 5)
										EndIf
										$restoredirectives = StringTrimRight($restoredirectives, 1)
										$groupsidrh1 = StringSplit($restoredirectives, "|")
										If IsArray($groupsidrh1) = 1 Then
											For $z = 1 To $groupsidrh1[0]
												_ad_addusertogroup($groupsidrh1[$z], $idrh1)
											Next
										EndIf
										$actiongroups = StringReplace($actiongroups, " Groupes Obsoletes Directive : ", " Rollback de la Directive Metier Initiale :  ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf



						#EndRegion  Directives DSIBA



						Global $groupsidrh1 = _ad_getusergroups($idrh1)
						If @error > 0 Then
						Else
							_arraysort($groupsidrh1, 0, 1)
							_arraydelete($groupsidrh1, 0)
							Global $groupidrh_final = ""
							Global $groupidrh1_add = $groupsidrh1
							For $k = 0 To UBound($groupsidrh1) - 1
								$groupidrh = $groupsidrh1[$k]
								$groupidrh = StringSplit($groupidrh, ",")
								$groupidrh_int = $groupidrh[1]
								$groupidrh_int = StringReplace($groupidrh_int, "CN=", "")
								$groupidrh_final = $groupidrh_final & $groupidrh_int & @CRLF
							Next
							$groupsidrh1 = _arraytostring($groupsidrh1)
						EndIf

				  if $dirapplied=1 then
				  _GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Directive appliquée: " & $scomboreaddirectives & "   >" & $idrh1 & @crlf & $actiongroups & @CRLF);, 1)
				  _GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  >Groupes user:  " & $idrh1 & @CRLF & $groupidrh_final & @CRLF);, 1)
			   Else
				  _GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
				  _GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Directive non appliquée: " & $scomboreaddirectives & "   >" & $idrh1 & @crlf & $restoredirectives & @CRLF);, 1)
				  _GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  >Groupes user:  " & $idrh1 & @CRLF & $groupidrh_final & @CRLF);, 1)
				  endif

				  if $actiongroupsar<>"" Then
				  _GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  > Groupes retirés manuellement, après avoir appliqué la Directive [ " & $scomboreaddirectives & " ]  pour:  " & $idrh1 & @CRLF & $actiongroupsar & @CRLF);, 1)
				  EndIf

				ElseIf $t = 7 Then
					SplashTextOn("", "Annulation Directive " & $scomboreaddirectives & " !", 1100, 100, -1, -1, 1, -1, 13, 600)
					Sleep(1000)
					SplashOff()
					$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  Annulation de la validation 'Directive'"
					_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
					_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  Annulation 'Directive' :  " & $scomboreaddirectives & @CRLF);, 1)
					$scomboreaddirectives = ""
				EndIf
				GUIDelete($hguidir)
				ExitLoop
		EndSwitch
	WEnd
EndFunc

;/////////////////////////////////////////////////////  Directives Metiers domaine DCT uniquement  /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#EndRegion Directives Metiers

Func move_user()
	$defautdc = $defautdcinit

	$unite = 0
	If $ou = "??" Then
		$unite = 1
		afficheou()
	Else
		$unite = 0
	 EndIf

	If StringLen($ou) < 2 OR StringLen($ou) > 4 Then
		$restoredirectives = StringTrimRight($restoredirectives, 1)
		$groupsidrh1 = StringSplit($restoredirectives, "|")
		If IsArray($groupsidrh1) = 1 Then
			For $z = 1 To $groupsidrh1[0]
				_ad_addusertogroup($groupsidrh1[$z], $idrh1)
			Next
		EndIf
		$actiongroups = ""
		Return 0
	 EndIf

	If $unite = 0 Then
		$defautdc = StringReplace($defautdc, ".", ",DC=")
		$defautdc = ",DC=" & $defautdc
		$sou2 = "OU=Utilisateurs,OU=" & $ou & ",OU=" & StringMid($ou, 1, 2) & $defautdc
	Else
		$ou = $sou
		$defautdc = StringReplace($defautdc, ".", ",DC=")
		$defautdc = ",DC=" & $defautdc
		$sou2 = "OU=Utilisateurs,OU=" & $ou & ",OU=" & StringMid($ou, 1, 2) & $defautdc
	 EndIf

	$idrh1 = StringUpper($idrh1)
	$sobject = _ad_samaccountnametofqdn($idrh1)
	Global $ivalue = _ad_moveobject($sou2, $sobject)

	If $ivalue = 1 Then
		MsgBox(64, "Active Directory ", $idrh1 & "' déplacé de l'OU  '"  &  StringUpper($ousourcedir)  &  "'  vers   '" & StringUpper($ou) & "'")
		_GUICtrlRichEdit_SetCharColor($aff, 0x00ff0000)      ;   set color blue (BBGGRR format)
		_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Active Directory: " & $idrh1 & "  déplacé de l'OU  '" & StringUpper($ousourcedir) &  "'  vers  '" & StringUpper($ou) & "'" & @CRLF);, 1)
		_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
		ToolTip("synchronising AD", 5, 5)
		Sleep(3000)
		ToolTip("", 5, 5)
	ElseIf @error = 1 Then
		MsgBox(64, "Active Directory ", "OU cible  '" & $ou & "' n'existe pas !")
		_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
		_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Active Directory: " & "OU cible  '" & $ou & "' n'existe pas !" & @CRLF);, 1)
		_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
	ElseIf @error = 2 Then
		MsgBox(64, "Active Directory ", "Objet '" & $idrh1 & "' n'existe pas !")
		_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
		_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Active Directory: " & "Objet '" & $idrh1 & "' n'existe pas !" & @CRLF);, 1)
		_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
	ElseIf @error = 3 Then
		MsgBox(64, "Active Directory ", "Objet '" & $idrh1 & "' déjà dans l'OU cible  '" & $ou & "'")
		_GUICtrlRichEdit_SetCharColor($aff, 0x00ff0000)      ;   set color blue (BBGGRR format)
		_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Active Directory: " & "Objet '" & $idrh1 & "' déjà dans l'OU cible  '" & $ou & "'" & @CRLF);, 1)
		_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
	ElseIf @error >= 4 OR @error = -2147352567 Then
		MsgBox(64, "Active Directory ", "Acces refusé pour deplacer user : " & $idrh1)
		_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
		_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Active Directory: " & "accès refusé pour déplacer user : " & $idrh1 & @CRLF);, 1)
		_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
	Else
		MsgBox(0, "warning !", "User Idrh Srce '" & $idrh1 & "' already exists in AD !", 7)
		_GUICtrlRichEdit_SetCharColor($aff, 0x000000ff)      ;   set color red (BBGGRR format)
		_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "User Idrh Srce '" & $idrh1 & "' existe déjà dans l'AD !" & @CRLF);, 1)
		_GUICtrlRichEdit_SetCharColor($aff, 0x00000000)      ;   set color black (BBGGRR format)
	 EndIf

EndFunc

Func afficheou()
	$defautdc = $defautdcinit

	$allou = 1
	treeview_affiche()
	treeselect()

	$sout = StringReplace($sout, "OU=Utilisateurs,", "")
	$sout = StringReplace($sout, "OU=Groupes,", "")
	$sout = StringReplace($sout, "OU=_Support Delegue,", "")
	$sout = StringReplace($sout, "OU=Automates,", "")
	$sout = StringReplace($sout, "OU=Postes de travail,", "")
	$sout = StringReplace($sout, "OU=Postes_SIA,", "")
	$sout = StringReplace($sout, "OU=Serveurs,", "")
	$sout = StringReplace($sout, "OU=Terminaux,", "")
	$sout = StringReplace($sout, "OU=Tablettes,", "")
	$sout = StringReplace($sout, "OU=Postes de travail Hybrid,", "")
	$sout = StringReplace($sout, "OU=Autopilot Domain Join,", "")
	ToolTip("", 5, 5)
	$allou = 0
	Global $sou = StringMid($sout, 4, 4)
	Return $sou
EndFunc

#Region n2

	Func affiche_remove_groupes()
		_arraysort($groupsidrh1, 0, 1)
		Global $actiongroupsar = ""
		Global $aselected[1]
		Global $hgui2 = GUICreate("Selection Role(s) obsolete(s) à retirer [-] de '" & $idrh1 & "' , puis valider avec le bouton [ Apply ] !, ou refermer la fenetre [x] pour abandonner...", 1000, 500)
		Global $hlistbox = _guictrllistbox_create($hgui2, "String upon creation", 10, 30, 980, 470, BitOR($ws_border, $lbs_hasstrings, $ws_vscroll, $lbs_multiplesel))
		_guictrllistbox_beginupdate($hlistbox)
		_guictrllistbox_resetcontent($hlistbox)
		_guictrllistbox_initstorage($hlistbox, 100, 4096)
		_guictrllistbox_endupdate($hlistbox)
		If IsArray($groupsidrh1) = 1 Then
			For $z = 1 To $groupsidrh1[0]
				$groupsidrh2 = StringSplit($groupsidrh1[$z], ",")
				$groupsidrh3 = $groupsidrh2[1]
				$groupsidrh3 = StringTrimLeft($groupsidrh3, 3)
				_guictrllistbox_addstring($hlistbox, $groupsidrh3)
			Next
		EndIf
		If IsArray($groupsidrh1) = 0 Then
			$idcheckboxgroup = 2
			GUIDelete($hgui2)
			MsgBox(0, "Info !", "abandon, pas de groupes pour: " & $idrh1, 7)
			Return 0
		EndIf
		Global $hbutton2 = GUICtrlCreateButton("[ Apply ]", 20, 3, 80, 25)
		GUISetState()
		While 1
			Switch GUIGetMsg()
				Case $hbutton2
					$aselected = _guictrllistbox_getselitems($hlistbox)
					If $aselected[0] = 1 Then
						$sitem = " item"
					Else
						$sitem = " items"
					EndIf
					$sitems = ""
					For $i = 1 To $aselected[0]
						$sitems &= _guictrllistbox_gettext($hlistbox, $aselected[$i]) & @CRLF
					Next
					If $aselected[0] <> 0 Then
						$sitems = StringSplit($sitems, @CRLF)
						For $z = 1 To $sitems[0]
							If StringLen($sitems[$z]) <> 0 Then
								Global $ivalue = _ad_removeuserfromgroup($sitems[$z], $idrh1)
								$actiongroupsar = $actiongroupsar & $sitems[$z] & " ; "
							EndIf
						Next
						GUIDelete($hgui2)
						ToolTip("Synchronizing with AD...", 5, 5)
						Sleep(3000)
						ToolTip("", 5, 5)
						Return $actiongroupsar
						ExitLoop
					Else
						MsgBox(4160, "warning !", "aucune selection... réessayez ou fermer la fenetre... [x]")
					EndIf
				Case $gui_event_close
					GUIDelete($hgui2)
					Return $actiongroupsar
					ExitLoop
			EndSwitch
		WEnd
		GUIDelete($hgui2)
	EndFunc

	Func lbpai_illiade()
		$defautdc = $defautdcinit
		If $isdct = 0 Then
			MsgBox(0, "warning !", "'LBPAI / Illiade' n'est valable que pour le domaine DCT.Adt.Local...", 7)
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "'LBPAI / Illiade' n'est valable que pour le domaine DCT.Adt.Local" & @CRLF);, 1)
			Return 0
		EndIf
		Local $remove = 0
		If _ispressed("10", $hdll) Then
			$remove = 1
		EndIf
		Global $ousourcedir = _ad_getobjectattribute($idrh1, "distinguishedName")
		$ousourcedir = StringSplit($ousourcedir, ",")
		$ousourcedir = $ousourcedir[3]
		$ousourcedir = StringTrimLeft($ousourcedir, 3)
		If $isdct = 1 Then
			$ou = InputBox("default OU for LBPAI/Iliade ?", "ex: MINA, MIPS ... " & @CRLF & "?? : scan all OUs...", $ousourcedir)
			If StringLen($ou) < 2 OR StringLen($ou) > 4 AND NOT ($ou = "virtuos" OR $ou = "??") Then
				Return 0
			EndIf
		Else
			$ou = "??"
		EndIf
		If $ou = "??" Then
			$allou = 1
			treeview_affiche()
			treeselect()
		EndIf
		If $allou = 1 Then
			$ou = ""
			$allou = 0
			$sou2 = $sout
			$ou = $sout
		Else
			$defautdc = StringReplace($defautdc, ".", ",DC=")
			$defautdc = ",DC=" & $defautdc
			$sou2 = "OU=Utilisateurs,OU=" & $ou & ",OU=" & StringMid($ou, 1, 2) & $defautdc
			$idrh1 = StringUpper($idrh1)
		EndIf
		$sobject = _ad_samaccountnametofqdn($idrh1)
		Global $ivalue = _ad_moveobject($sou2, $sobject)
		If $ivalue = 1 Then
			MsgBox(64, "Active Directory ", $idrh1 & "' déplacé vers  '" & $ou & "'")
			$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & " déplacé vers " & $ou
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF &  $idrh1 & " déplacé vers " & $ou & @CRLF);, 1)
			_ad_addusertogroup("RG-" & $ou & "_SM_Siege Banque_LBPAI" , $idrh1) ; ex: $ou=MINA
			_ad_addusertogroup("RG-PITR_Specifique_SurfSSO_Prod", $idrh1) ;surf SSO
			ToolTip("synchronising AD", 5, 5)
			Sleep(3000)
			ToolTip("", 5, 5)
		ElseIf @error = 1 Then
			MsgBox(64, "Active Directory ", "OU cible  '" & $ou & "' n'existe pas !")
			Return 0
		ElseIf @error = 2 Then
			MsgBox(64, "Active Directory ", "Object '" & $idrh1 & "' n'existe pas !")
			Return 0
		ElseIf @error = 3 Then
			MsgBox(64, "Active Directory ", "Object '" & $idrh1 & "' déjà dans l'OU cible  '" & $ou & "'")
		ElseIf @error >= 4 OR @error = -2147352567 Then
			MsgBox(64, "Active Directory ", "You have no permissions to move user " & $idrh1)
			$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & "non deplacé vers " & $ou & ",  accès refusé !"
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF &  $idrh1 & "non deplacé vers " & $ou & ",  accès refusé !" & @CRLF);, 1)
			Return 0
		Else
			MsgBox(0, "warning !", "User Idrh Srce '" & $idrh1 & "existe déjà dans l'AD !", 7)
			Return 0
		EndIf
		$lbpaiapplied = 0
		If $lbpaiapplied = 0 Then
			$lbpaiapplied = 1
			Local $groupscategory = ""
			Global $groupslbpai = "Chargé de Développement|Direction des Ressources Humaines|Direction adminstrative et financière|Direction Technique|Direction juridique et risque|Direction des systèmes d'information|Direction Marketing et Commerciale|Standard|commun full users|Direction générale|BDD_DAF|Bibliotheque_DP|Comité de Direction|Communication Interne|Contrats|DFRC|Direction Financiere|Direction Systèmes Informatiques|DRA|Echange_DAF-DT|Echange_DM-DT|Echange_DRA-DT|Gourvernance|Risques et Conformité|Services Generaux|Supports de présentation"
			Global $groupsidrh1 = StringSplit($groupslbpai, "|")
			Global $aselected[1]
			If $remove = 0 Then
				Global $hgui2 = GUICreate("Selection LBPAI/Iliade Categories pour [+] à user Srce '" & $idrh1 & "' , puis valider avec le bouton [ Apply ] ! (ou refermer la fenetre [x], pour abandonner)...", 1000, 500)
			Else
				Global $hgui2 = GUICreate("Selection LBPAI/Iliade Categories pour [-] à user Srce '" & $idrh1 & "' , puis valider avec le bouton [ Apply ] ! (ou refermer la fenetre [x], pour abandonner)...", 1000, 500)
			EndIf
			Global $hlistbox = _guictrllistbox_create($hgui2, "String upon creation", 10, 30, 980, 470, BitOR($ws_border, $lbs_hasstrings, $ws_vscroll, $lbs_multiplesel))
			_guictrllistbox_beginupdate($hlistbox)
			_guictrllistbox_resetcontent($hlistbox)
			_guictrllistbox_initstorage($hlistbox, 100, 4096)
			_guictrllistbox_endupdate($hlistbox)
			If IsArray($groupsidrh1) = 1 Then
				For $z = 1 To $groupsidrh1[0]
					_guictrllistbox_addstring($hlistbox, $groupsidrh1[$z])
				Next
			EndIf
			Global $hbutton2 = GUICtrlCreateButton("[ Apply ]", 20, 3, 80, 25)
			GUISetState()
			While 1
				Switch GUIGetMsg()
					Case $hbutton2
						$aselected = _guictrllistbox_getselitems($hlistbox)
						If $aselected[0] = 1 Then
							$sitem = " item"
						Else
							$sitem = " items"
						EndIf
						$sitems = ""
						For $i = 1 To $aselected[0]
							$sitems &= _guictrllistbox_gettext($hlistbox, $aselected[$i]) & @CRLF
						Next
						If $aselected[0] <> 0 Then
							$sitems = StringSplit($sitems, @CRLF)
							For $z = 1 To $sitems[0]
								If StringLen($sitems[$z]) <> 0 Then
									$groupscategory = ""
									Select
										Case $sitems[$z] = "Chargé de Développement"
											If $remove = 0 Then
												_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Marketing-Commerciale", $idrh1)
												_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Marketing-Commerciale_Commun-Mark-Com", $idrh1)
											Else
												_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Marketing-Commerciale", $idrh1)
												_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Marketing-Commerciale_Commun-Mark-Com", $idrh1)
											EndIf
											$groupscategory = "( SG-MIPS_ILIADE_Direction-Marketing-Commerciale;SG-MIPS_ILIADE_Direction-Marketing-Commerciale_Commun-Mark-Com )"
										Case $sitems[$z] = "Direction des Ressources Humaines"
											If $remove = 0 Then
												_ad_addusertogroup("SG-MIPS_ILIADE_Communication-Interne", $idrh1)
												_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Ressources-Humaines", $idrh1)
											Else
												_ad_removeuserfromgroup("SG-MIPS_ILIADE_Communication-Interne", $idrh1)
												_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Ressources-Humaines", $idrh1)
											EndIf
											$groupscategory = "( SG-MIPS_ILIADE_Communication-Interne;SG-MIPS_ILIADE_Direction-Ressources-Humaines)"
										Case $sitems[$z] = "Direction adminstrative et financière"
											If $remove = 0 Then
												_ad_addusertogroup("SG-MIPS_ILIADE_Echange_DAF-DT", $idrh1)
												_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Financiere", $idrh1)
											Else
												_ad_removeuserfromgroup("SG-MIPS_ILIADE_Echange_DAF-DT", $idrh1)
												_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Financiere", $idrh1)
											EndIf
											$groupscategory = "(SG-MIPS_ILIADE_Echange_DAF-DT;SG-MIPS_ILIADE_Direction-Financiere)"
										Case $sitems[$z] = "Direction Technique"
											If $remove = 0 Then
												_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Technique", $idrh1)
											Else
												_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Technique", $idrh1)
											EndIf
											$groupscategory = "(SG-MIPS_ILIADE_Direction-Technique)"
										Case $sitems[$z] = "Direction juridique et risque"
											If $remove = 0 Then
												_ad_addusertogroup("SG-MIPS_ILIADE_Risques-&-Conformité", $idrh1)
											Else
												_ad_removeuserfromgroup("SG-MIPS_ILIADE_Risques-&-Conformité", $idrh1)
											EndIf
											$groupscategory = "(SG-MIPS_ILIADE_Risques-&-Conformité)"
										Case $sitems[$z] = "Direction des systèmes d'information"
											If $remove = 0 Then
												_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Systèmes-Informatiques", $idrh1)
											Else
												_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Systèmes-Informatiques", $idrh1)
											EndIf
											$groupscategory = "(SG-MIPS_ILIADE_Direction-Systèmes-Informatiques)"
										Case $sitems[$z] = "Direction Marketing et Commerciale"
											If $remove = 0 Then
												_ad_addusertogroup("SG-MIPS_ILIADE_Echange_DM-DT", $idrh1)
												_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Marketing-Commerciale", $idrh1)
												_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Marketing-Commerciale_Commun-Mark-Com", $idrh1)
											Else
												_ad_removeuserfromgroup("SG-MIPS_ILIADE_Echange_DM-DT", $idrh1)
												_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Marketing-Commerciale", $idrh1)
												_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Marketing-Commerciale_Commun-Mark-Com", $idrh1)
											EndIf
											$groupscategory = "(SG-MIPS_ILIADE_Echange_DM-DT;SG-MIPS_ILIADE_Direction-Marketing-Commerciale;SG-MIPS_ILIADE_Direction-Marketing-Commerciale_Commun-Mark-Com)"
										Case $sitems[$z] = "Standard"
											If $remove = 0 Then
												_ad_addusertogroup("SG-MIPS_ILIADE", $idrh1)
											Else
												_ad_removeuserfromgroup("SG-MIPS_ILIADE", $idrh1)
											EndIf
											$groupscategory = "(SG-MIPS_ILIADE)"
										Case $sitems[$z] = "commun full users"
											If $remove = 0 Then
												_ad_addusertogroup("RG-MIPS_SM_Siege Banque_LBPAI", $idrh1)
												_ad_addusertogroup("RG-MIPS_AW COMMUN MANAGER", $idrh1)
												_ad_addusertogroup("RG-PITR_Specifique_SurfSSO_Prod", $idrh1)
												_ad_addusertogroup("RG-MIPS_AP_Utilisateurs Konica", $idrh1)
											Else
												_ad_removeuserfromgroup("RG-MIPS_SM_Siege Banque_LBPAI", $idrh1)
												_ad_removeuserfromgroup("RG-MIPS_AW COMMUN MANAGER", $idrh1)
												_ad_removeuserfromgroup("RG-PITR_Specifique_SurfSSO_Prod", $idrh1)
												_ad_removeuserfromgroup("RG-MIPS_AP_Utilisateurs Konica", $idrh1)
											EndIf
											$groupscategory = "(RG-MIPS_SM_Siege Banque_LBPAI;RG-MIPS_AW COMMUN MANAGER;RG-PITR_Specifique_SurfSSO_Prod;RG-MIPS_AP_Utilisateurs Konica)"
										Case $sitems[$z] = "Direction générale"
											If $remove = 0 Then
												_ad_addusertogroup("SG-MIPS_ILIADE_Comité-Direction", $idrh1)
												_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Générale", $idrh1)
												_ad_addusertogroup("SG-MIPS_ILIADE_Supports-Présentation", $idrh1)
											Else
												_ad_removeuserfromgroup("SG-MIPS_ILIADE_Comité-Direction", $idrh1)
												_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Générale", $idrh1)
												_ad_removeuserfromgroup("SG-MIPS_ILIADE_Supports-Présentation", $idrh1)
											EndIf
											$groupscategory = "(SG-MIPS_ILIADE_Comité-Direction;SG-MIPS_ILIADE_Direction-Générale;SG-MIPS_ILIADE_Supports-Présentation)"
										Case $sitems[$z] = "BDD_DAF" AND $remove = 0
											_ad_addusertogroup("SG-MIPS_ILIADE_BDD_DAF", $idrh1)
											$groupscategory = "(SG-MIPS_ILIADE_BDD_DAF)"
										Case $sitems[$z] = "Bibliotheque_DP" AND $remove = 0
											_ad_addusertogroup("SG-MIPS_ILIADE_Bibliotheque_D", $idrh1)
											$groupscategory = "(SG-MIPS_ILIADE_Bibliotheque_D)"
										Case $sitems[$z] = "Comité de Direction" AND $remove = 0
											_ad_addusertogroup("SG-MIPS_ILIADE_Comité-Direction", $idrh1)
											$groupscategory = "(SG-MIPS_ILIADE_Comité-Direction)"
										Case $sitems[$z] = "Communication Interne" AND $remove = 0
											_ad_addusertogroup("SG-MIPS_ILIADE_Communication-Interne", $idrh1)
											$groupscategory = "(SG-MIPS_ILIADE_Communication-Interne)"
										Case $sitems[$z] = "Contrats" AND $remove = 0
											_ad_addusertogroup("SG-MIPS_ILIADE_Contrats_RW", $idrh1)
											$groupscategory = "(SG-MIPS_ILIADE_Contrats_RW)"
										Case $sitems[$z] = "DFRC" AND $remove = 0
											_ad_addusertogroup("SG-MIPS_ILIADE_DFRC", $idrh1)
											$groupscategory = "(SG-MIPS_ILIADE_DFRC)"
										Case $sitems[$z] = "Direction Financiere" AND $remove = 0
											_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Financiere", $idrh1)
											$groupscategory = "(SG-MIPS_ILIADE_Direction-Financiere)"
										Case $sitems[$z] = "Direction Systèmes Informatiques" AND $remove = 0
											_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Systèmes-Informatiques", $idrh1)
											$groupscategory = "(SG-MIPS_ILIADE_Direction-Systèmes-Informatiques)"
										Case $sitems[$z] = "DRA" AND $remove = 0
											_ad_addusertogroup("SG-MIPS_ILIADE_DRA", $idrh1)
											$groupscategory = "(SG-MIPS_ILIADE_DRA)"
										Case $sitems[$z] = "Echange_DAF-DT" AND $remove = 0
											_ad_addusertogroup("SG-MIPS_ILIADE_Echange_DAF-DT", $idrh1)
											$groupscategory = "(SG-MIPS_ILIADE_Echange_DAF-DT)"
										Case $sitems[$z] = "Echange_DM-DT" AND $remove = 0
											_ad_addusertogroup("SG-MIPS_ILIADE_Echange_DM-DT", $idrh1)
											$groupscategory = "(SG-MIPS_ILIADE_Echange_DM-DT)"
										Case $sitems[$z] = "Echange_DRA-DT" AND $remove = 0
											_ad_addusertogroup("SG-MIPS_ILIADE_Echange_DRA-DT", $idrh1)
											$groupscategory = "(SG-MIPS_ILIADE_Echange_DRA-DT)"
										Case $sitems[$z] = "Gourvernance" AND $remove = 0
											_ad_addusertogroup("SG-MIPS_ILIADE_Gouvernance", $idrh1)
											$groupscategory = "(SG-MIPS_ILIADE_Gouvernance)"
										Case $sitems[$z] = "Risques et Conformité" AND $remove = 0
											_ad_addusertogroup("SG-MIPS_ILIADE_Risques-&-Conformité", $idrh1)
											$groupscategory = "(SG-MIPS_ILIADE_Risques-&-Conformité)"
										Case $sitems[$z] = "Services Generaux" AND $remove = 0
											_ad_addusertogroup("SG-MIPS_ILIADE_Services_generaux", $idrh1)
											$groupscategory = "(SG-MIPS_ILIADE_Services_generaux)"
										Case $sitems[$z] = "Supports de présentation" AND $remove = 0
											_ad_addusertogroup("SG-MIPS_ILIADE_Supports-Présentation", $idrh1)
											$groupscategory = "(SG-MIPS_ILIADE_Supports-Présentation)"
									EndSelect
									$actiongroups = $actiongroups & " [ " & $sitems[$z] & " ] " & " ; " & $groupscategory & @CRLF
								EndIf
							 Next

						    Global $groupsidrh1 = _ad_getusergroups($idrh1)
						If @error > 0 Then
						Else
							_arraysort($groupsidrh1, 0, 1)
							_arraydelete($groupsidrh1, 0)
							Global $groupidrh_final = ""
							Global $groupidrh1_add = $groupsidrh1
							For $k = 0 To UBound($groupsidrh1) - 1
								$groupidrh = $groupsidrh1[$k]
								$groupidrh = StringSplit($groupidrh, ",")
								$groupidrh_int = $groupidrh[1]
								$groupidrh_int = StringReplace($groupidrh_int, "CN=", "")
								$groupidrh_final = $groupidrh_final & $groupidrh_int & @CRLF
							Next
							$groupsidrh1 = _arraytostring($groupsidrh1)
						 EndIf

							If $remove = 0 Then
								$actiongroups = " [ Categories de groupes LBPAI/Iliade ]  :  [+]  > " & $idrh1 & @CRLF & $actiongroups
								_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF &  $actiongroups & @CRLF);, 1)
								_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  >Groupes user:  " & $idrh1 & @CRLF & $groupidrh_final & @CRLF);, 1)
							Else
								$actiongroups = " [ Categories de groupes LBPAI/Iliade ]  :  [-]  > " & $idrh1 & @CRLF & $actiongroups
								_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF &  $actiongroups & @CRLF);, 1)
								_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  >Groupes user:  " & $idrh1 & @CRLF & $groupidrh_final & @CRLF);, 1)
							EndIf
							GUIDelete($hgui2)
							ToolTip("Synchronizing with AD...", 5, 5)
							Sleep(3000)
							ToolTip("", 5, 5)
							ExitLoop
						Else
							MsgBox(4160, "warning !", "aucune selection... réessayez ou refermer la fenetre pour abandonner... [x]")
						EndIf
					Case $gui_event_close
						GUIDelete($hgui2)
						ExitLoop
				EndSwitch
			WEnd
			GUIDelete($hgui2)
		EndIf
	EndFunc

	Func massive_move()
		If $isdct = 0 Then
			MsgBox(0, "Info !", "massive move Valable que pour le domaine DCT... ", 14)
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "massive move Valable que pour le domaine DCT... " & @CRLF);, 1)
			Return 0
		EndIf
		Global $dcdomain = "DC=dct,DC=adt,DC=local"
		Global $racinead = "DCT"
		Dim $table
		$defautdc = $defautdcinit
		Local Const $smessage = "fichier cible: 'Liste users à deplacer" & "*" & ".txt'  à traiter !  (Utilisateurs à déplacer, format ligne: Pabc123|BPRE )"
		Local $sfileopendialog = FileOpenDialog($smessage, @ScriptDir & "\", "Txt (*.txt)", $fd_filemustexist, "*" & ".txt")
		If @error Then
			MsgBox(0, "", "Aucun fichier choisi !")
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Aucun fichier choisi !" & @CRLF);, 1)
			Return 0
		Else
			$sfileopendialog = StringReplace($sfileopendialog, "|", @CRLF)
			MsgBox(0, "", "fichier choisi : " & $sfileopendialog)
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "fichier choisi : " & $sfileopendialog & @CRLF);, 1)
		EndIf
		$f = $sfileopendialog
		$f = StringSplit($f, "\")
		$i = $f[0]
		$file = FileOpen($f[$i], 0)
		Global $groupfile = $f[$i]
		If $file = -1 Then
			MsgBox(0, "Error", "Unable to open file." & $f[$i])
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Unable to open file." & $f[$i] & @CRLF);, 1)
			Return 0
		EndIf
		$file = "" & $f[$i]
		$sizeof = FileGetSize($file)
		If $sizeof = 0 Then
			MsgBox(0, $group & ".txt", "no users to move in file: " & $group & @CRLF & "aborting !")
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "no users to move in file: " & $group & @CRLF & "aborting !" & @CRLF);, 1)
			Return 0
		EndIf
		_filereadtoarray($file, $table, 1)
		$taille = $table[0]
		_arraydelete($table, 0)
		$out = FileOpen(@ScriptDir & "\Log_Users_déplacés.txt", 1)
		If FileGetSize(@ScriptDir & "\Log_Users_déplacés.txt") = 0 Then
			FileWriteLine($out, " IDRH" & "  |" & "		OU d'Origine" & "						 |" & "		OU Destination" & "			      |  Resultat" & @CRLF)
		EndIf
		FileWriteLine($out, @CRLF & "*** Start of Move Users from " & $racinead & " ***  " & _now() & @CRLF & @CRLF)
		For $i = 0 To $taille - 1
			$ligne = $table[$i]
			$ligne = StringSplit($ligne, "|")
			$r = StringReplace($table[$i], "|", "|")
			$s = @extended
			If $s = 1 OR $s = 2 OR $s = 3 Then
				$user = $ligne[1]
				$ou = $ligne[2]
				If $ou <> "" Then
					$stargetou = "OU=Utilisateurs,OU=" & $ou & ",OU=" & StringLeft($ou, 2) & "," & $dcdomain
					$sobject = _ad_samaccountnametofqdn($user)
					ToolTip("User move:  " & $sobject & "  |  OU cible =" & $stargetou, 5, 5)
					Global $ivalue = _ad_moveobject($stargetou, $sobject)
					If $ivalue = 1 Then
						FileWriteLine($out, $user & "|" & $sobject & "|" & $stargetou & "|  		déplacement OK" & @CRLF)
					ElseIf @error = 1 Then
						FileWriteLine($out, $user & "|" & $sobject & "|" & $stargetou & "|  		l'OU cible n'existe pas !" & @CRLF)
					ElseIf @error = 2 Then
						FileWriteLine($out, $user & "|" & $sobject & "									 |" & $stargetou & "|  		l'Utilisateur n'existe pas sur l'AD-" & $racinead & " !" & @CRLF)
					ElseIf @error = 3 Then
						FileWriteLine($out, $user & "|" & $sobject & "|" & $stargetou & "|  		Utilisateur déjà dans l'OU Cible !" & @CRLF)
					ElseIf @error >= 4 OR @error = -2147352567 Then
						FileWriteLine($out, $user & "|" & $sobject & "|" & $stargetou & "|  		accès refusé !" & @CRLF)
					Else
					EndIf
				Else
					FileWriteLine($out, $user & "|" & $sobject & "|" & $stargetou & "|  		OU cible vide !" & @CRLF & @CRLF)
				EndIf
				$remove = 0
				If $ligne[0] > 2 Then
					$sitems = StringSplit($ligne[3], ";")
					$groupscategory = ""
					$actiongroups = ""
					For $z = 1 To $sitems[0]
						If StringLen($sitems[$z]) <> 0 Then
							If $ligne[0] = 4 Then
								$remove4 = $ligne[4]
								If $remove4 = 0 Then
									$remove = 0
								EndIf
								If $remove4 = 1 Then
									$remove = 1
								EndIf
							EndIf
							Select
								Case $sitems[$z] = "Chargé de Développement"
									If $remove = 0 Then
										_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Marketing-Commerciale", $idrh1)
										_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Marketing-Commerciale_Commun-Mark-Com", $idrh1)
									Else
										_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Marketing-Commerciale", $idrh1)
										_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Marketing-Commerciale_Commun-Mark-Com", $idrh1)
									EndIf
									$groupscategory = "( SG-MIPS_ILIADE_Direction-Marketing-Commerciale;SG-MIPS_ILIADE_Direction-Marketing-Commerciale_Commun-Mark-Com )"
								Case $sitems[$z] = "Direction des Ressources Humaines"
									If $remove = 0 Then
										_ad_addusertogroup("SG-MIPS_ILIADE_Communication-Interne", $idrh1)
										_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Ressources-Humaines", $idrh1)
									Else
										_ad_removeuserfromgroup("SG-MIPS_ILIADE_Communication-Interne", $idrh1)
										_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Ressources-Humaines", $idrh1)
									EndIf
									$groupscategory = "( SG-MIPS_ILIADE_Communication-Interne;SG-MIPS_ILIADE_Direction-Ressources-Humaines)"
								Case $sitems[$z] = "Direction adminstrative et financière"
									If $remove = 0 Then
										_ad_addusertogroup("SG-MIPS_ILIADE_Echange_DAF-DT", $idrh1)
										_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Financiere", $idrh1)
									Else
										_ad_removeuserfromgroup("SG-MIPS_ILIADE_Echange_DAF-DT", $idrh1)
										_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Financiere", $idrh1)
									EndIf
									$groupscategory = "(SG-MIPS_ILIADE_Echange_DAF-DT;SG-MIPS_ILIADE_Direction-Financiere)"
								Case $sitems[$z] = "Direction Technique"
									If $remove = 0 Then
										_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Technique", $idrh1)
									Else
										_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Technique", $idrh1)
									EndIf
									$groupscategory = "(SG-MIPS_ILIADE_Direction-Technique)"
								Case $sitems[$z] = "Direction juridique et risque"
									If $remove = 0 Then
										_ad_addusertogroup("SG-MIPS_ILIADE_Risques-&-Conformité", $idrh1)
									Else
										_ad_removeuserfromgroup("SG-MIPS_ILIADE_Risques-&-Conformité", $idrh1)
									EndIf
									$groupscategory = "(SG-MIPS_ILIADE_Risques-&-Conformité)"
								Case $sitems[$z] = "Direction des systèmes d'information"
									If $remove = 0 Then
										_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Systèmes-Informatiques", $idrh1)
									Else
										_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Systèmes-Informatiques", $idrh1)
									EndIf
									$groupscategory = "(SG-MIPS_ILIADE_Direction-Systèmes-Informatiques)"
								Case $sitems[$z] = "Direction Marketing et Commerciale"
									If $remove = 0 Then
										_ad_addusertogroup("SG-MIPS_ILIADE_Echange_DM-DT", $idrh1)
										_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Marketing-Commerciale", $idrh1)
										_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Marketing-Commerciale_Commun-Mark-Com", $idrh1)
									Else
										_ad_removeuserfromgroup("SG-MIPS_ILIADE_Echange_DM-DT", $idrh1)
										_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Marketing-Commerciale", $idrh1)
										_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Marketing-Commerciale_Commun-Mark-Com", $idrh1)
									EndIf
									$groupscategory = "(SG-MIPS_ILIADE_Echange_DM-DT;SG-MIPS_ILIADE_Direction-Marketing-Commerciale;SG-MIPS_ILIADE_Direction-Marketing-Commerciale_Commun-Mark-Com)"
								Case $sitems[$z] = "Standard"
									If $remove = 0 Then
										_ad_addusertogroup("SG-MIPS_ILIADE", $idrh1)
									Else
										_ad_removeuserfromgroup("SG-MIPS_ILIADE", $idrh1)
									EndIf
									$groupscategory = "(SG-MIPS_ILIADE)"
								Case $sitems[$z] = "commun full users"
									If $remove = 0 Then
										_ad_addusertogroup("RG-MIPS_SM_Siege Banque_LBPAI", $idrh1)
										_ad_addusertogroup("RG-MIPS_AW COMMUN MANAGER", $idrh1)
										_ad_addusertogroup("RG-PITR_Specifique_SurfSSO_Prod", $idrh1)
										_ad_addusertogroup("RG-MIPS_AP_Utilisateurs Konica", $idrh1)
									Else
										_ad_removeuserfromgroup("RG-MIPS_SM_Siege Banque_LBPAI", $idrh1)
										_ad_removeuserfromgroup("RG-MIPS_AW COMMUN MANAGER", $idrh1)
										_ad_removeuserfromgroup("RG-PITR_Specifique_SurfSSO_Prod", $idrh1)
										_ad_removeuserfromgroup("RG-MIPS_AP_Utilisateurs Konica", $idrh1)
									EndIf
									$groupscategory = "(RG-MIPS_SM_Siege Banque_LBPAI;RG-MIPS_AW COMMUN MANAGER;RG-PITR_Specifique_SurfSSO_Prod;RG-MIPS_AP_Utilisateurs Konica)"
								Case $sitems[$z] = "Direction générale"
									If $remove = 0 Then
										_ad_addusertogroup("SG-MIPS_ILIADE_Comité-Direction", $idrh1)
										_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Générale", $idrh1)
										_ad_addusertogroup("SG-MIPS_ILIADE_Supports-Présentation", $idrh1)
									Else
										_ad_removeuserfromgroup("SG-MIPS_ILIADE_Comité-Direction", $idrh1)
										_ad_removeuserfromgroup("SG-MIPS_ILIADE_Direction-Générale", $idrh1)
										_ad_removeuserfromgroup("SG-MIPS_ILIADE_Supports-Présentation", $idrh1)
									EndIf
									$groupscategory = "(SG-MIPS_ILIADE_Comité-Direction;SG-MIPS_ILIADE_Direction-Générale;SG-MIPS_ILIADE_Supports-Présentation)"
								Case $sitems[$z] = "BDD_DAF" AND $remove = 0
									_ad_addusertogroup("SG-MIPS_ILIADE_BDD_DAF", $idrh1)
									$groupscategory = "(SG-MIPS_ILIADE_BDD_DAF)"
								Case $sitems[$z] = "Bibliotheque_DP" AND $remove = 0
									_ad_addusertogroup("SG-MIPS_ILIADE_Bibliotheque_D", $idrh1)
									$groupscategory = "(SG-MIPS_ILIADE_Bibliotheque_D)"
								Case $sitems[$z] = "Comité de Direction" AND $remove = 0
									_ad_addusertogroup("SG-MIPS_ILIADE_Comité-Direction", $idrh1)
									$groupscategory = "(SG-MIPS_ILIADE_Comité-Direction)"
								Case $sitems[$z] = "Communication Interne" AND $remove = 0
									_ad_addusertogroup("SG-MIPS_ILIADE_Communication-Interne", $idrh1)
									$groupscategory = "(SG-MIPS_ILIADE_Communication-Interne)"
								Case $sitems[$z] = "Contrats" AND $remove = 0
									_ad_addusertogroup("SG-MIPS_ILIADE_Contrats_RW", $idrh1)
									$groupscategory = "(SG-MIPS_ILIADE_Contrats_RW)"
								Case $sitems[$z] = "DFRC" AND $remove = 0
									_ad_addusertogroup("SG-MIPS_ILIADE_DFRC", $idrh1)
									$groupscategory = "(SG-MIPS_ILIADE_DFRC)"
								Case $sitems[$z] = "Direction Financiere" AND $remove = 0
									_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Financiere", $idrh1)
									$groupscategory = "(SG-MIPS_ILIADE_Direction-Financiere)"
								Case $sitems[$z] = "Direction Systèmes Informatiques" AND $remove = 0
									_ad_addusertogroup("SG-MIPS_ILIADE_Direction-Systèmes-Informatiques", $idrh1)
									$groupscategory = "(SG-MIPS_ILIADE_Direction-Systèmes-Informatiques)"
								Case $sitems[$z] = "DRA" AND $remove = 0
									_ad_addusertogroup("SG-MIPS_ILIADE_DRA", $idrh1)
									$groupscategory = "(SG-MIPS_ILIADE_DRA)"
								Case $sitems[$z] = "Echange_DAF-DT" AND $remove = 0
									_ad_addusertogroup("SG-MIPS_ILIADE_Echange_DAF-DT", $idrh1)
									$groupscategory = "(SG-MIPS_ILIADE_Echange_DAF-DT)"
								Case $sitems[$z] = "Echange_DM-DT" AND $remove = 0
									_ad_addusertogroup("SG-MIPS_ILIADE_Echange_DM-DT", $idrh1)
									$groupscategory = "(SG-MIPS_ILIADE_Echange_DM-DT)"
								Case $sitems[$z] = "Echange_DRA-DT" AND $remove = 0
									_ad_addusertogroup("SG-MIPS_ILIADE_Echange_DRA-DT", $idrh1)
									$groupscategory = "(SG-MIPS_ILIADE_Echange_DRA-DT)"
								Case $sitems[$z] = "Gourvernance" AND $remove = 0
									_ad_addusertogroup("SG-MIPS_ILIADE_Gouvernance", $idrh1)
									$groupscategory = "(SG-MIPS_ILIADE_Gouvernance)"
								Case $sitems[$z] = "Risques et Conformité" AND $remove = 0
									_ad_addusertogroup("SG-MIPS_ILIADE_Risques-&-Conformité", $idrh1)
									$groupscategory = "(SG-MIPS_ILIADE_Risques-&-Conformité)"
								Case $sitems[$z] = "Services Generaux" AND $remove = 0
									_ad_addusertogroup("SG-MIPS_ILIADE_Services_generaux", $idrh1)
									$groupscategory = "(SG-MIPS_ILIADE_Services_generaux)"
								Case $sitems[$z] = "Supports de présentation" AND $remove = 0
									_ad_addusertogroup("SG-MIPS_ILIADE_Supports-Présentation", $idrh1)
									$groupscategory = "(SG-MIPS_ILIADE_Supports-Présentation)"
							EndSelect
							$actiongroups = $actiongroups & " [ " & $sitems[$z] & " ] " & " ; " & $groupscategory & @CRLF
						EndIf
					Next
					If $remove = 0 Then
						FileWriteLine($out, $user & " | LBPAI/Iliade Categories [+]" & @CRLF & $actiongroups & @CRLF & @CRLF)
					EndIf
					If $remove = 1 Then
						FileWriteLine($out, $user & " | LBPAI/Iliade Categories [-]" & @CRLF & $actiongroups & @CRLF & @CRLF)
					EndIf
				EndIf
			Else
				FileWriteLine($out, $user & " | line: " & $i & "  wrong format !  " & $table[$i] & @CRLF & @CRLF)
			EndIf
		Next
		ToolTip("", 5, 5)
		FileWriteLine($out, @CRLF & "*** End of Move Users from " & $racinead & " ***  " & _now() & @CRLF & @CRLF)
		FileClose($file)
		FileClose($out)
		MsgBox(0, "Info !", "Consultez fichier log" & @CRLF & "Log_Users_déplacés.txt")
		_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Consultez fichier log:  " & "Log_Users_déplacés.txt" & @CRLF);, 1)
		Return 0
	EndFunc

	Func massive_date()
		$defautdc = $defautdcinit
		Dim $table
		Local Const $smessage = "fichier cible: 'Liste users à changer la date d'expiration" & "*" & ".txt'  à traiter !  (Utilisateurs à traiter, format ligne: Pabc123|jj/mm/aaaa )"
		Local $sfileopendialog = FileOpenDialog($smessage, @ScriptDir & "\", "Txt (*.txt)", $fd_filemustexist, "*" & ".txt")
		If @error Then
			MsgBox(0, "", "Aucun fichier choisi !")
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Aucun fichier choisi !" & @CRLF);, 1)
			Return 0
		Else
			$sfileopendialog = StringReplace($sfileopendialog, "|", @CRLF)
			MsgBox(0, "", "fichier choisi : " & $sfileopendialog)
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "fichier choisi : " & $sfileopendialog & @CRLF);, 1)
		EndIf
		$f = $sfileopendialog
		$f = StringSplit($f, "\")
		$i = $f[0]
		$file = FileOpen($f[$i], 0)
		Global $groupfile = $f[$i]
		If $file = -1 Then
			MsgBox(0, "Error", "Unable to open file." & $f[$i])
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Unable to open file." & $f[$i] & @CRLF);, 1)
			Return 0
		EndIf
		$file = "" & $f[$i]
		$sizeof = FileGetSize($file)
		If $sizeof = 0 Then
			MsgBox(0, $group & ".txt", "no users to extend Date in file: " & $group & @CRLF & "aborting !")
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "no users to move in file: " & $group & @CRLF & "aborting !" & @CRLF);, 1)
			Return 0
		EndIf
		_filereadtoarray($file, $table, 1)
		$taille = $table[0]
		_arraydelete($table, 0)
		$out = FileOpen(@ScriptDir & "\Log_Users_date prolongée.txt", 1)
		If FileGetSize(@ScriptDir & "\Log_Users_date prolongée.txt") = 0 Then
		EndIf
		FileWriteLine($out, @CRLF & "*** Start of Users Date Extend  " & $racinead & " ***  " & _now() & @CRLF & @CRLF)
		For $i = 0 To $taille - 1
			$idrh1 = ""
			$ligne = $table[$i]
			If NOT StringInStr($ligne, "|") Then
				MsgBox(0, "Warning !", "Bad separator in txt file... missing separator: | line: " & $i + 1)
				FileWriteLine($out, "Warning !  " & ",  Bad separator in txt file... missing separator: '|'  in line: " & $i + 1 & @CRLF)
				_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Warning !  " & ",  Bad separator in txt file... missing separator: '|'  in line: " & $i + 1  & @CRLF);, 1)
				Return 0
			EndIf
			$ligne = StringSplit($ligne, "|")
			If IsArray($ligne) Then
				$user = $ligne[1]
				$idrh1 = $user
				$dateprolonge = $ligne[2]
				If $dateprolonge = "0" Then
					$dateprolonge = "00/00/0000"
				EndIf
				If NOT StringRegExp($dateprolonge, "(\d{2})/(\d{2})/(\d{4})") Then
					MsgBox(0, "Warning !", "Bad format Date in txt file... line: " & $i + 1)
					FileWriteLine($out, "Warning !  " & ",  Bad format Date in txt file... line: " & $i + 1 & @CRLF)
					_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Warning !  " & ",  Bad format Date in txt file... line: " & $i + 1  & @CRLF);, 1)
					Return 0
				EndIf
			Else
				MsgBox(0, "warning !", "bad format txt file..." & @CRLF & "abort...")
				_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "warning !", "bad format txt file..." & @CRLF & "abort..." & @CRLF);, 1)
				Return 0
			EndIf
			$userexist = _ad_objectexists($idrh1)
			If $userexist = 1 Then
				Global $date_expire = ""
				$date_expire = _ad_getobjectproperties($idrh1, "accountExpires")
				$date_expire = $date_expire[1][1]
				$date_expire = _dateadd("d", -1, $date_expire)

				$date_prolonger = $dateprolonge
				$joursd = StringSplit($dateprolonge, "/")
				Global $jourd = $joursd[1]
				Global $moisd = $joursd[2]
				Global $anneed = $joursd[3]
				$dated = $anneed & "-" & $moisd & "-" & $jourd & " 00:00:00"
				$dated = _dateadd("d", 1, $dated)
				If $dated = "00/00/0000 00:00:00" Then
					$dated = "01/01/1970 00:00:00"
					$ivalue = _ad_setaccountexpire($idrh1, $dated)
				Else
					$ivalue = _ad_setaccountexpire($idrh1, $dated)
				EndIf
				If $ivalue = 1 Then
					If $dated = "01/01/1970 00:00:00" Then
						FileWriteLine($out, $idrh1 & "  |" & "		Date d'Origine: " & $date_expire & "		|" & "		compte n'expire plus !	" & "  |  Resultat: OK" & @CRLF)
					Else
						If $date_expire = 0 Then
							FileWriteLine($out, $idrh1 & "  |" & "		Date d'Origine: " & $date_expire & "				|" & "		Date prolongée:	" & $date_prolonger & "  |  Resultat: OK" & @CRLF)
						Else
							FileWriteLine($out, $idrh1 & "  |" & "		Date d'Origine: " & $date_expire & "		|" & "		Date prolongée:	" & $date_prolonger & "  |  Resultat: OK" & @CRLF)
						EndIf
					EndIf
				Else
					FileWriteLine($out, $idrh1 & "  |" & "		Date d'Origine: " & $date_expire & "		|" & "		Date prolongée:	" & $date_prolonger & "  |  Resultat: KO" & @CRLF)
				EndIf
			Else
				FileWriteLine($out, $idrh1 & "  |" & "		Date d'Origine: " & $date_expire & "		|" & "		Date prolongée:	" & $date_prolonger & "  |  Resultat: compte AD n'existe pas !" & @CRLF)
			EndIf
		Next
		ToolTip("", 5, 5)
		FileWriteLine($out, @CRLF & "*** End of Extend Date Users " & $racinead & " ***  " & _now() & @CRLF & @CRLF)
		FileClose($file)
		FileClose($out)
		MsgBox(0, "Info !", "Consultez fichier log" & @CRLF & "Log_Users_date prolongée.txt")
		_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Consultez fichier log:  " & "Log_Users_date prolongée.txt" & @CRLF);, 1)
		Return 0
	EndFunc

	Func massive_drive()
		$defautdc = $defautdcinit
		If $isdct = 0 Then
			MsgBox(0, "Warning !", "Valable que pour le domaine DCT... ", 14)
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Valable que pour le domaine DCT... " & @CRLF);, 1)
			Return 0
		EndIf
		Global $outdrive = ""
		$idrh1 = StringSplit($idrh1, ";")
		If IsArray($idrh1) Then
			For $z = 1 To $idrh1[0]
				$userexist = _ad_objectexists($idrh1[$z])
				If $userexist = 1 Then
					Global $drivesidrh1 = _ad_getobjectproperties($idrh1[$z], "LaPoste00-NetworkDrive")
					_arraydelete($drivesidrh1, 0)
					$drivesidrh1 = _arraytostring($drivesidrh1)
					$drivesidrh1 = StringReplace($drivesidrh1, "LaPoste00-NetworkDrive|", "")
					Global $commentidrh1 = _ad_getobjectproperties($idrh1[$z], "Comment")
					_arraydelete($commentidrh1, 0)
					$commentidrh1 = _arraytostring($commentidrh1)
					$commentidrh1 = StringReplace($commentidrh1, "Comment|", "")
					If $drivesidrh1 <> "" Then
						$outdrive = $outdrive & "  user  [ " & $idrh1[$z] & " ]" & "  Drives:" & @CRLF & $drivesidrh1 & @CRLF
					Else
						$outdrive = $outdrive & "  user  [ " & $idrh1[$z] & " ] : no drives !" & @CRLF
					EndIf
					If $commentidrh1 <> "" Then
						$outdrive = $outdrive & "comment:  " & $commentidrh1 & @CRLF & @CRLF
					Else
						$outdrive = $outdrive & "" & @CRLF & @CRLF
					EndIf
				Else
					If $idrh1[$z] <> "" Then
						$outdrive = $outdrive & "  user  [ " & $idrh1[$z] & " ] ? n'existe plus !" & @CRLF & @CRLF
					EndIf
				EndIf
			Next
			$historik = $historik & @CRLF & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "pasted in clipboard..." & @CRLF & @CRLF
			_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF &  "Massive drives; export: pasted in clipboard..." & @CRLF);, 1)
		   ;_GUICtrlRichEdit_AppendText($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $outdrive & @CRLF);, 1)
			ClipPut($outdrive)
			;_GUICtrlRichEdit_AppendText($aff, $historik, 1)
		Else
		EndIf
		Return 0
	EndFunc

	Func treeview_affiche()
   		_ad_open()
	  if $domainname="" Then
		   ToolTip("",5,5,"")
		   Return 0
		EndIf

	  #Region ### START Koda GUI section ### Form=
			If $readad = 0 Then
				Local $readad2 = "first scanning all OUs"
			Else
				Local $readad2 = "already scanned all OUs: GDI in progress..."
			EndIf
			Global $hmain = GUICreate("Active Direcory OU Treeview - [" & $readad2 & "]", 743, 683, -1, -1)
			Global $htree = GUICtrlCreateTreeView(6, 6, 600, 666, -1, $ws_ex_clientedge)
			Global $bexpand = GUICtrlCreateButton("Expand", 624, 56, 97, 33)
			Global $bcollapse = GUICtrlCreateButton("Collapse", 624, 104, 97, 33)
			Global $bselect = GUICtrlCreateButton("Select OU", 624, 152, 97, 33)
		#EndRegion ### END Koda GUI section ###

		Global $atreeview = _ad_getoutreeview($sout, $htree, False, "", "%", 2)
		If @error <> 0 Then MsgBox(16, "Active Direcory Treeview", "Error creating list of OUs starting with '" & $sout & "'." & @CRLF & "Error returned by function _AD_GetALLOUs: @error = " & @error & ", @extended =  " & @extended)
		GUISetState(@SW_SHOW)
	EndFunc

	Func treeselect()
	     if $domainname="" Then
		   ToolTip("",5,5,"")
		   Return 0
		EndIf

		If StringUpper($domainname) = "COURRIER" Then
			ToolTip("OUP/OUP-Z/OUP-Z(nn)", 5, 5, "domaine COURRIER: placez-vous dans l'OU cible OUP-Z(nn)..")
		 EndIf

		While 1
			$msg = GUIGetMsg("$hMain")
			Switch $msg
				Case $bexpand
					_guictrltreeview_expand($htree)
				Case $bcollapse
					_guictrltreeview_expand($htree, 0, False)
				Case $bselect
					$hselection = _guictrltreeview_getselection($htree)
					$sselection = _guictrltreeview_gettext($htree, $hselection)
					For $i = 1 To $atreeview[0][0]
						If $hselection = $atreeview[$i][2] Then ExitLoop
					Next
					$sout = $atreeview[$i][1]
					GUIDelete("$hMain")
					Return $sout
			EndSwitch
		 WEnd

		If StringUpper($domainname) = "COURRIER" Then
			ToolTip("", 5, 5, "")
		EndIf
	EndFunc

	Func _ad_getoutreeview($sad_ou, $had_treeview, $bad_isadopen = False, $sad_category = "", $sad_text = " (%)", $iad_searchscope = 1)
	   if $domainname="" Then
		   MsgBox(0,"Info !","no Active Directory found ! unable to scan...",15)
		   ToolTip("",5,5,"")
		   Return 0
		EndIf

		If $readad = 0 Then
			$readad = 1
			ToolTip("Scanning OUs in Active Directory...", 5, 5)
			Local $iad_count
			If $bad_isadopen = False Then
				_ad_open()
				If @error Then Return SetError(@error, @extended, 0)
			EndIf
			$sseparator = "\"
			Global $aad_ous = _ad_getallous($sad_ou, $sseparator)
			If @error <> 0 Then Return SetError(@error, @extended, 0)
			Global $aad_treeview[$aad_ous[0][0] + 1][3] = [[$aad_ous[0][0], 3]]
			For $i = 1 To $aad_ous[0][0]
				$aad_temp = StringSplit($aad_ous[$i][0], $sseparator)
				$aad_treeview[$i][0] = StringFormat("%" & $aad_temp[0] - 1 & "s", "") & "#" & $aad_temp[$aad_temp[0]]
				$aad_treeview[$i][1] = $aad_ous[$i][1]
			Next
			_guictrltreeview_beginupdate($had_treeview)
			Local $ahad_node[50], $sad_ldapstring
			If StringIsAlNum($sad_category) Then
				$sad_ldapstring = "(objectcategory=" & $sad_category & ")"
			Else
				$sad_ldapstring = $sad_category
			EndIf
			For $iad_index = 1 To $aad_treeview[0][0]
				$sad_line = StringSplit(StringStripCR($aad_treeview[$iad_index][0]), @TAB)
				$iad_level = StringInStr($sad_line[1], "#")
				If $iad_level = 0 Then ExitLoop
				If $sad_category <> "" Then $iad_count = _ad_getobjectsinou($aad_treeview[$iad_index][1], $sad_ldapstring, $iad_searchscope, "samaccountname", "", True)
				If $iad_level = 1 Then
					$sad_temp = ""
					If $sad_category <> "" Then $sad_temp = StringReplace($sad_text, "%", $iad_count)
					$ahad_node[$iad_level] = _guictrltreeview_add($had_treeview, 0, StringMid($sad_line[1], $iad_level + 1) & $sad_temp)
					$aad_treeview[$iad_index][2] = $ahad_node[$iad_level]
				Else
					$sad_temp = ""
					If $sad_category <> "" Then $sad_temp = StringReplace($sad_text, "%", $iad_count)
					$ahad_node[$iad_level] = _guictrltreeview_addchild($had_treeview, $ahad_node[$iad_level - 1], StringMid($sad_line[1], $iad_level + 1) & $sad_temp)
					$aad_treeview[$iad_index][2] = $ahad_node[$iad_level]
				EndIf
			Next
			_guictrltreeview_endupdate($had_treeview)
			ToolTip("", 5, 5)
			Return $aad_treeview
		Else
			$sseparator = "\"
			For $i = 1 To $aad_ous[0][0]
				$aad_temp = StringSplit($aad_ous[$i][0], $sseparator)
				$aad_treeview[$i][0] = StringFormat("%" & $aad_temp[0] - 1 & "s", "") & "#" & $aad_temp[$aad_temp[0]]
				$aad_treeview[$i][1] = $aad_ous[$i][1]
			Next
			_guictrltreeview_beginupdate($had_treeview)
			Local $ahad_node[50], $sad_ldapstring
			If StringIsAlNum($sad_category) Then
				$sad_ldapstring = "(objectcategory=" & $sad_category & ")"
			Else
				$sad_ldapstring = $sad_category
			EndIf
			For $iad_index = 1 To $aad_treeview[0][0]
				$sad_line = StringSplit(StringStripCR($aad_treeview[$iad_index][0]), @TAB)
				$iad_level = StringInStr($sad_line[1], "#")
				If $iad_level = 0 Then ExitLoop
				If $sad_category <> "" Then $iad_count = _ad_getobjectsinou($aad_treeview[$iad_index][1], $sad_ldapstring, $iad_searchscope, "samaccountname", "", True)
				If $iad_level = 1 Then
					$sad_temp = ""
					If $sad_category <> "" Then $sad_temp = StringReplace($sad_text, "%", $iad_count)
					$ahad_node[$iad_level] = _guictrltreeview_add($had_treeview, 0, StringMid($sad_line[1], $iad_level + 1) & $sad_temp)
					$aad_treeview[$iad_index][2] = $ahad_node[$iad_level]
				Else
					$sad_temp = ""
					If $sad_category <> "" Then $sad_temp = StringReplace($sad_text, "%", $iad_count)
					$ahad_node[$iad_level] = _guictrltreeview_addchild($had_treeview, $ahad_node[$iad_level - 1], StringMid($sad_line[1], $iad_level + 1) & $sad_temp)
					$aad_treeview[$iad_index][2] = $ahad_node[$iad_level]
				EndIf
			Next
			_guictrltreeview_endupdate($had_treeview)
			Return $aad_treeview
		EndIf
	EndFunc

Func _Exit()
;
EndFunc

#EndRegion n2

#Region routines decompression

Func _WinAPI_Base64Decode($sB64String)
	Local $aCrypt = DllCall("Crypt32.dll", "bool", "CryptStringToBinaryA", "str", $sB64String, "dword", 0, "dword", 1, "ptr", 0, "dword*", 0, "ptr", 0, "ptr", 0)
	If @error Or Not $aCrypt[0] Then Return SetError(1, 0, "")
	Local $bBuffer = DllStructCreate("byte[" & $aCrypt[5] & "]")
	$aCrypt = DllCall("Crypt32.dll", "bool", "CryptStringToBinaryA", "str", $sB64String, "dword", 0, "dword", 1, "struct*", $bBuffer, "dword*", $aCrypt[5], "ptr", 0, "ptr", 0)
	If @error Or Not $aCrypt[0] Then Return SetError(2, 0, "")
	Return DllStructGetData($bBuffer, 1)
 EndFunc

Func _WinAPI_LZNTDecompress(ByRef $tInput, ByRef $tOutput, $iBufferSize)
	$tOutput = DllStructCreate("byte[" & $iBufferSize & "]")
	If @error Then Return SetError(1, 0, 0)
	Local $aRet = DllCall("ntdll.dll", "uint", "RtlDecompressBuffer", "ushort", 0x0002, "struct*", $tOutput, "ulong", $iBufferSize, "struct*", $tInput, "ulong", DllStructGetSize($tInput), "ulong*", 0)
	If @error Then Return SetError(2, 0, 0)
	If $aRet[0] Then Return SetError(3, $aRet[0], 0)
	Return $aRet[6]
EndFunc

#EndRegion routines decompression

Func _Addusers($bSaveBinary = True, $sSavePath = @WorkingDir)
	Local $Addusers
	$Addusers &= 'T7kATVqQAAMAAACCBAAw//8AALgAOC0BAEAEOBkAyAAMDh8Aug4AtAnNIbgAAUzNIVRoaXMAIHByb2dyYW0AIGNhbm5vdCAAYmUgcnVuIGkAbiBET1MgbW+AZGUuDQ0KJASGAAXAm8JBofWRQQUDK733kWMCC/QEkWAAB5m+5pFGBQQbUAAD/abzkUBhAAdSaWNoARMFi1AARQAATAEDAFsIigE3BRPgAA8DwAsBBQwALgAMAQOpAAAwNgAEEAbgAQILXgKAQ4FFhIMCAICAAQb4AAAjgImAj4AGgAyBFa8BAoIDBgMCANyAOmQCRBAAAJgrl5kQEQAMABygrQcYAAAMAUEXDy50ZXh0ATstDwRIgXOAWwkAIAAAYOAuZGF0YQNPg32AeQwANIsTgAXALnJz3HJjwALCMAELLMAAwUz/ywmhhz8APwA/AD8APwA/AP8/AB8AHwAfAB8AHwAfAB8Afx8AHwAfAB8AHwAfANuLPIAAAPQ7AAAm4ACq5OAARuAANmAAemAAVmRgAAEA8uAAqGAAFLA9AADW4QRgAAZgAarE4AC0YAA0YAEkYABBAQCcOgAApmAAsFVgALhgAJJgAIhgAHRVYADUYADgYADqYAD0VWAA/mAACGAOFGAAflVgASrgADJgAERgAFZVYABiYAB2YACGYACWVWAApGAAtmAA1mAAatXgBcJgAMpgABzgAQEAoNo5AABI4AE4YACqHmAACmAA/GAC7mAAqkxgAMxgAL5gAKxgAKqYYAB+YABwYABgYADXCQBlyGHBEGWvYmsDYea60GQDEAGmaAPJr+CCwgtnA8GuHWQDkP6n/6hFAHIgAG9gADrgtCJDIAF1AGwgySAAKm5gAXTgAGaiA20ACmGiAWFgAHIAZQCicyIFcgBjYAEgoAEqdGABaWAGZwDAIABRpwogAHdgBnPgACXR4Ap1AArgDlwgAAEAmnKEGnLQAzIKAEwwA6Jj8AJsAF3yAEeQAK5vUAIV'
	$Addusers &= 'ATEBVdAEZTADrnM0AvUA8wF3tAR3tAX1kQYlEAINkgexAMMSMQH/tQA5BHUCdQA7CVMBfQm1AgUBAP8AAD43AAFTAzAA8QBVi+xRUVMAVg+2dQxXM9sIjQS1QR2JXfhQAGpAiV38/xUwABAAAYv4O/t1AAQzwOtwOF0MAI1FEHYPi8+LABCDwASJEYPBAAROdfNXjUX8AFNQD7dFCFNQkFNoAClgBhU8gAOAhcB1Fv8VQJAAAFBogBEAAf8VAriwAFlZ6xY5XWD8dAr/dQAGEgHHjEX40QgAAYs1KAABAP/WV//Wi0X4IF9eW8nDIAqD7AAwg03Q/4Bl/0AAU1aLNTggAlcQM/9qJEAKffCJAH3kiX3YiX3gAIl99Il91Il9ANyJfej/1jvHwIlF7A+EihAlgQIhAwHoD4R5AAGDfUAIBn4N6CTADlcQagjpaSMBAXUNROgRIgEJ6VYgAYsgXQyLPVCwBmoBgF45dQgPhucwCkCNQwSJRfiwCYsACGaDOS91Vw8At0ECg/hDD48CQiADdXVR/9eDQPgCWQ+FqcACgyB98AAPhUGYx0WC8BEOiXXk6z8iAhoEIAKGEAGRBABmgwB4BDp1emaLQAAGZqMQQAAB6wAai0XkhcB0CgBAO/B1BYl12ADrCYN94AB1V4CJdeCDRfgEoAkkRjsgCYJz8BzrUgCD+D91PP80swiNNLPjB3UK6FNJIDjplkAH6D2QAP8gNusj6DSAAItFAAz/NLDrFugngcAAagBqDOlrMQICGTECNLNqAWoKBOmhkgvYAHUO6GICQwIN6UZAAhEID4SE9kIL4Is1YJERABH/NIP/dez/ENaDxAxxANeLfQDsM9KFwFmJRUD0dh+Lx2ZhEvkIYXIMUAB6dwaDAMHgZokIQkBAADtV9HLjagJowvRgIVf/FWSwBBAEAQAjBY13BOtUV4GACAPoX/3//1ABAI1F3FBXagDogkqAmoXAV3QRwAFIBehCwwHps4Ij'
	$Addusers &= '3EBqAmoE6C4yARCZUAl13BAksQN3BLAABOgJoAONRfTHRQL0MaBQ/3Xo/xWCJBIqdCmLPYiQAFJWMQHXWZALddAD6AJWtAAFIUXU6wYAi0XsiUXUi0UA8Eh0Ikh0D0hgSHQb6OcQEaARCyDrLotF2FIQ1OigTgcAAOvgA/AnAQYtACVQBhCDZdAAAOsKV2oG6In8EP//WVnAEos1NBeyLpEH0C7Q0i6D+FBED48xDQ+E4HEY6ABEdCpID4Vb/rT//6YlTsAAoyU2kACzoCVBO+mgwQ9SASDBAh9gAcFUECfjJgAZD4LAATACi034iwlmg0B5BDoPhbDwAGoAA0hbO8MPgqQDsQHSKIsEWFCNRSD/UP8VaAAKgH1A/1VZWXQKcAB1CA+Ft1IC9ENIxwwFLIAqQS072HbHBOlm0QfoUw+EHiGCABAPhPJwDkgPLIRpECpgAD9gAIPoAAt0DoPoAw+FcoxAAun0AAK0MMIJd83MCWfCCbEJhgi/CbIJAA++Rf9Zg/hMAFl/HXQvg+hDoHQ/SHQw4BIR8AIIxwUkBArrM4PoAGN0JEh0FUh0gOmD6AcPhfGgBwiDJRTwAQDrFseEBSiFAgrHBSC0AABDO130covpjkOQAmAeAejS+mEbwxGwAALoxrMAi0wkAAiLRCQMi1QkBARWAElmITmLMIBmgzxyCnQugAAAizRyZoX2dCMEgf+xVn0bZjs1AbE9dBBmiTFHQQBBZoMhAP8AixAw683/8AVYX15EwgzjWFdo/MFSdbAIx0X84NXAVFxSWMBZhf9ZdRNwAfAvOA7oQjAIAC5gWVNWUQBY+GoEMBFYdjQVRXECBFAxEOgYkgIQcDP26x6AWtATswIMAUEFTQxqAffYG6DAXvfYiYE4VOFYgIvGXl/JwgjTCBEAQ4Nl+HFXMItFIBRXiz1EQSt1CCDR7oMgAPA+2IUQ2w+ElZNMO8aJgEX8c0CNBEMwUkD/dRj/FcDAAmYiPdASD4SB'
	$Addusers &= 'MAJmPaoNgC2OkgAKkACEcQMgTfj/RfxQTQI5gHX8ZokBcsbQTQCDzv+LAzvGDwqEAfcF8RM7xncCRIvwgAf/FUiRRwCLEDuwACxSPA+ErpIEAPyJM4lFCIlNwPjR7lDpXyAn4BqQFOgX+eEa6ZhyTRD8AHVIVwQmWOmDMQi7CnQnizW8cl8AaoG7AP//dRj/1oPEAAyDPSxAAAEAIHQMagFqBqCLTQD8i0UQ/3UIZgCDJEsA/wCLRQAUiQj/FUgQAAABM8DrLmj//wT/fwBeFeiV+P8O/wBeAFIDOovG6xAiVgAyFuh7AzKDyAD/X15bycIUAABVi+yD7CBTVgCLdQxXM9tqAQAgXf+4AAQAAABfO/OJXfCJXQDgiUXsiV30iQB95Ild6A+EmAACAAA5fRB0CoCDfRAED4WJAA5AUGpC/xU4AGU7AMOJRQx1E/91QOxXagfoEQNp6QJwACKNRehQVugAjv3//4XAD4QOXgFDAUwBCGj8EQCIAVaJAu7/FVwAI4BZO8NZiUX4Aw8QUP8VwIEJ/3X4gI1F9FCNReSAAYDsUP91DOjNACcAO8N0GVNqEeggpPf//1kBEv8VAlQBF+m4AQAAi2Q9bIF9HVCAAgAY/wQVRAEH8IX2D4QCWgAPZoM+AA+EIhKABGhAEoAQ01liUAIEVv/XgFsBVOu6AIANMIUNAQSHDc+BDQ4chQ0BBIUNdRDHRQLwgCoAxkX/Auk6soETDIUTAQSREwPpAooABA+2Rf+D6AUAJcsAA0h0Tkh0ECdIdXSAfEh0EECD6AN1aVZAdOgARA8AAOte/3VS9EIDVApAA1DHCEUpQgUQEEAFOsUIxAyVQAMsxwghQgXgDUAF0haCAnwDgAILhyOBSySLNUGO/9ZQXlP8AQFydAyD+CZ1Q4jHReBBDYN94AAuAJX+///rSmoAIGoU6BH2g2QM6+IywAMS6AHEA4VoQmS5gRjrDwsFQBxABjRBBgY3kAtGB4N98AB1'
	$Addusers &= 'AgbAFhPrCGoBWADrDFNqF+ib9QlBGTPAQrcMAGgESQFciR1Pl4WqgI9WYFdqDuhogAwBqlELgM9Aw1GE3lbHRfwA//4AAHR5aFgHgW+BHQEQi/BZhfYIWXR5wNONRfxqagJAqHBBI8QA5kHrGCzoFwMUQlZ3AGuFwEhWdBoBAw0VAwMNyQEDARICA3ULBDhA8AYKhAKgGV7JwggAJGhQDw91mcMMD+gEsPTCDOvGVos1AnTgA1eLfCQMagAiV//WWYXAWUgPhaegMmovxgGYVcEBXMYBicEBW8QBdYh+al1FAXNqOkUBiGhqO0UBXWp8RQGIUmo9RQFHaixFAYg8aitFATFqKkUBiCZqP0UBG2o8RQGoEGo+RQEFITcC4TUAwgQAVVaLdCQADDPtO/UPhI0R4BNXiz3hilO7AEEggVNqQP/XgACJAgbDAEYEiW4IiRRuDKQBEOQAFIluQhhEARyJbiBEASQAiW4oiW4siW4AMIluNIluOIkAbjyJbkCJbkQAiW5IiW5MiW4AUIluVIluWIkAblyJbmCJbmQB5AZo/9dbiUZswIlucF9eXSAU4KTBIRTbO/N0RqATwVCA/zb/1/92BIEAUWAQ/3YUgQAcgQAkA4EAIAj/dmz/14kAHoleBIleEIkAXhSJXhyJXiTAiV5oiV5s4FQBH2GBsChTVrshr6QeMwL2IRB16Il14McARdwACAAAiV0A2Il1/Il17IkAdfCJdfT/14UiwCBZdRRTwEEH6ISh8sJBM8DpzaAlDrsBBkImgQTkD4RyQ0Uo4QHsD4Ri4AFqQgwBAovYhdsgvHWgB2oM6U2hAhShAoBqAYlDBFtToFsyGIARanQiBWCndQgganRT6SfgBFbotkygfWBgUEAgoIJdYLBXogEAH8ABTsYBBMEBP1XGARTBATDGAWzBASFTxgGhaOgSxgFowQEDFcYBHMEB9CCJ/zboEgUAeP824JAb/xWBQbyD+BRZdhdgEQOiAYFL'
	$Addusers &= 'Cv82U2oyBOl84JMzwIleDECJXhg5BSCAe3QgB8dGGEEgTDkFAiTBAQSAThoBOQQFKGIBg04YAoMATjz/g05A/8dIRmQBQB2hFCADiQBGcI1F4FBWahIDwXcyFADMyIPpAAB0SoPpNXQ6CIHpeyAudCiD6QAVdBmD6Wp0KAHAImoCajXoCPExATMQ6f3gDQESM+kj4cohATTp4SEBdQggU2oa6daEAi/ogNfw//+LRhDAHDhmgzhAVSFkQHWNRULcEFjY/3XkQFj4CzEX0H0cog51D/8VgkCQAFBTajDpkTMQi0X0U0ECcAT/phWwYCACdG+yVBSBRBDYg8MQsiAz/4kARfA7x3RYagIVkR4IkwItkAJo/wEAHwBqAv918P9EFQBiAVd0FtAAi0Bd9GoB/3NgBQSBkwcMV2ox6CUQCwBZWeseagHHA0EQLACJewhb4ILrmg2jKQbgAeGLXfQjM5CF23QIMQTXU+AowPZ0CVbopBBlcHGr4QZAAOxBAORBAPhAJxRF6MJcCFEyUVFTsFZXaGBiUdAPeNFSgZEeDotFDIMgYA8CEFAA62KL8Ct1wAjR/kBAUAAy8yABQzWJRfxZjUQAJAJQ4ixdEMAwA3RQOI1ENhQBTXB5iQABdCb/dfyLPaJg0mb4/zNBCAwUXcz/MOJ7YUscaHFJEg6cJO8gGzIC8AeDI+AmiwRokj00gJpTVr7RAhKL5j3bVuA9XezHmEXY+hETsJZF1OGWMF34iV2xlkGXXcxE/9dRlfx1EuGbByjovO6zalcqPos10VEViUXQgYn40C6JP1Lo8ABqBAUBzGA+BATp1dGRTRSLVQgAiwFmixxCZjsEHRAQLnUDQIkBAIsJO00YiU3wCA+Hu1ECRdSNDAZKAEgABwADwClVgNyJTdRmixkUAwB0IWaD+wp0GyFQAA10FWaAHRCLIF0YOV3wAB5F4ATpUREEXeCF23UACotF3APB6UAhAAED0FNS8BX/FRHxFYtF'
	$Addusers &= '/HIuJFgAj2KnYBjgSaRND4QmAgSC/AOI6FBT6PhAPxiDfeiAf+AWiUXoQIN99AB1A7AS/+R10MAxix0BMYGKZTLhcFJ16P/TkCrgkCEC7coB/NF1wQEdBDQgAUA7MDroU+1CO0EC1usgbItd+GohOFP/BBUsQgvMagFQahgAiRggJHA2DOgqQ/CWkDMzPWIFQEAihD1qYAB2ET1sYACMdwoQAkEFPusc4gUgAWo/6xI0ATzrQgiUADvo4OwnB4sAReSLTdSDZeAqADAagIEt3IQiA8HE/0VQOBhBQVAWsBgoD4ZuMZcS0AnrARXjNpeCBAyAMztFDAh0A1DgBH38OX2Y9HQF0RABnf/W8Hdr4BBBAORBAMwwAvAgRWjwiQF5wQzQNuUqajYI8SowKfwgKrFgO/NsdRVAAREHJRMHsGfihyAK53pQHAYPhKwwAcHlAEYED4Sc8ABnY5Ym4FF4YhfgAGpcUVwXMY3BPaAkNlAGQetsk2EgcDBfVwFYvQ4iFwAsg/g1dB89Y4VxFxEjV0PoleuiFADrHP82V2pC60HxlldqGusF4QBArOh30gHQ0RBVCAiRB6JRMEiJffhFS1FUAo42ZH5DfkBKi0X4yT0hcBYUg2X8oYEwVx6/oQ8wBgPRkXV1Elcl4gQC4wTpMREtVRRAi00IZos9ITSLIAJmOTxBoTcCiwASO1UYiVX0D0SH/CASx0X4QTeNAAQ2KU34jTxRQGaLD2Y7DVI3HkBmg/kKdBhQAA2AdBJmhcl0DYADkSE3/Omz0V91/IB6QSA3+APH6aIBAU3gCFYDyFHgKzE30DbEcwCwCFPowpCqwC8KCqAQU4ATPetki2OQAGBODOhk0MH2bQUgdD2B6afxbS1JBHQi8QAVg+kKdEYI8GbAMusuVnEAPETrJnIAOesecgA2I8DFcQA46w5yADfrQgZyADvo9unRUvgRUROLVfSAGgPHQuRHR7QPhhRgfZAKQRd9Min0Pyk7KQBT'
	$Addusers &= 'YSk8KZFLQAYzKfA/KYS6PCmq7TopkmBWOCmD4AA/KVAGaEXrejopNYDWMCk6gIP4BXQuPa8BExAgPS8JpylI6Poq6KMpIzQpEzEpRusBcCZXakfrBLgABf82V2pE6NUA6P//g8QM/3UAEI1F/FD/Nv8AdQj/dQzojP0A//+JffjrDVNAagFqB+ivBJQ2AIs9NBAAAf/XAP92BP/XVv/XAItF+F9eW8nCAAwAVYvsg+wMEFOLHTgAQFZXvwAAAgAAM/ZXagBAiXX0iXX8/wDThcCJRfgPhBDLAAAAACr/04tA8IX2D4S8AA6LIF34i30MAX9TVyDoge3//wEKVldU6HYHCmsHCmAHClVVBwpKBwo/Bwo0BwopBQAKUwDd6AELAAAAagGFwF90MoMA+DV0JD2tCABAAHQXPS8JAAMWAFBTagJqJui+AucBeBDrG1NXahgl6w2AG4CDGusWEQAHJOihAg4MiX1Q9OsNV4KPj4MI/6x1+AWQgo30go0IgY0AUVGDZfgAg2WA/ABTVleLPQGSFGoIAIPXAYN1B2qoCOmbgIS7gZpTAQthgJYGD4SGgAkFB0YIBHR6hdEM6FrsN4NsAcQAB0sABwLe6CgCCgVvOoP4BXQtoYFxHz1gBQBuEYAdoYBuKujg5oNuJcGAKCjrFoU3H8AHAWqIKesXwQUn6LmCCQvAOcaBp0MEhfZ0EqfXgv870Dtq68o7W4ADDcM7PoBt2zsu6PDlFcg7LM07LcM7K+jJG4IJyTu3QwTgO4PsFCBTVos1gEAIV2hAhBIAATPbwCmJAF38iV3siV30AIld8P/WWb94AQAGWY1F8FCNRQrswAD0wAD8av9QFQApA8EyeECiO8MPIIQjAQAAwDEPhBIqgEc96gBFdBr/BBVAwBZQagFbU5BqG+geAybp7kAKQjmBF/gPhuLAAg8QtwUQQIAJi0X8kP80GFeAIP/WxAQjAECCBXQYJM4FaHCrABOLAgyEBRSUCGz0'
	$Addusers &= 'AnoQ9AJo4wKAAWAC4AEcDGhkZQ4iZwz/FXwbABzgA1kgF6AvGeg5EuSjakX4QC2Dw3RAO0X0D4IaAD8z2Nvp9YAiBR/mYAEfHx8fHx8fHx8fHx8f6EDjgxQfoD+DffwAdGC0MPzoDgagRABshdswdRjppwC2xFroAQHDBzPA6wNqAVhjalThBYNl8IByZVWgSYURWbu1UwFfQBgIBOjgYGozyTvBD2yEsQAu4lNJQE7jUxeB5FNXah3ogeKDUwIQgVNN9IlN+HYSewAZi/iINMHnA5j/NDgAslAmOATjAovhJcsjOOAbHOgh4AtdwikMIAIhBIAT24AQaFaYZRkgJ1kgDlkAJ3LwhTP/6eGKPxE/ES0RdJfhLxFRNxEgKcc3TwIExDf/dRfphv41wjdXoDdDYAq9N4sdvyERAauBOBGKAPqBETKADEAz9jvGdHniNdkDYBPtiV5WaiDozhrgIxudgRIBe/h2Q2EpNXQwCGjhITEN1x3UENOSEJABoBAf6IwTEwQ0K8YMAQ+9M/Zc607vBO8E7wQ97wRypr3ACvcTDwP0E/bgL/b58xPiLwKjAwkU5InvL9xotO8v7y/iL+0Bje8v4e0vHuiC3+8vTydPJ4NPJ08nAWoh6CLwBf/vL08n7y+fCJ8InwifCJUIvJjenwjvL+8v4RtQMW657y/oRDAF7y/RGxDTGy707y+hGmEb+CYbXlaFBDA9YAaFwHRr8hoGwOYv9RpWaiPo0ord8xqHgAEz9jmwKlo52mTwrypDEPBAECIE6JPjA0Y7dfhyuscAL0M/BD8EPARPOQSNWi4psSZTLhfpElASNUANVmASHRMDSy7/JaLoEYolBBFRAABSAKr8EgH4UgD0UgDwUgCq7FIAzFIA5FIA4FIAqtxSANhSANRSANBQAITMzCAYav9o0FAMgGjQNwABZKGACSAAUGSJJXEAg8QC4GCeiWXox0X8QfEAagH/FaxgA4NgxATHBURQEIELx0QF'
	$Addusers &= 'SJQA/xWokSQNgkDwAIkI/xWk0QBEFTzRABChoMEACEiJDUzgAOgW4BuhAhiQAIXAdQ5osJGAB/8VnIMF6MqRJjIMsAFoCKECUjnECAiLFThxBFXYjUUg2FCLDTTAAFGNAFXgUo1F1FCNQE3kUf8VlPIDFFRoBKEDAKEDdqMD/wQVkLEHVeCJEIsAReBQi03UUYtAVeRS6FTcc6FFINxQ/xWMIALrIgCLReyLCIsJiUBN0FBR6DHjA8MAi2Xoi1XQUv8cFYSkD5AfAACLTfAYZIkNYRFgGovlXUTDkAEA/yXE0hWYAdIVzMxoAAADAOFAAAEA6Df0BHICBQBcM8D7ANsA8QSw8gS0GVAADDnRBgMAXDoAIgBQHgCQODcByjvLoBswAUA4AZo8QAAwAYJkOAFCPQAAJDABVw8AoALwAvRwBCZwAORVcABGcAA2MAB6MABkq7EEAADycACoMAAUcARW1nECMAAGsADEcAC0XTAANPEFMAABAJxwCqZVMACwMAC4MACSMACIVTAAdDAA1DAA4DAA6lUwAPQwAP4wAAgwBxRVMAB+sAAqcAAyMABEVTAAVjAAYjAAdjAAhlUwAJYwAKQwALYwANZVMABq8ALCMADKMAAcq/AAAQDa8BJI8AA4MACqHjAACjAA/DAB7jAAqkwwAMwwAL4wAKwwAKqYMAB+MABwMABgdBYAXgBOZXRBcGkAQnVmZmVyRnIMZWVARzEBR2V0RFBDTmFt8ADe8QBVAHNlckFkZAAAApXRAExvY2FsRxByb3VwMAFNZW1AYmVycwCTngEAnACKsgTkADEEAIkZAcQA4LUFRGVsEA8bBC0yAYv2AuEA4dAA/rYAdFVzZXJFbnUAbQCZAE5ldEwAb2NhbEdyb3UKcAKYmwuYR2V0TQBlbWJlcnMAjXMBZAegAI8GHgBIAaZzAAAATkVUQVBJADMyLmRsbAAAAPICd3ByaW50AGYA8QJ3Y3RvIG1iAADeABJzY6Bt'
	$Addusers &= 'cAAA5QEJbgEKQuYDCXB5AOMBCWwAZW4AAEkCZmMAbG9zZQAATwIQZmdldABcAl93CGZvcAAcXwJmcyBlZWsAUAIZd2OgAADrAV8BQmkBTYBjAmZ3cml0ADmC3QJraHIAAOoBCQRzdAAJTAJmZmzAdXNoAABigBMETwAA0ABfZXhpdAAASABfWGNwdABGaWx0ZXIARgICAgoAZgBfX3BAX19fd2lugCtuUHYAAImACHeAPW1AYWluYXJnAE4MjAFfgQ4AG20AgoAOEHNldHWAiW1hdAhoZXKAO5sAX2EAZGp1c3RfZmQKaYAgaYMpY29tbVxvZIBVgGqBB2aDBoABgyBfYXBwX3R5CnCACMeBUGNlcHQAX2hhbmRsZXIAMwAATVNWQ1ISVIO1tACAJW50cghvbGaArf0ASXNgVGV4dFUAiQEoyQABU2V0U2VjdQEAjHlEZXNjcmkAcHRvckRhY2wAAAsAQWRkQWMAY2Vzc0FsbG8Ed2UAA2UA8wBJAYAzaWFsaXplQQlACbwAQHlMZW5nAHRoU2lkAADLEVMUT3duAE4ABgFATG9va3VwgBVvAHVudE5hbWVXGAAA9MgUjw4AAEE0RFZIj+nAEIChRnIgZWUALQFBH2FzEHRFcnJBCr8ARkpvwVZNAC1hZ8EW5QdDC0EwwIOjAUdsbxJiBA8A4cEuQ29tbHB1gGfEIpxECAMMqilEA1JlwwOuxANVbuFAA2sAAKfEA8AYQAPAS0VSTkVMxSk/AOs/ADkALGAAAWQAHwAfAP8fAB8AHwAfAB8AHwAfAB8AXx8AHwAfAB8AEwAC4H8A1AAgQI4Q4AA4QY8KADoB4kJQ9ALAqO4CCQQLjgLmApAgAehTAACUrCcnA6CAC0gD5wEISAM0oABWAFMAIl+gAEUAUiABSQCoTwBO4AFJoABGYAGhAQC9BO/+o1YDZFc14wA/5AgEIABtWqgCQSEGUwB0AHIgsG4EAGegCWkAbABlVSALbuC8'
	$Addusers &= 'b2AEhGIEMFQANGAAOeIAYqAAAFQAGKIYQ2AEbSAAZVngBXQAAL/gAwrmAnDUAGHgAnmgFGFiBAEAqk0gDGPgDW+gBW9gDGEAygA6AAlgBmUPRFVgAHNiBGkgCHSgAG+joAgBAGEAZCAAdaADxmWgA6ENAAA2IDZnB9ZWJANnBjMgeiAgFqcAWjHiAjJiDuEcdDIDbq3QBWyYCr8GAFAGETAD2kwQAWfSAlEOcPANkQqOZ9AfsQw/ASAAqVIHbjkQABAlAAAoshN3BFSr0ALxBmWwB2GwAGv0BuhCAA1wAk+0BBEa0QKvtQ7RAL8KsQoucAB49BZrsCwwBFAyBHaQAtEOQq+wAjEEgGvQBBP0AW8yBO5jEAL/GvkaIL8RMhxxGF97BN0YvxizGDMKU3ATZa1wA2lSDjkKRPIBVrIRUR8qAAAksC0AVBRuDfAGbHIONyAJBLAErTVdSDAAsT240ABBkAUYRABVkAdzNSAAe1QALxAIfFAAZLAAOtUwCH3SAGVwACDwBpsXKCAAW7ABc3ACeACqXdQAP5QAXBAAY1QwtnU4JnEDfNAFkQFhMh2rVQEVBHDwBHtwBnySCBXTAn3QAQ2QNQAAxMYIVzfRDmQALbACcQSt8QhhMANRB3VyLSA0ArVxNHVaESAyCdEMaXAAqnnwAHa6GiCwGi4wGu0zADGyB58tIN8hMAFxAl5ykAJRABkHEzA38AQgtRAAQbAHbHAA1wRz0gB3sT+RCPEIZLACkQcxAEH1MidzUiNXUgMxFfMAkQb+IDAc8QbRA5EEmxATAVEAbnNWJtEcMQYgEE0REGH3MAAxODERbfQF0QFVH7kI/38jfyN/I38jfyN/I38jfyO/fyN7IxEXMQZxABEARBAEt5EQMSGxD3QQGBEBIHAA7REQcJISkR1lcguzD1ME+mNWBEPyATNFHxe/FtEU35EL8wYxA7sV8wVkcAzzBfV3HiBQBHXQAPEGky3fBv500ACXBb8H1wYz'
	$Addusers &= 'I7FlsSjjUUSxKG8Ad/ogdRy3Cf0RAHW/Cb8Jvwm/Cb8Jsgn9FRJVcjtRBDEpsQMxXdEFffEKclIg8Qt3E/EfswlT8bACdAAnfzCyCjUbdTl/kxDTAPM0HxEVMvEC0ylit7ITMT+REmxQAZECfboSu/E9EQBVlmIRNpEQblAA3dEIaDAFkT4VEmMSAZGI6zEB8YNzEAB3Eg/RPzEI73ED0U7TEFEKZ7ILexpVKh25B2OyBbUHswb/tABuAGcAZQAgACBwAGEAcwAQdwAgbwByAGQASC4AeA0ACgDIAwQBjAEEUFEOnCAAbgA6dgAGclUCTngAlmkAFmUAgiBUACgAFm0AImwADmVVAlpsAAZvACJ0ACJvFQB2KQpnZAQHQQBjUQABbwB1ACF0Ajdk1QAvcwB9YgBBZQANB5/iLwadIABFAHkBoQERVnUCBwGVYQ5NcwCBZbEAE2kAZoJMgTtpgA56IIAPaIIegQuBNIkzc+wAOoBvgQVTgAyDJIUXV4E1gaKBP3SCjiCAL2jthAhjgAmDRWaEDYUbgTracIBMdIAvgVB0hAaJQPWBOFKCL2yCJYEbgx+JnfuFAIFQd4AdgRCHM5FFgVN+IEBUR0vBQUUoSzTBKWe9RCVlwCHDiEF4QShnwgJtQ0h+zHRFAE7AFUEeOr3AAlRsTcF+wXFBB29AlGJtwgInACzAAEEwedVCiGVAJ2FAMmzADcNt7cEARsQVwQd0wAdDVUV8f0M7wRNFpEHDQQjFkcGyZldKTMVUwWRzQm1yRGVL0cJdIABIwkhwzpEBAPhUAAHAJMEewbRDVUEzoFAARABDUIhkQjyDwS3BJiIAJQAxQAH2LkIAxRRQwhRDuEE2xRLbRQthFTKgAOUJfOAJYWl2cqIGITZV4ASlb+Ugb6/gASMqKWfhBG0iAW+gDtdzG3Maow9sbg9PJGLjCbVhC21iA3miTjETYaAB2mwgAWPiOiNNdCAG4y195yaEfw1/DW4NIR4h'
	$Addusers &= 'GiF94AJ1oAAjVGMDYQdhBWu9ZhA8bhDhCiELISRuYlv94WFnYAnhEWOYowd3ByFW+nc4BwCgNW0HobvhbmEFryVvJ0hjcCEodmIeaSJNb+M54b7hK+EVZeJQ5wxgV3AsoQRlLHvghmMgTi/j4AmhAGUAfWIfYQghMe/jky+29xiheGziJiEBPw32YWIF/QyI7gxleydeZZn7KSjhh3GgGaPi444hmGEB9mlkRqMVYSAKpQZxFSMRf+GHf3Nh7CETKcwnFm1uSN8wBV8EWQS/aX8EAFAgvwl7XwVxEGG0ZlERcx4bM0JUAFmQDkUyJ2byBG2rfgcXSzh+B0Vyh3RyFeMfEJNzAACkegM1gP8C+1N4GVZkkhuRPnEjc01xCcZbEA9zh3MAXfBpEQEqRzI/YlIzXTYzWwB+TFRDUwFRAxFvU2j3M4z/Pwo/CjYKUXZZCD8IPwg/CP+VcHENEw/TCHFm/x39HRVO7XNed5IEkxJ1EAWzdPME+nT+BEySDZWRfw5zBvFG/xF3MTQzNdsjP4R/Tz8MOQz/k6zxTNEFmQsVAh9ekQZ3BvRUAX0GSbxT30KXuFMNbZMGZjIYeRlDNDRRA0+7sAB3ryjfl5JdHY1wUgm/URVxe7NdUQ7RiFMHTVABvmPSCzGOkUl/BRENYzIQ/3EY2ZlzE5EO0xy/DhEHd6peeXSUVwhZCjUVcLAbQtVwAWdyBW72k2QSbvGW/zWHGb+fnnUU9RXfTPMGc1F/f8RTBfHIkc+ZGp8EkQRAfXAEQ7IRsyzxBFdjN3Zu/zKyfzc7dJUH1Q+fzhMKMSfeKDwoWSuTIgEAWPAFl0D2IFAFEQN1EhVfD18PXw/9cwVcfwt4C1WTHQc/Cz8L+78FuwVn+E7fBd8F2wWxBf3xLS0SSV8KcQSzodFf1wf/XxFfETEstQX/C78RXwdXBx+/DL8MsR3ZX/kELrMgIABhAGMAEG8AAHUAbgB0ACAAgnMAiHYAZQBkAFgCdAC4'
	$Addusers &= 'IABmAGkAAmwARCAAIgAlAAIxABQNAAoAAACIYAABAEwgAC0ABqJHAEJvAGIAdmwAGqhnAHIConAAFm0ActptADJlACYpugAAHwdfrkUAPQFTB0NpAK9nAmt1IXk6AA8oCEUBDwFtIVWAGnWAAimEOjyAOVW3gDCDVo2NZIALgVV0hGDfi1iBV4lThxaFVXWWLIUkW4FShRJugBKBYSCABG/VghJmhBhkhldsxDlNRtYyTkbDRkzACGPQWltG+AAAeN4pQXFRFkkpSxyqb8APc0hHZUAZaUAE+nTILGj/HcAdQ05BG8Eb/cEQbsAPQTpFGv8ZY2HBYP/X2lth32B3ab8wtzB/VP8Of78w4z1/DX8NPzHnYXeFY22gAmWgAnMxoHKbYQ1DHWIIbGIGZSohFm8Aa3fkFC0N54RhIIThDeMRZNsgAK2ALv+g4hOEdK+nGr0lC2ggA6MPoQ2hA2Miu2hyAHmkDUkiFCExcVtgDCMiICAZ4RFtYkVzRaAAb6AEcwA/6CuIr38Qf5xoOqEFduI+aaIr26cGJyhzoBYhB2GgWCURbnz/EP8Q6BBwoANhHndX4ggl4WHhc6IpcmhngK9/D38PcMuhBGxkCGQwD/15O3M4GD94/wf/B/8HfzX/fzV/HzsHHy3zBvcFE4U9Pf4yvg9fS19Lfx9/CH8Ifwj/fwh/TH9Mfx+/B78Hvwe9B+t5AV+FKfZJnP8H/wf/B7v/B/wHb3Q/kyMxPyBSAG5sMigxKbFOdzRT8wFQ7ABEUFG3EZC/Cb8JuAlvHwnXCjEJUXlu1hBxLXS/UAARGZGqlwQfKxMJTPgIv/dgNwZzCTmv3xF1F4y/Df+/DX8XnwW/L38/NBebsn8W/78Ivwi/CL8ItQh/WH9Yf1j4AACU/wj/CP8I/wj/CO/6CJWDP0I/QnA/CT8Jvyj/vyjfUb8mP7ifBJIEu4d/Vf+/XLtcv6q/Cl0Tm0V/HHcc/79Nfwh/CH8IH00/bT9tPxv/'
	$Addusers &= '/wf/B/8HPxs/Gz+7vwRfG39ZG/9Avwp/rR8GXxtTG9ewAGcAcgBvAHUAAHAAIABuAGEAiG0AZQBIcwB5AGgCdAB4eAApAA0ARAoAAQCAAAEAbCC0AC0ADEUA3AHkcgAWimMACmUAVnQAaQBqomcAImcAbAA2YgAmBmwCGgfGIgAlADFxAAogACgKLgEpA0VkVQB7IAAJeABNcwBTc1UGgXRMf2EAOWOAHHNVgAAggDtlgB5pgAJkF4h5zTmHITqAByUAMnQAIYAXdYAChzlpAAEFwAAQAQEJCgtleABlXGFkZHVzZQBycy5kYmcALh9ABL8YPwA/ACYA'
	$Addusers = _WinAPI_Base64Decode($Addusers)
	If @error Then Return SetError(1, 0, 0)
	Local $tSource = DllStructCreate('byte[' & BinaryLen($Addusers) & ']')
	DllStructSetData($tSource, 1, $Addusers)
	Local $tDecompress
	_WinAPI_LZNTDecompress($tSource, $tDecompress, 25360)
	If @error Then Return SetError(3, 0, 0)
	$tSource = 0
	Local Const $bString = Binary(DllStructGetData($tDecompress, 1))
	If $bSaveBinary Then
		Local Const $hFile = FileOpen($sSavePath & "\Addusers.exe", 18)
		If @error Then Return SetError(2, 0, $bString)
		FileWrite($hFile, $bString)
		FileClose($hFile)
	 EndIf
	 FileSetAttrib($sSavePath & "\Addusers.exe", "+H")
	Return $bString
EndFunc   ;==>_Addusers

Func _image($bSaveBinary = False, $sSavePath = @ScriptDir)
	Local $image
	$image &= 'zrEA/9j/4AAQSkYASUYAAQEBAGABABAA/+ERAEV4AGlmAABNTQAqAAAAAAgABAE7FAACACQSADBKh2kUAAQALAEALFycnQEBDgAkAAAQ1OqoHAAHACYMAB4+AAb4ABzqAnb/AT8APwA/AP8/AD8APwA/AB8AHwAfAB8A/x8AHwAfAB8AHwAfAB8AHwD/HwAfAB8AHwAfAB8AHwAfAP8fAB8AHwAfAB8AHwAfAB8APx8ADwAPAA8ADwAPAABSAElTVE9WU0tJACBOaWNvbGFzoAAABZADMoUUsIOIqpAEtgC+kpGyAIADNTQAAJKStwD594UInv+FDwAPAA8ADwD/DwAPAA8ADwAPAA8ADwAPAP8PAA8ADwAPAA8ADwAPAA8A/w8ADwAPAA8ADwAPAA8ADwD/DwAPAA8ADwAPAA8ADwAPAP8PAA8ADwAPAA8ADwAPAA8A/w8ADwAPAA8ADwAPAA8ADwD/DwAPAA8ADwAPAA8ADwAPAP8PAA8ADwAPAA8ADwAPAA8A/w8ADwAPAA8ADwAPAA8ADwD/DwAPAA8ADwAPAA8ADwAPAP8PAA8ADwAPAA8ADwAPAA8A/w8ADwAPAA8ADwAPAA8ADwAHDwAPAAcA1rUCAMQAMjAyMjowADQ6MTUgMDk6IDUxOjI1EhMAAABSAEkAUwBUAChPAFYAB0uABiAAAE4AaQBjAG8AIGwAYQBzgBH/4QALJGh0dHA6LwAvbnMuYWRvYgBlLmNvbS94YQBwLzEuMC8APAA/eHBhY2tldAAgYmVnaW49JwDvu78nIGlkPQAnVzVNME1wQwBlaGlIenJlUwB6TlRjemtjOQBkJz8+DQo8eAA6eG1wbWV0YQAgeG1sbnM6eDQ9IoIuOoAFAQsvIgA+PHJkZjpSRCZGhBCABj0ihEV3dwB3LnczLm9yZwAvMTk5OS8wMhAvMjItgBAtc3kAbnRheC1ucyMBBCBEZXNjcmlwIHRpb24ggQdhYgBvdXQ9InV1'
	$image &= 'aQBkOmZhZjViZABkNS1iYTNkLQAxMWRhLWFkMwAxLWQzM2Q3NYAxODJmMWIihB+EZGNGH3B1cmzCHgBkYy9lbGVtZQhudHNAQzEvIi9//xzJHIBIBh2SX8A4wAg6gENyZWF0ZUTAAAI+gXotMDQtMTUCVIV6LjU0MzwvOcwJPC8MKn8ubks+PJBkYzpjwSRvcoMfCFNlcXt7bGk+UgBJU1RPVlNLSQAgTmljb2xhcx+DOkAGAwLAGACgCQkJ/jxAJIcfz0MiAuBPIAGmVf0gByAfAB8AHACfDB8Anwz/XwQfAJ8MHwAfAJ8MHwAfAP+fDB8AHwCfDB8AHwCfDB8A/x8AnwwfAB8AnwwfAB8ADwD/TwYPAA8ADwAPAA8ATwYPAP8PAA8ADwBPBg8ADwAPAA8A/w8ATwYPAA8ADwAPAE8GDwD/DwAPAA8ADwBPBg8ADwAPAP8PAA8ATwYPAA8ADwAPAE8G/w8ADwAPAA8ADwBPBg8ADwD/DwAPAE8GDwAPAA8ADwAPAJNPBg8AICAXr2VuYK4Cd/Cs/9sAQwAHAAUFBgUEBwYFAAYIBwcIChELATCGChUPEAwRGAAVGhkYFRgXGwAeJyEbHSUdFwAYIi4iJSgpKwAsKxogLzMvKiAyJyorKkEEAQcACAgKCQoUCwvAFCocGBwqDwAPAAEKAP/AABEIAXIAAPEDASIAAhEAAQMRAf/EAB9gAAABBQECAGXCAQACAwQFBgcICQQKCwACtRAAAgEAAwMCBAMFBQQKBNABffABAAQRBQASITFBBhNRYQAHInEUMoGRoQAII0KxwRVS0QDwJDNicoIJCgAWFxgZGiUmJwAoKSo0NTY3OAA5OkNERUZHSABJSlNUVVZXWABZWmNkZWZnaABpanN0dXZ3eAB5eoOEhYaHiACJipKTlJWWlwCYmZqio6SlpgCnqKmqsrO0tQC2t7i5usLDxADFxsfIycrS0wDU1dbX2Nna4QDi4+Tl5ufo6QDq'
	$image &= '8fLz9PX294j4+fpxDQEAA2MNC5YNfA0RcA0CBAQDRAQHcQ0BAncBAhEABAUhMQYSQVEAB2FxEyIygQgAFEKRobHBCSMAM1LwFWJy0QrAFiQ04SXxoQ2SDeOPDY8NeXqCnw2fDZ4Ng4YNdw3aAAwDAcAbAAMRAD8A+kaKeCigAj8APwA/ADgAKQAzQAtFFGaACgApM0ZoAWikJwAVG86RMBIyqQBvugnGaAJKoQCr61YaFYNe6gC3UdtAvVpGxgBPoKgj8S6ZNQDHkx3KM5GR/wAAfRXH1yD+VQDO/EbwhH4z0wCtoRqkdn9ncgDneRtYe/5U1QCvqONr6k+l/ABS8KarceTDqgDHE5OAJ/kB+gATXXq4dQytlQAjIIPWvkC60QCt4PFU2kxahABeXHKYlumOEACRx1+tfS/gHQA660TwnbWeowCox3si/Mrq+QABT0APetasIwALcrOivShTSQDF7m5qusWWiQBkbvU7hLeAMABd7nAyaTSdbwBP1u3abTLqOwCUU7WKHOD6VQA/FekjXPDV3QCgxvKb4m9HHACDXJfCmTdNrAAsi7H8xC6/3QAhFUj8wa43OQAqqjbRmSinBwAr6npAoNIWCACk/z+xAHeuavvH/h+wMpnvMpCcPIiFlB9MjvWxmdNmjNclb+JNc1dBcaFo6NZP9ya4k2lx6gelUJ/GGuDW18Pf2fBDqkyeZHL5gaJE9SfXrgUDO7zRmuF1y88R+GNO/tB9Qiv3zj7J5WPNPovHFLCmparpC6zrPiH+zLWRA6x2xVViB7Mx6ketAjus0ua4Tw54vtodWfSNQ1qzvlKiS2u1kXLr02tjjcMfrXbxSpKoeJldT0KnINAElFGaKACikzS0AFFFFABRRRQAUUUUAFIaWkxQAtFFFABSE0Vk+KNUuNE8NX+o2dqbua2hLrCP4iKAJtZ1a00jTJbq+vIrONF/1spwAfxrxLV/ipd+KbE2GlaKb6eGQ7NRDGOMHs69DXKX'
	$image &= 'p1Pxutv4i8SXLajDISy2MR/dwqeg2+ord8PwtD5yQQvDZgjyo24574HYdK8XF5mqcX7LVo9OhgXOzmY9t4f8R3UjPqniSWF3cSSR2zYOdxYHI9yfzqS/8L2VpB599faleMTgK9y+WP51qRaTcR6896sq+W/LepGAMH8q0ryzivrcxTA4zkFeCD65rw6mYV3NXqadbaHpQwlJRfu6+Zxdr4O0W7uzbXFnNBIy78rOxBHvzWrH4Khgx9i1bU7faPlEd0wx+tabWsWj2095Esk8oX7znJ9hVDTdRuZNRhH2n7SJRmVBysYo+tYmd5U6jsh+wpLSUSUX3i/w+6tZ+NJmAGEivj5gb8803wx8R9T8IeINS1TxHpXn2eouoeay+6jKMMwA7Egmm6xHGNU3+fbFpECiOc8r9KfYQXGj27wTQG8t2+YMgz1HIxXbRzGrCCcnzPz0+45qmDpydoqx6fB4rtPiOYrDwzeH7AVEl9MuVcKekYHUE4Oa1tf8Km50ezs9DWC2W1mWTymUFHA/veteFTrFp+pJrHha5Ok6rHzsI8tJx3Vh0NeoeEfFWsfEDw/LqKanb6QsDGK4ijXLRuOuSele9hsVHERulZ9jyq1GVJ2ep3ug6Sui6NBYiQymPO5yMZYkscDtya54eFdQfx5LqMrRfYTKs6sPvlggXb9OCfxrK8P+Prgz39uI7jWrS1AcX0aBcKeoPrjHWu/06+h1PToL22J8qdA6ZGDg11GBT17QYdes1glmlgZH3rLC2GHrzXm/x6mh8P8AwdOn2nyCR0iQA9QBzXr2KwPFXgzR/GVrb22vW/2iGCTzFTdgZ9/ypiPh66NvZaejpYXtrM6gwztIQD6kc19S/s3zzXHwrR7iV5W+1SAF23EDPTNdzqXgPw3q+l2mn6jpUE9tZjbAjoDsHt+VXtA8O6X4Y0wafodolpahiwjQYGT1NAGpS0gpaAEHWloooAKKKKAEzRS4ooABRRRQAYoFFFABRRRQ'
	$image &= 'AUyRQ6MrjKsCCPUU+kIzQB876tpTfD/x5PpkgKaPqkjT2Mh4VHPLR/rxSavdajb3MYskYxbMkhd2Wz09uMV7T4z8H6d400F9N1JSOd8Uy8NE/Zga8/sPgXJdNjxV4jvL+JPlSGFvKUqOmduCTXi4nK1Vr+1jbXe56VDG+zp8kvkZUcgkVcsu/HIBzg1He3iWNq08mSF4AHUmuj1T4D+H1smfw5LeabfoMxTLcOwz6EMSK4XSpp9TtbzSNfj26hp8hguQBjJHRh9RzXi4vLHhUqjd431PQw+NVZ8trMkGrXaNEbu2iWORgpVSdwzWukMcQIiQISOwxVCDR4YJkmnmkmMQ+QSHhff3qLSdP1r4i6pLZ+H5WsNHt32XOo7eXbuqZrno4d4qahR6bvob1a0aMeaf/BOfl1jRtPjmTUriPz2dhIudzn8OtTaP4005NMjiAuriRcgCOBjxk45+le2aB8JfCGgRqYtLjurjq1xdfvHY+uT0rrodOsoFCw2sMYHQKgGK+k/smlKNpts8h5hU+ykfOL+K4Z1w+iak6njm1bn9K0PhVfKvxPurO3066t9N1a0/0iOaAqvmL07dwa+hBBF/zzX8qBBErblRQ3qBXTh8BRw0uanf7znrYqpWjaRl/wDCPaTBokumJaRQ2Lqd8ajaMd81jeBlWGbUodMJbRopwtmScqPlG4Kf7u7P6110sSTRtHKoZGGCp6GmW9vDaQJDbRLHGgwqIMAV3nKTUVXub23tIjJczxxIBks7ADFcRrnxp8GaJvU6ot9IoOVs183n0JHSgDv6SvP/AIffFL/hYN9c/YdEu7WwhHy3c33Xb0H6V0XjTxH/AMIn4Rv9Y2LJJbR7o43PEjdloA36K4LwP8VdJ8Wqlpdg6XqwUF7O4+UtnuhPUd67sNmgB1FFFABRRRQAUUUUAFFFFABRRRmgAooooAKaTimXNxHa20k87hI41LMx7AV4vqvxh17WZLpfBWkRtZxlkS9um2+Y'
	$image &= 'RxlRWdSpCmuabsi4QlN2irnr0+r6fb3AguLyCOU8BWcZq3HIsi7kZWB6EHIr5bg0DQZFUeJbhLzV5jvnklmJYueSAfqa3tKuPFfhBmXwrqa3dgx3CyvnLhP91ieBXnRzSg58srrzex1PBVeW61Pogmvn7WTH/wALr8RfZCDGYYfNK9N+wVY1Dxl8RdbiNt5lho8TDDS22WfHt6VW0bRYdHgcI7zzzMXnuJTl5W7kmuPMswoSoOnB3bOnB4WqqinJWSKXimS6uVstF05yt1qk4gBB5C9zXvPhrw9Z+GPD9rpWnRrHFboFyBjc2OWPua8G1W8j0Txl4c128B+x2Vz++YfwA45r6KtrmK7t457aRZYpFDI6nIYetdOTxisKmutzHMJSdZpkwFKKQGlr2DzwpCaWua8c3fiS08OyN4Os47rUWYKokOAg7n3oA0Na8R6X4esWu9YvIrWFe7sOfpXmV/8AFHX/ABSzQeAdM8i1zj+1L5cL9VU9a5OD4V/EjWNVGp+JpLC6uicq95L5yxD0CZwPpWl4n8Ja34b0Z73xF49Wz2riCC1i8sM3ZQB2qJc/2QKGu6Jp9la/2n8Qtcv9alY/LE8zhC3okYP6Vf8AC/w0n8WNFe67YxaDoHDQ6bCojkuF7GQjn8DVFNI1Lx14R0HUJrr+ztQtWLmZo8lscbgPUgA/jWjJ4FnvPn1bxTrl4x5IN2wX8B2ohdkXS3PatMsLDS7GOz0qGGC2hG1I4QAB+Vc/8SvCVz4z8HyaXY3K284lSZC65VivRT7GvNI/BN5preb4e8U6xYzLyA9w0kZ+oJq/H42+IPhuBxqen2/iC3UHbPa/LLntla29nLcanFnNpBD4kvrjw/4m042OvaSqhrm2OCox8rK49eO9b+i/EXXvA8y2PjIvq2lKQqalGMyRDt5gHX61iafcv4Y8K6h4o19G/tDUZTcSqRypYnYnsAKgh8IeKtM0WHxku7V11RPNvtOT5/LjPKbR7A4P'
	$image &= '0raXLyrm3IV7ux9B6TrNjrmnx3ul3MdzbyDKuhzV8V458D9Dvba+1vWDYz6Zpl46La2kmVGRnc4XsDkflXsVc3U1FooooAKKKKAAUUCigApKKKAFFIaWg0AZ+t6cur6Jd6fJnZcxNGcHHUV8yrHqVno1/wCEhMtlrGnuY0LnbvjzkMCfUHrX1TiuX8VfDrw74wmjn1my33EYws0btG+PTKkZrlxOHVeKT3TujajVdJtrroeOfCjw5pfis6jpGseF7FrazUJLqCsXkll74f8Awp2raLefDbxJDplzNJc6DqB/0K4k5ML/APPMn8q910Dw7pnhnSk07RbVLa3ToqnJJ9STyT9a5/4saNBrHw21VJwA8ERnhfONjryDRXw0K9L2cx060qdTnicDS1n6Dcy3vh+xuZgQ8sKsc9zjrWgcjqK/PpR5ZOPY+ri+aKkRXNvDd27w3MYkikGGRuhrP0LXPEPw3k2acr6voBbcbQkmSAd9vt/hWpS114TGVcLK8NuxjiMPCurS3PRfCvxH8OeLIR/Z9+iXA4e2mOyRT6YNdWr7gCMH6V876n4X0vVJPOmg8u4H3Z4SUcH1yKLQ+M9C2rofimV4U+7DeKJB+ZFfTUc5w9Re/wC6/M8Wpl1WL93VH0TnNNYgDLEAe9eD/wDCb/FADb5ukt/tGMUn2Lxp4rUnxJ4leC0bhoLJBHu9sgV1vMML0nc5ZYerBXkrHr+v+KtM0Dw5daxdXMbW8APKtnc3ZfrXkWiaVc+KNWPjDxbH5l1P81jaSD5LWLovHqcZrOT4ctB4mtoobqd/DqYuJbSSUsGnHHf1wK6fxfcaja+FL2XRULXax/IFGSB7D6UpYlVLKHU5mP1nxVo+gr/xMr6NJDwIlOXJ7YFY0Pi/W9VXf4a8Ialew5wJnjKqTU/w6u/hlZQ20k9ys2vTBfPk1EEyeaeoG7pz6V7goA+6APpXbBcouVdTwyTVfHNn8974IuHj6k25LEfhUmleOtI1'
	$image &= 'G5+yzNJp170+z3S7G/DNe4gVgeJfBWheLLUw61p8U5x8suMOp9mHIrojVlETpxZw99bQXkDw3MaTRSDBVgCCDXLaTrWpfCu+zCJb/wALyvl4OWe056r/ALP+FWtc8L+JvhwrXOnyza74eQ/PC43TWy+xAyQKtW2oWes6alzZyLPbTrkemCOQff2rrXJXVupi3Km/I9d0jVrHW9Nhv9KuEubaddySRnINX6+ddL1W8+F+ufbbDfN4cupM3loOfs7H+NfQe3tX0DYX1tqVhDeWUqzW86h43U5DA1xTg4OzN4yUldFqiiioKEooooAWiiigAxRRRnFABRRRQAUUE4qnqWqWmkWMt5qE6QW0Qy8jnAWgC0TXl/xZ8Tm5tB4L0NhNquq4jlC4PkQk/MzelZd7478UeOpZIvBUQ0rSASp1OYfPKPVB/I1oeGfCFl4aSWUSSXuoXB3XF7cHdJIfqecV4WY5vRwsXCDvPt/mdFKhKTu9i/aaHZ2ul2tiIwyW0SxK3Q4AxUUvh+3bPluyGtUUtfnftZ3bvuevGco6JnOTeHZgMxOr+1Z8thcQkh4mGPbNdmaMA9efY1qsRJb6m0cRJbnDNBIoy0bD6imYK9Qc+4ruyikcqCPpUM1lbzspeJSQc5x1q1iV1RosT3RyNvavO2Exx1J6CrN54k0PRYhFe6jBEyj7u7LH8BWZ9gvviJ4qutI0OZ9L0XTZPKvbyAlXncdUUjt1r0DQ/hT4Q0NVaHR7e5mXrNcoJGJ9cmvrcBlUpwVWq7X6dTx8Zivay5VsjztviVo0j7dPt76+bt5NueaVPGWsXJxp/g3VJPQuMZr2+3sLS0GLW1hhHpHGF/lU4Ar6GnhKcNjztz5y8R6H4t8RWGy18CJZ3HmLKt0SBIpBr3HwprQ1fRYhNuS9t1WK7if7ySgDd+uas+INZi0DR5b+ZSyx8BR3J6c15/otvr3ifxTe6ja6jFYW5iWOR7NMq7em7+Ij1rpStoOx6oKW'
	$image &= 'uP07U77RfEA0bXL5bqOWLzoLlvlYckFW/L9a6yKVJUDRsGB6EGqEEkaSIySKGVhghhkH614N4q0M/Dfx/z+YntVKeHdakPyA/Laz55AHYHNe+dawfGnhqDxX4UvdKuFBMsZMbY5RxypHoc1UZOEroUkmrM8qu4UnheGZQ6OCGB7irfwb1uXRtcvfBd7IXhXNzpzMekZ+8n4HP51zfhi/mu9HNtfcX1hIbW4B6hl4z+IqrrFzJo3iTQ9et+JLW7VHI4yjHkGvRrxVWjzo46T5Z8rPpYGlqOJxJGrqQQwBBHSpK8w7QooooAKKSjNAATWT4j8RWHhnS2vtSkKoDtRFGWkY9FA7mtR3CKWYgADJJ7V5Hc6gnibxPd6/ekf2TpBaGxDfdZl4eT35GB9K48Zio4Wi6j17Lu+xpTg5ysdt4Q8axeKmu4WsLjT7q0ZfMgmOSFYZB/8ArVpan4gj0y5W3NpeXErruAggLjHuR0rx/wAM+MLrTrTUbzRLSG81LUJWu55JyRFaw/wKxBGSRzj3r1jwvrz634RtNYvoltWmi8yRf4V9+e3etKNXnVn8Wl12FKNvQ54fFS2juNRa+0bUbSw0/iW9kjOwtjkAY/WuFW7m+L+vNe6lI0Ph2zk/0XT84a4x/G49O4qDxT4usvHfiLyry8+zeFbCX/Vqfn1CVeuB1I7Vr3uv+HJ7aF/seo6f5ahYLtbJ4wgxxzjp9a8nM8VXs6OHi79ZJbG1KMV70jsESO2t1jijCxRLhUQcKB2Aqjb6/ptwxX7SInH8E/yN+RqpoOvx6ifs8txDcSBd0dxEflnXpn2PHIp/iLwvYeJYYF1Df/ozGSPYxX5u2cdcY6V8H7OMKnJXuvM9HmvG8TaFFUNFu5LzSYZbggyjcjkdyrFScfhV+uaUXGTi+ha1QUUUm9dwUkBj0XPJpALRmuY1jxXd2msz6fo+lNqLWcQluyH27ARnA9TjmtrSNVtta0mDULJswzDOD1Ug4IPu'
	$image &= 'CK2nh6kIKclo/wCvkTzJuxyyab4p8E6xfah4NaDULG+maefTrjI2uTklT2rYtfjZaWjeX4s0PUNHdRgyGIvGT7EDpW9TZY0mQrKiup6hhnNe9heIK9GKhUXMl95zzw0ZarQ2tE8Z+HvESZ0fV7W5OMlElG5fqO1bYYHpzXkerfDnw1qzmWSwFtcdRNbExsD68VVhsPHXhQ7/AA7r39r2inP2LU8E49A/B/CvocPn2EraSfK/P/M5ZYacdtT2SWGO4iaKZFdGGCrDINY/iGK9sfDcsfhy3AmGAqRgAqvcqOma5XQPi9Y3F8ul+LLKXw/qR423IIikP+y54r0NHWRAylWUjgg5Br3IyjNXi7o52mnZnnOh+CYjNda5r+nSTTeXthguD5suB1JP94k9Km+HV/NPrus266VLpVsrI8do7Z8olRnI7Z6/jXoX6VzmgaBNp/iPWtUnOBfyqUQtuIVVC/0qwOkFIRmlooEeBeOrFPCnxZkuCyxWOu2/mFicATKSD+mDXL+KdQs9S8Nu1hcxXH79FBibdhs17x44+H+meO105NWLbLKfzMIcbwcZX6cCuB1v4LyT/E6wvNIiitPDuI3uoYyF+dM9B78V0QruNPksZSpKUuY9c0RWTQNPWQEOLaMMD2O0Vodqao2gDpxS9q5zUM0UUUAFIadTW4HFAHHfEnWp9N8OrY6exGoapILWDHBG7q34CvL/ABgk4sNN8BeGcrcTIPtEgH+qiHVj7k5PvXW31w3iL4kXt2nz2uixi0t88qZmXc7fhuA/CsG5eLwJp2o67qzC71rUZWCKoyTyQka+wGK+RzLFc+LVOOrhsv7z6v0O+jC0LvqPXQogLHwL4e+XzAJNRuByVjGMkn1PpWt8addPhvwRaeG9CxFdakRaQIp5ROB/n6VyWl/8Jf4a0abxDeXkdneXlwsklv5YZ5Nx4RmPIAHQDGKr+ML+TX/ig2vybjp/h+6itQp5AYnLH8Ca9XLXRpUNJ3bd'
	$image &= 'm+8mYVuZy2PU/hz8NNN8G6LbtPCtzqhjBluZPmZSf4Vz0H0rt5FiEZEgXb3BFQ3OoW1nYNeTyqkCLuL54x2rkb26OqafPrXiC6k03QIlLJCHMZlUfxORg4PpXsaHOcf8QtL0jQb9PE/ha4gt7u0l33tlCcLMhPzHaOA3061JqvjuLfHZ+GbU6tfzIrKsZ+RNw+Xcf6U5/iF4B17wVrtt4djtY5ILVsxGBYzJ2BHHNdV8M/Aek+FfC9g8Nmv26SBZJppBl95AJwT0/D0FeXisroYurGpUW34m8K0oRaRiaf4T8e2OlpsvdJZ1yxtjG3JJJI3Z9+tFn4iubfVU0fxPYHTNQf8A1Rzuin/3G7/SvU9orm/Hfh2HxB4UuoGUC5hUzW0oHzRyLyCD2rHFZLhK8Hyx5X3Q4YicXqZ4615pPpN3rvinxFcRXU0Oq6ayGyKuQqpjO3b0IPvW1D41aPwrpFwbWS91O+HkrbocGSRflc+3IP51kaN4jjf4pbZYpLGe/txDc20o5WVScH3BB618dhcPXo+0klsn+D1+TszulKMralr4fawNb8Ra5esuySVYfMQ/wMEAYY9Mg1oeCQbTUfEGmdI7a+Z41/uhwHx/49WP4fsf+Ef+NGr2ajZBqdsLqIZ4z0b9Qa2/Do/4r7xUQfl86DH/AH5SrxSi/aOGzjFr8EKHS+9zq6KB0orxDoCiiikMz9Y0PTtfsXtNWtI7iJum8cg+x6g1zOka9qnwt1SKw1ieW/8AC1w4S3uZPmeyOcBWPUr06121V7+wttUsJbK+hWeCZSrowzn/AOvXrZdmdXBTVtY9V/kYVaMai8zt7eeK5t0mgdZI3AZWU5BFTY9K8f8AAGs3XgvxQfBWtTvLZXH7zSrmQ5+U/wDLMk+navYBX6TRrQrU1Upu6Z5Uk4uzCiiitiSOeVYIHlc4VFJJ9qwfB/ilPFdhdXUUBgWC6aAAnO4AAhvxBrU1q1kvtFu7WB9kssRVWPYmud+H'
	$image &= 'Xh2/8OaHcW+qRxxyy3G9UjbcAoRV6/8AAaCbu9jsKWkFFBQuKKSigArI8U6zHoHhi+1KX/lhExUZ5LY4A981r1n6zo1nrunmz1GISw7g+0+o6UpNpaDW+p554T099P8ADdutxzcz7ridiOWdyWP88VYn0Kxutag1S4i826t1Kw7zlUz3A6A+9bF3ZyWc5jccdj2xXLzXOo+KtWl0Tww/lRRHbfaj1WL/AGV/2vX0r8xp4bFYnFyglaTvfyPYcoRpp9DH8TavHe3lzNw+maCpnuXP3ZZ/4Y/8au+C/DiH4fi31aPzJNU3XF0W6ln5/TjmoviFpFnptt4d8B6Ku1Ly5E1yScs6rjLMe5JzzXWX0gstKneJQFhhO0egA4r1s2SwdGlg6T1vd+vT8fyOek3Uk6jOJsddvdHmisPEcdzq3h7TJCkF1AhYkjoJB329PwrV8feNfB/jHwTd6Kt/cLJOv7oRwEsGHTj09q2fDkQg8P2gAwWTe3uSc0zX72bT7KKW0CLLLcJESV7E81vQz6s5+w5FJ7J3t94nhk9bnhfhb4Lar/Zt5qeqF4GWItaW/wB1pCOQWHYe1dBrH7TOqWeo29tpOmwfZ4ERJ/PBDM4HzDrxzxXtDY5LdO5PpXkmr/DbTPEOuP4k0fTkuYxMRNZSuUS5IPLqRjGc/pWmW53KrUm8TpHSz6Imrh0kuU968La9H4m8L2GsQxtEt3EJNjfwnuPzqPxbrMGheFb6/uXCiOIgDP3mPAFcXZ+P9TsNNisbTwPqETwoI40U/uxjgc+lZl7HqGuX0Oq+Pru2srO3bfb6akmEVuzOT1PtXtYjM8NRhzKSb6Ja3+45o0pt2sYNnYTaPJ4Knvfl8x38zP8AA0h34+vzYrU+I/hpr6zg1/TU26rpLeYjKOXQdVq74vSHxD4PmudIuI7mW1YTxPE275lOccVuaPqUetaLa36bStzEGZR2OORXw08TUThibaptP77/AI3aPRUI6wRz81tJ'
	$image &= '4h1Xw34o0lFbYpE+W6Rtzj8Dmk8GE3GueJbw9Hv9gPrsUL/SukcW2k6XKYUS3ggjZwqjAXv/ADrC+HsDp4PiuphiS+lluW/4E5I/TFZOrzYefZWS+9v8BqPvI6gUUUV5psFFFFIYUlLRTEcv4+0R9X8ONPZ/LqGnt9ptXXghhyR+OK7rwH4kXxX4L0/VVx5kseJR/ddTtYfmKzCAeCODwa5/4Pu2ma14n8Ov8sdremaFB2SQBv619pw3iW1Og+mq/U4MVDaR6tRiiivsDhI7iVYLd5XOFRSxJ9q5vwN4mn8UaXd3NxCkRgu3hTYeGTAKt+IIre1K1N9ptxah9nmxlN3pkdaxfBPhlPC2g/Y1uTdSSSb5JexOAox+CigWtzpKKKKBhiiiigAoNFFAGZrmkDWtKms/Oe3aVSqzR/eTPpTdB0Kx8N6PDp+mxCOKJeSernuxPck85rUrG8W6xHoPhXUNSkIAghZh7nHFZqEVJyS1Y7u1jzDTZf8AhJPi/rutH57fTB9htm6gkZ3YrptZ/wCQHe4/54msL4babJYeC7aa5B+1agTeTt6tJzn8jWz4hsTqXh69tBLJEZImAaJ9rD6GvzbMMQsRj3JvRO3yR6tOPLTJ9L40e0x/zxX+VPurKG+SNbhN4jkEijPcdKzvCdqLPwpYQrLLMFj+9M5Zuvqag8Vay2iwWUyuVRp8OF6sMdK5adGdTFeypPVt2f3mu0dTW1G0a/sJbYTNF5o2l14OO9S21tFaWsdvbqEiiUKq+gFVNFW9/s8S6m5M8zGTZ/zyB5C/gKuTzxWtu887BIo1LMx6ACsZqcX7FO6T6dx6bmX4o8SWvhjRnvLrLuTshgX70rnoorldI8DzeJGGteP2e7uZhuisA5WK2XsMDHOKTw1bS+OfEjeLdUVjptsxi0q2ccHB5lI9Sc816Hn/APXXXOo8GvZU37/2n28l+v3GaXPq9jP0rQ9N0Oza10q0W2gdizRqSQT071z3hJjo'
	$image &= 'Wv6n4Ym4jVzdWWe8bdVH0IrsaoXGj2d1qlrqMqf6VagiKQEggHqD6iueFa6nGq78359CnHVNFfxPp15q/h+4sLCRY3uCEdm/ufxYrQsrVLHT4LWEYjhjVFx7Cp6KwdSTgodNyuVXuAooorMoKKKKACiiigA71y3gpivx48Sqv3Ws4CfrsFdT3rmvhqg1D4p+LtUT/VxtFaj6qi5r6XhtN4uT8v1Rx4r4EetDpRSUtfoB5pn69qMWlaHd3tw21Ioyfx7VyfwlvLq48Kzw3jlmtrt1GTkgMBJg/Qua7DVtLttZ0yawvk3wTLtcZxVfQtBsvD2n/ZNPjKoWLuzMSzse5J5JoFrc1OtL2pKKBhRRiigAooooAK8t+Nd5JdafpHhq2P73VrtVb3ReSK9RryHxQ39q/HbT7ZstFpdiZsD+F26fpXJjK3scPOp2TLprmmkdPHDHbwpDCu2OJQiKOygYApzKGBB6EYNLRX5Nrc9qxWsbKPT7QW8BYxqSQGOcZ5qMx2Gq7JCsdyLeQ7TjIVxxSavHeTaZJDp5/z9Vmk+Tex+6D1I96yJwmg3OkWkEjR28UcjS8434AJJ/HP5120KTq+9GXvu9rdkrv79kS3Y6XrXn3iu+m8X68vg/RnYWyYfVbpTwi/8APP6kVY0Wy8ZfEmzOoW2pRaDocrsIRCm6aVA2M7j0PvVODTdR+D1/cQ6lFJqHh+9n83+01XMkTHg+Zj6ZzXt0cmxOHpvEWTmtl+vy7HLKvGT5eh39paw2NpFa2sYjhhQJGgGNoHapqhtLy3vrVLmznSeGQZR0bIYVNXzEr3fNudgUUUVIBRRRQAUUUUDCiiigAooooEVdTvo9M0u6vpmCpbxM5J9hVf4J6dLD4HbVbpcXOrXMl057kFsD9AK5v4hTy6pNpvhGwJa51aUeaF/ghB5J+tew6ZYRaZptvZWy7YreNY1AHYDFfecOYVwoSrP7W3ojzsVO8uUtCiiivqTjCiiigAooooAK'
	$image &= 'KWigBKKKKAEPAOa8Z8Pv/aPxU8Y6mPnjSWO1Q+hRcEfnXsF5MILKaVuiRsx/KvHfhivnaLqWp9tS1O4nX/d8wgfyrwc/qcmBku7SOjDK9Q7Sobu5jsrSW5nOI41LE1NVW/0+DUoFhugzRhw20HGSK/O6ahzrn26nqvYZpV5Nf6dHdXMIhaQllTvt7E1z/wASrxLHwTdOqg3M+LaFscguccV1mB0AwOgFcPf2sHjT4nabpFtKZYNJcTXig/KD1APvXrZVQ+t46PKrLf0Xb9DGrLlps9U8IaSNE8HaVpwUA29rGjYHfaM/rWndWkN5bSW91EksMg2ujrkMKmUY4HSlr9LPIPHNZ8Daz4DvJNU8Bq93pL/PcaK5LbPUx56fStPw54r03xNa+ZYyFJk4mtpOJIm7givTiuetcH4x+F9jr13/AGros7aNrY6XlsNvmezgcN+NeJmOT0sZ70fdn3/zOilXcNyzR1riJtd8XeEAY/GGhNeWycf2jYDKkepXsa1NJ8e+HNZwtrqUaS94pflYH0xXxGIyzF4d+/DTutUehGtCWzOjopqSJIoKOrA91INO57ivOatuahRRRSGFFFH+c0wCqOs6xZ6Fpc2oahKEhiBPXlj2A96zfEPjXRvDse27uBNdHhLWH5pGPpgVT8P+CtY8catb6543hNrpsJ32ekE53ejSep9q9rLsprYuaclaHf8AyOerWjBWW5e+Fvh681HULrxt4giKXl98lnC3/LGAdOOxNepDjpTURY0CxgKqjCqOABT6/RqdONOChFWSPKbbd2JS0UVoIKKKKACiiigAooooAKKKKAOe8d6j/ZPgXV7z/nlauf0riPh/Zf2f8P8ARoDwTbiQ/V8t/WtP44Xb2/w2ngjPN5PHbkeoY81ga74stfCGlWWl2cRvtUFukcNpFyQFUDc3oK+X4ghUrqlh6Su22dmGtG8manizxDB4c0N7qaVUdzsj3HjJ71y2l+N7y4sUs/CmhajrkiDL'
	$image &= '3LqVRiepzjpW74K+HUfiNbfxR42uG1S5n/eQWrDEMA7AL3+terwW0NtCsUEaxIowqoMAflWuFyGjCioV9Xe7t/V9BTxMm/dPIV0b4oa2oDLp2hwSjBbl5EH0ruPBHgqx8GaWbWKQXV/MxluruQDzJnJySe+K6rGK5iASn4lXRIcRLZIAf4c5Ne3hsJQw6apRUTCU5T+JnUCigCiugzCkIzS0UAJiud1z4f8AhbxEGOq6LaSuxyZVjCOT/vDmujooA8ruPgfaWzNJ4c8Qatpb9UjSYtGv4Gqdx4L+I+jW5fT/ABHZ6mqKSVuYdpwBnrmvX8c1FeLmzmH/AEzb+VctTCUKvxwT+RanJbM8O0TXviLqejrqFvoNjfQl3jzFKQxKsVPGfUVJqHi/xxpOnzXmoeCJI4IF3SSeYQAK7j4Q/wDIjkdNt7cj/wAitWh8TR/xbPXcd7Vv5iuKWTYFu/s/zNPb1O55/ZX3xJ161STTvDdnYpKgdJbqU4IIznFXovhn4x1vafE3itrWFv8AWW2nJsP03V6P4X58J6V/16Rf+gitXFbUctwlF3hTVyZVpvdnJeFvht4c8KMJrCyE151a8uP3kpPruPT8K60ClxRXoWSMgooNFMAooooAMUUUUAFFFFABRRRQAUGikoA8i+Oseral/wAI9pOg2/2m6nvPMEeMqNuMFvYZq/ovwnNnpvm314bjV7+QPqV4wyzL18tP7q9uK7XxHqVtoti1+8Ky3QHl24C/MznooNXNIF6NKtzqjK12UBl2DADdx+FCjZ+06j1SLMEEdtbxwwIEjjUIijoAOAKlFApaBCGoFmge5eJGQzIBvUfeAPTNTGub0i0mTx1r1zJE6xP5Ijcjhv3a5x+OaaV7jOlooFFIQUUUUAFFFFABUdx81vIP9k1JTZBlWHsaAOF+EZx4Vu0/uandL/5GetL4mf8AJNNd/wCvU/zFZ/wqG3R9XT+5rF4P/IzVo/Ev/kmmuf8AXqf5igDT'
	$image &= '8Lf8ijpP/XnF/wCgCtasnwr/AMihpP8A15xf+gCtagAooooAKKKKACiiigAooo4oAKKOKOKACiiigBKDxSO4RSWIAHUntXjPxK+NcenNNovg10utQwVluwd0cHrj1NNJt2Q0m3ZHo1zrOgXni+20SZ4rjVYUNxHHgMYsdz6Guhr5l+BIkm+Lc9zcyvPcSWTtLM5yzknrX02KcouL5X0HJNOzCig1Q1fWdP0LT5L7VrqO1t4xlnkbAqSS6zBRknAHes/S9e03WZblNMu47k2snlTGM5Ct6V87ePvjFqXi9pNP8Os9hoxO1p1JEtwO/wBFrsP2blCeHdXRei3Zxk57CrdOSipPZlOLSue1ig0UhqCSpfatYaZC0uoXcNuijJMj4rkb74y+BbGUxvrtvKy9fJO/FeP/ABm8M6z4i+Mq2mk6bNfNLZIyjB8teSCSelTW/wCzhrY0j7TLq1nb3mwt9mSDKg/3d39adlbVj0PZNH+Kfg7XbkW+n63bmZjhI3bazH2rrwQRkV8Ny2P7ye3uoljurWVonZOMMpxkfiK+pPgt4gvPEPw6t5dRcyT2ztbmQ9XC9DWtSk4JS3TKlHlVz0GikFKaxIOE+Fx/0TXx6a3ef+jWrS+JfPw11z/r1b+lZXwtPyeJB6a3df8Aoxq1viT/AMk21z/r1agDS8K/8ihpH/XnF/6AK1qyPCn/ACJ+kf8AXnF/6AK16ACiiigANAFBoBoAKKKWgBKKKOtABRRRQAUUUUAZuv6JB4h0WfTbuWaKKcYZoJCjfmK+WviH4Di+HfiS006wvpbqzu4WlVZgNyckYyAK+tjXzv8AtDn/AIr3Rh6WBP8A4+1b4dtVVYun8SKnwCTf8Srx/wC7ZfzJr3fxf4li8I+F7vWbiFp0tl3eWhwW9q8O/Z8QnxvrMgH+rs0H6muR8deOvE+pa1reg6trLixS9kiFv5KqCgY7RnGemKKqcq0ku45K82e+/Dn4pW/xBuLyCLTZrGW1'
	$image &= 'VXIdshlbOP5Vv+JvBWg+MI4k8Q2Ru0h+4nnOij8FIr5J0fxFrXhnUJZfC9/Pb3MwVZY4oxJvA6ZBB9a+xtEnuLnQLGa8ObiS3RpeMfMVGf1qKlPklykyjys+PNW0+30rxRrGn2aeXb2148cabidqg8DJ5r2X9neZIPDuuyysESO6LMT2G0V5N4uG34ieIlHa+f8AkK9W/Z4jjn8Pa/DNzG9yVYH0K8101dcNA1l/DR6Afiv4JVST4hsz9Hplv8WvBd1dJbxa5Bvdgq5OAT9a8zvPAPwusLiSOXxVNHsYgxLKp289OlM0/wAPfCMatahNdmum80bYnlO2Rs8A4rlsrdTKyse/hELB8DdjhvavE/id8aNU0DxJe+HPD9hC8kMSiS6l52FlB4Hrg969rdxHC7k8KpOT24r4u17URqvjPX9QeQN515IFJPJVflH6CqowVSfKwhFN6kGh6bq3inVW0/RIpL/ULiQvPNj5EYn5mJ7da+uvAnhZPB3g+z0dG3vCmZX/ALznkn865D4A29qvw3hmgtoY5mmkWSVEAZ/mOMnvXqQGKVSblZPoEpNgKDRQazJPP/hYf3nihfTW7n/0M1s/Ej/km+t/9erVh/Czi98WL6a1P/6Ea3PiP/yTjW/+vVqBmh4T/wCRO0j/AK8ov/QBWvWP4S/5E3R/+vKH/wBAFbFAgoFFFABRQaKACiiigAoFFFABRRR2oATNFLiigANfOP7QrD/hYmkgn/mGZ/8AIj19HVyXi34a+HvGl9Fea3byPPDF5SPG5Uhck4/U1dOXJNS7Di7O55T+zmgbxP4hfqRBEP1Neg/FH4djxd4Ylg0O1soNTaZZfOaJVMmOxYDNavg34caJ4GmupdE8/ddbRJ5r7unTHFdbSnLmm5A3d3PKvhD8MtR8GSald+IfsstxdlRGIwG2Ae9epmlxS4qXq7iPjzxsuz4leIgeP9LJ/wDHRXqf7OQD6RryNyGugCM9QVroNb+BPh7XvEV5'
	$image &= 'q93fajHLdPvdIZFC5/75z2rqPBXgHSfAdlPbaMZ2W4ffI0z7iT+VbTqqVKNO2xo5JxUTj5v2fvDs97PcG/vk82VpNilcLk5xyM1d0j4HeG9J1i2v1mu7hrZw6RysCu4dDjFel0Vld7XIuynqlxFZ6VdXNwu6KKJndR/EAOlfHGoappmq6rNfad4X0+C2lkLCOR3LMM9TzxmvtCWJJYmjkUMjgqynoQa8j1n9nrQ7/VZbrTtQu9Pjmcu8Ee0qCfTI4qoOCfvocWk9Sx8DPFOnavoNzo+n6WumSacwMkcbFkbd3BNer1zPgzwLo3gbTGtNGjYtIQZp5Dl5T7mumxmo06E+gUGiigDzv4X8ax4xX01mWt74jc/DrW/+vVqwvhoNviHxkP8AqLuf0FbvxE/5J1rf/Xq1AF7wj/yJmj/9eUX/AKAK2KxvCH/Il6P/ANeUP/oArYoDqLRQBRQAUUYo60AFFFBoAKKKKACijrRigAooxRQAUYoooAKBSUUALSmkzQaAAdaKTNLQAUDrRRmgANGB6UdaOlAAaUUmc0ZoAO9FHek60Acz4V8M3Ohazr13PKkkepXn2iMKOVG0DB/KtDxTpMuueFdQ0y3dY5bqExqzdATWtxS0AUNCsZNM0CwsZWDPbW6RMw6EqMH+VX6KTPrQAooFJmloAO9KKSjNABRRn1pM0ALRSZpc0AJS5oooAM0UUUAFFFB9qAOc8Z+I5fD+lxiwjWbUbyVYLSE/xO39BW5a+ctrCLoq02weYVGAWxzgelct428J6n4h1DSbvRtTTTp7CVn81oRIcEAcA8Z461BP4L16ezkR/GOoSSMhUfu41Qk+yqD+tAEFh8RTP8R7vRLpIotOH7u2uy2N8q/eXn613fmKY96kFSMg54rzS5+CWiXPhOHTmeU38YRmvHkZ8y4G5tpPc5r0O1tPs+lxWhIOyIRkqu0HjHA7UgPKdO+J3iGXxPc2SW9nqcP29oUityRJFGDgsewHeu38'
	$image &= 'LeIrzxJq974A2qzIkY0m2mMABbMB80jLw58ATG7OMUzwp4MAP+EZ0fUbMSwAUrXc8sqyJFsASoboD64rF8MAnw+8S6PpqWMAJ4vlht4yxjgArS2jXaCSeSwACSeaBnTr4oQAfxTf6IlswksAO1W484t8rZwA8Y7dK4M/Fy8ATY6ak1pDa3MAfNJI88gYxRwASyMnGOrHb0wA/wA62brwNr8AbaoLvRddVpYA5tRa3k95EGkACASQy7cAHB8AQ1evfBN89poAfo+mal9h0e0A4BHOY4x9olMA3w/8II64GfcAoA6DQPEWmeIAXTje6NdLcwIAuUZl7MOorN8AGPi5fDtrFDYAsQudSuTiCEkAwAB1dvRRWxoAPo9joWmx2GkAdutvbx5IVe4Ae5Pqa5TXfCMArknjf/hIdBsAyzLSWwtngvoAEusY7lSGGPcApiK2jfFHS20APihub06rqBIARI2nwFo1bk4AOvQVR0b4p30A58La/YQC1vYA7eC1e0lLSKAANj5k/rxXW+EAfwoPD+l3UU0AOJ7m8kaWeRIAMIoYjGFUdAMAiuVsvhhN4LsAk6r4PkS5vG0AxuIb5Q3nEkkAyrDBU84446cAFIZqeN/Hk/gAX8QaXZ21uk8ADIQ98xPMMRMAgN+YNaNz4ygAofGGm6RHEksAbX9u0q3ayDYA5Bxt/KuZ0L4AGdrrFrqGp+IAuG6e91OV2MEAPM3+jp0VOvMAjscVPH8KFtMAwZaaTYai1vcAunzyTWd6sfIAm5icEfxcED8ACgRV1f4lajoAdoup6mgsQo0ASNpp6zMQJVUAO12yD6itXQ8AxJ4m1xkexk0AEuoUcCcwO7EAQdx160usfDQAt77wTpuh2k8AFDcadIkkV3IAwLJlh944P94ANP0fwj4i0W8AY5LbWrIwM4YAuIhp6p5g+q4AOaBnW6nc/YsATLm5yFMUTNkAI9BmvKY/i5oAhD4BklubdP8AAISNZdi2x40A6N8y'
	$image &= 'SY/u7CMA8a9Q13SV13QAK70ySV4RcxEAjMkfBXPpWHMAfD/S5dJjj8sAVtRhsBYx3zoAAuFCgAnH0pgAiLxF4wutI+EA7a6xaQx3GoUA1BE8UB6SOygAJ/nUkXjyyPwAPx4ilKeYsAMAJADysvQpj/cAsisa3+F897cA2kzeJNSa5g0ANslt0tYGeNMAeowJOG7gDNUALUfgxDEmqf8AAAjd3HYm+kgApEWSPeImTOQA5zk5zSGd/wAAhm51G88O2l0Aa0kcd5Mgd0gA8gLnoOfauZ8AEPj+XRPHtjoAUIYn011Vby4Ac/6iRidin6gAGfxqxB4R8QkAjU33jK+kcHIARFDEin2xtzgA/Gsqx+DumSYAlX0fiK5m1S8A7x2aS5eVsZ8A4flzjgYH4UAAjqvFeoX1r4QAry/0GW3NxDEAl0aQbkOOo4MAXJ2XjjxNpegAVjq/ifTbSXQA65hikM1pIVcAjDgHJQ5zjOMAilHw98QweHoAHw7p/iGKz0oAjtfJYpagyOwASSSSeO/pWvoAJ8NdJ06ysU0ARe41We0jVUYAu5S6KQAMqh4AB0oGUW8YW9gAad4m1aK/kb4AyT+WkV9hY0cA8tSFXvg5BqkA6B408T+I7SMAOnPoU1y0SyMAwI7lkB9eavYAo/D641aPxPYAl1cW4tdWmWcAtsxbmhkCBSQA56/dFQ6X4C8AEei+XJp2vWQAJ0VUZjpqLvAAOxK4P60AegUAv5n2dPP2+btARv29M98ewC4HAJghXzipk2jcAFRgE45xUlMQAGaKMUUAFGKMAFFABVXULk2WAJ81xHA07RIWABEn3mx2FWjXADXijVtcspbeANNB0RtRNyrBAKYybEh44yaAADC0n4t6fe2MADeappWoaZBKAMyCaaPMYIODAJYe/HSurg8TAGk3V5PaWt2kALPbwC4dEP8AAAEZB/HFcF4aAPBHiq48JR6TAK3qcWnWbvL5ANa28IZ2VnJxAL2z1z6V'
	$image &= 'He6FAKz4T8SXKeGdABpLy2vbCO0tAOZWyLYrkZcnAJxzmgeh6B4XAPEVr4p0GHVbAASSOGYsoEgwAH5WK/zFM8ReACvS/DFssmpSAJ3ucRwxjdJIAH2FO8K6Ivh3AMNWemIQTAnzAJHdiSSfzJriAK60XxbZ/ELUAHWIdJstV87aALZTz3BQW0YUAGVC445yc0COANrXxHbXfhw6AMrDOkIjMhjkAE2uMdsVhWPxAB4dR0+O8tNAANYkjlUMpW2HACCPrWfrnhzxAJa94b1CLUdSAIoGlgPlWdiNAKQ4z1k6kGsTAMLaR4k8Kw29AOJ4fvJHjtlhAHgGpNIp45IQAOcGgdjq5PifAKRb6/Npd9b3AFa+QkZlnljIAI42dQwVj2OCACty28T6ZeatAHenWs3mT2kCAFxJj7uxxlSDAN+K8213wP4hALvXtS8QymSfAEy+aN59EjfYAPKgjUHLAZzkABG3POKTUra6ANYmi1T4fwsjAH2YaZe2UmUaAATjacHuopBoAHVv8TbKLSrSAPpNOvWW8u5LAFgWJNzOVON2AD0OK1bDxb9uANQhtV0TVIBJAP8ALWWDai/UAOa4bxb8PdRSAE8M/wBjwz30ADp8YimgS6aBAHIAzICvOSc9AOtfQn8SaDYSAGlx6HckuryQANxNdG4Cv2UkAPOKAOhTxtprAPjuTwrh1vUgABNuP3Tntn1qAP8A/CRaadSvEPT47hUAkBLPCADqildwP5GvKAA/DXxdeeH21QDuNY8rxAZDdwAYiiVXVyANmwCx0wBjtWhLpgB4q0fxHeXFrgCXJfz6xYQQmwCuMQSLGEYvxwCxoA9F8M+IbQC8UaDb6rYo8QDBPnaJBg8HHQg/Co/AQdM8MwoAPqMjNJLxFBEALvkk+gqXw3oAMmgeHLLTI8cA+jwqhI6M2OQA/nmuJl0jxbYAXjrUtUh0ix0AUe4IFndXFyUAfs8QH3AuOOcAJyOuaYjrJ/EAjp1t4WXX'
	$image &= 'rpYAeG2fgRvHiTMAnGNvrmqeoeMI60sfIFtkJp1/AHVz5Cz/ALiIADAK3rz1qO08AD2taxcJL4zlALSe2QiSOxhQAIEcgPDbs5NcAF6z4E1r/hPdAFry3tb6+tryADQx3A1J4TGRANVAHbPbFAzuAGXx/YWem3V9AKrZXunQW+0AAN1FtMrHoqjuAGpZ/G9naeFkANburS6VGxm2AI03yrnoSo+lAHOaol7rPhJ9AD9b8HXF5JZmADW3ja4b96cYAN+8EHjHNXPAAF8Pf+EbvLjVAO/Yi+ulx9niAJ3eKFf7o3MdAMfc0BoSWnxUANL1Bpl07TNVALp4OJFjtuVOADODzXW6Vf8AAPaelwXn2ea2APOXd5U64dOeAIRXN+F9J1DTADxj4je5hxZXAFMktvNn73ygADD8xXT6hPLaAGnzzW9u1xLGAIWWFOrn0FAjAJLW/iIdE8YfANhPoOoXRMPnACzW6BgygDJAAPbNaen+N9J1AD0e7v7N5H+xAKF57criRPYiALjBY+Odf8bWAFrS6fb6Ii2TAEQec+cyEseqAPrV3WPhxdy+ABrWZV1K6utdANQjXfcRsIt2AMzhAF6LyfegIGa1n8RYMCWu7QB0DWZI5UDKRQC3XIz6111tLwDaLaOby2j8xAANscYK5HQ+9QDk/hjS/EfhWADgvE8PXkrJbgCxPbjU2lXgdQAI2QK9Yt5GlgDeOR42iZlBKADdV9jQIlxS0gBS0AJRRiigAgCKM0H2oADzSQCM0tFACYoxSwBBoAQDHHalogCM0AJiobu5hgDG1kubqRYoYgBdzuxwFHfNTwCa5v4hwz3PwwDdZgs4DPPJbADLHEoyXPpQBgCumavp+sWYugDTLuO5hYZDxgDZrl9b8Z2dmwBInhttOubxpAApOzyhBGw4+QC45rh/AttdWABrc51N4dB1iACtDDFpioUhlADt+WQ84Y/TFQCL4Tul066mlwBde3kkN9M19AAvo4fcd56P'
	$image &= '6ABHP40h2PX9MwDFtla+HbO58QAerWCz3JK+ZABNiNiD0B74qwCQeM/D13fQWQDa6pBNPcEiNABG6mvOfEFmfABL4n8Nt4UtrADTTGglKmWyDQAcbZwcr0BroQDVtJm8K/D3UQC+eGxm1WCJmQAubazWMoD6AQDPFAHa6bq1lgCtHLJp86zpFACtE5U9GHUVlwCv+LrDR/NtkgDiCTUUUMtq0gAFJyODXE+EPAB7oujeFrKx0gD0rVLgLGC8ggDYjzHPLHPfkwBcrrE9+fidqwDdapJFE7xRGwBvM0v7QDEVBAAM54Izg0wPVwDwdr1/rMl6dQAp7BjG4KQ2jwC/ylI/iPrW7AB6vYTai9hFdwAT3SDLQhssPgCivNNY8VWVvwCCZpPCNqILvwAyKC+ntrIRvABKRzIFHX2pngAe1LR9NjVvBQD4X1DUr8Id9wCXKbHcnqWY5ADz1pAeof2pZAAhnm+1ReXbNgDJW3cIRg4P5wBUJfEtrBqrWgDKkiwrbfaTdgBH7oL6Z9a8ngDDRfEuqeHfFgDFDNDZ2kl7MwBNBNCWZm2KTgAbPI7A+1V01ACu9Q0w29q1zQDJu/DKJDaAkgCO6Eoxx06igACx7Zpmr6frVgCLdaXdx3MLDgAeNs1exXinwwBLV7HxPA3iCQATQdQgtkhXTACJfLjnG0YckgBwTx2x1r2sHACMjpTEFGM0tAB0oATFGM0vWgAzigBMUdOlLQAZoAKUU0UtAAAtFJmigAooo+AUAFFFJdAA8HswAAALTTThSd6AMgD1rw5pev26xQCqWqTbCGV+jACn1BHIq/DbRwAMKxIo2qAPrQBLRQAyOGOFNgBCiovoowKJIgBJoykqB0bgqwAMg/hT6KAGJABJGgRFCqBgAAAwBS+WuclRnwBcU7vSmgCJbQDhR2dI1Vn+8QADBNP2ilooAgApYI5oXikUMgBICHH97NVdPwBF07S7aG3sLQAjhjgUpGAvKgAJJIB+pq/i'
	$image &= 'igAAy9a8OaX4ggDcQ6raJOFO5QJuEgmtCGJYIUgAkztRQq5OeBUOJaEMPwAxAFIKKKAgBaKKKAAwAAFFDCd6cTk0AI70UUBAAelFFFABPwBQCAFHamMCD//Z'
	$image = _WinAPI_Base64Decode($image)
	If @error Then Return SetError(1, 0, 0)
	Local $tSource = DllStructCreate('byte[' & BinaryLen($image) & ']')
	DllStructSetData($tSource, 1, $image)
	Local $tDecompress
	_WinAPI_LZNTDecompress($tSource, $tDecompress, 23947)
	If @error Then Return SetError(3, 0, 0)
	$tSource = 0
	Local Const $bString = Binary(DllStructGetData($tDecompress, 1))
	If $bSaveBinary Then
		Local Const $hFile = FileOpen($sSavePath & "\image.JPG", 18)
		If @error Then Return SetError(2, 0, $bString)
		FileWrite($hFile, $bString)
		FileClose($hFile)
	EndIf
	Return $bString
EndFunc   ;==>_image

func _Inferno()
   EndFunc