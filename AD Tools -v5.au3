#pragma compile(FileVersion, 5.0)
#pragma compile(ProductVersion, 5.0)
#pragma compile(FileDescription, [AD Tools v5.0])
#pragma compile(ProductName, [AD Tools v5.0])
#pragma compile(LegalCopyright, (28/01/2022) - Nicolas RISTOVSKI)
#pragma compile(CompanyName, Nicolas RISTOVSKI / EAPI69)

Global $lastdatecompile = "	(c) 2018~2022 " ;about box
Global $lastdateupdate = " 28-01-2022 " ;main routine information

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
;#include "FModMem.au3"
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
	Global $scomboreaddirectives = ""
	Global $restoredirectives = ""
	Global $actiongroups = ""
	Global $restoredirectives = ""
	Global $groupsidrh1
	Global $idrh1
    Global $domainname=""

	Func _terminate()
		Exit
	EndFunc

	Global $hdll = DllOpen("user32.dll")
	Global $historik
	Global $hguilg

	Func about()
;	     FSOUNDMEM_Init()
;Global $tMem
;Global $bChipSound = _Inferno()
;Global $iLen = BinaryLen($bChipSound)
;Global $tMem = DllStructCreate("byte[" & $iLen & "]")
;DllStructSetData($tMem, 1, $bChipSound)
;$bChipSound = 0
;Global $hFMod = FMUSICMEM_LoadSongEx($tMem, 0, $iLen)
;FMUSICMEM_SetMasterVolume($hFMod, 192)
;FMUSICMEM_PlaySong($hFMod)
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
		EndIf
		$defautdc = $isdct
		$defautdcinit = $defautdc
	EndIf
	If StringCompare($isdct, "dct.adt.local") = 0 Then
		$isdct = 1
	Else
		$isdct = 0
	 EndIf

	While 1
		Global $defautdc = $defautdcinit
		Global $allou = 0
		GUIDelete($hguilg)
		ToolTip("",5,5,"")
		_extended()
		$iterad = $iterad + 1
		If $iterad > 11 Then
			$iterad = 1
			$historik = ""
			$historik = $lastiter
			$lastiter = ""
		EndIf
	WEnd

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
		If $iwidth > @DesktopWidth Then $iwidth = @DesktopWidth - 100
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
		Global $idrh1 = "", $idrh2 = "", $idgroup = "", $idgroup2 = ""
		Global $actiongroups = ""
		Global $actiongroupsar = ""
		Global $driveismanual = ""
		Global $result, $result2, $outputidrh1
		GUIDelete($hgui)

		#Region GUI
			Global $hgui = GUICreate("    [ AD Tools ]---[ La POSTE/DSEM/EAPI69/NR ]---[ màj. " & $lastdateupdate & " ]         domain: [   " & $isdct2 & "   ]     |     iter11=" & $iterad, 1240, 595, 50, 50, $ws_sysmenu)
			GUISetBkColor(14745599)
			Global $aff = GUICtrlCreateEdit("", 350, 5, 880, 555, -1, 512)
			GUICtrlSetData($aff, $historik, 1)
			_ad_open()
			$buttonidrh1 = GUICtrlCreateInput("", 50, 50, 85, 35)
			GUICtrlSetFont(-1, 14, 400, 0, "MS Sans Serif")
			GUICtrlSetBkColor(-1, 6908265)
			GUICtrlSetTip($buttonidrh1, "Idrh Source.", "", 0, 1)
			$labelidrh1 = GUICtrlCreateLabel("Source", 10, 50, 40, 15)
			$idcheckboxcommonname = GUICtrlCreateCheckbox("", 15, 66, 14, 14)
			GUICtrlSetTip($idcheckboxcommonname, "searching by ANR (common name)" & @CRLF & "or with Idrh...", "", 0, 1)
			$buttonidrh2 = GUICtrlCreateInput("", 50, 100, 85, 35)
			GUICtrlSetFont(-1, 14, 400, 0, "MS Sans Serif")
			GUICtrlSetBkColor(-1, 6908265)
			GUICtrlSetTip($buttonidrh2, "Idrh destination, work with: (Idrh only !)", "", 0, 1)
			$labelidrh2 = GUICtrlCreateLabel("Dest.", 10, 100, 40, 15)
			$buttonidgroup = GUICtrlCreateInput("", 22, 300, 125, 35)
			GUICtrlSetFont(-1, 14, 400, 0, "MS Sans Serif")
			GUICtrlSetBkColor(-1, 6908265)
			GUICtrlSetTip($buttonidgroup, "Group or Object ? : check [Source] box , for using it ! " & @CRLF & "separator is [;]  " & "..works with: (Idrh, SID key,common name with a white space.. ex: nicolas ;san ;)" & @CRLF & @CRLF & "[Shift] key pressed => formating results to: 'SID key'|'objectname' , for DACL migration [B]", "", 0, 1)
			$labelidgroup = GUICtrlCreateLabel("Get SID key ?", 35, 280, 100, 15)
			$buttonidgroup2 = GUICtrlCreateInput("", 22, 400, 125, 35)
			GUICtrlSetFont(-1, 14, 400, 0, "MS Sans Serif")
			GUICtrlSetBkColor(-1, 6908265)
			GUICtrlSetTip($buttonidgroup2, "adding multiple existant Groups [;] to Idrh 'Source' " & @CRLF & "separator is [;]  " & "..works with: (Idrh) only" & @CRLF & @CRLF & "[Shift]: to remove groups", "", 0, 1)
			$labelidgroup2 = GUICtrlCreateLabel("adding groups [;]", 20, 380, 120, 20)
			$labelcreategroup = GUICtrlCreateLabel("Gr?", 115, 380, 20, 14)
			$idcheckboxcreategroups = GUICtrlCreateRadio("", 135, 380, 14, 14)
			GUICtrlSetTip($idcheckboxcreategroups, "Group members ? , get members of group" & @CRLF & @CRLF & "[SHIFT] key: create one group" & @CRLF & " With separator [;] Massive groups creation", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labeldrive = GUICtrlCreateLabel("cpy drives ?", 190, 20, 90, 12)
			$idcheckboxdrives = GUICtrlCreateCheckbox("", 174, 20, 14, 15)
			GUICtrlSetTip($idcheckboxdrives, "cpy all drives from Srce to Dest., work with: (Idrh only !)", "", 0, 1)
			If $isdct = 1 Then
				GUICtrlSetState(-1, $gui_checked)
			Else
				GUICtrlSetState(-1, $gui_unchecked)
			EndIf
			$labelscan = GUICtrlCreateLabel("cpy Mes scan ?", 190, 50, 90, 12)
			$idcheckscancpy = GUICtrlCreateCheckbox("", 174, 50, 14, 14)
			GUICtrlSetTip($idcheckscancpy, "cpy comment 'Mes scan' from Srce to Dest., work with: (Idrh only !)", "", 0, 1)
			If $isdct = 1 Then
				GUICtrlSetState(-1, $gui_checked)
			Else
				GUICtrlSetState(-1, $gui_unchecked)
			EndIf
			$labelgroups = GUICtrlCreateLabel("cpy Groups ?", 190, 80, 90, 12)
			$idcheckboxgroups = GUICtrlCreateCheckbox("", 174, 80, 14, 14)
			GUICtrlSetTip($idcheckboxgroups, "cpy all groups from Srce to Dest., work with: (Idrh only !)", "", 0, 1)
			If $isdct = 1 Then
				GUICtrlSetState(-1, $gui_unchecked)
			Else
				GUICtrlSetState(-1, $gui_checked)
			EndIf
			$labelreinitpwd = GUICtrlCreateLabel("Pwd Reset 'Srce' ?", 190, 110, 90, 12)
			$idcheckboxpwdreset = GUICtrlCreateRadio("", 174, 110, 14, 14)
			GUICtrlSetTip($idcheckboxpwdreset, "reseting password, work with: (Idrh only !)", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelairwatch = GUICtrlCreateLabel("Airwatch 'Srce' ?", 190, 140, 90, 12)
			$idcheckboxairwatch = GUICtrlCreateRadio("", 174, 140, 14, 14)
			GUICtrlSetTip($idcheckboxairwatch, "Adding Airwatch group, work with: (Idrh only !)", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelairwatchbpe = GUICtrlCreateLabel("AW ?", 293, 140, Default, 12)
			$idcheckboxbpe = GUICtrlCreateCheckbox("", 280, 140, 14, 14)
			GUICtrlSetTip($idcheckboxbpe, "AW ? = selectionne un autre groupe Airwatch (ex: BPE, Ma French Banque...)", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelairgroup = GUICtrlCreateLabel("Add group 'Srce' ?", 190, 170, 90, 12)
			$idcheckboxgroup = GUICtrlCreateRadio("", 174, 170, 14, 14)
			GUICtrlSetTip($idcheckboxgroup, "Adding group, work with: (Idrh only !)", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelmultiplesrce = GUICtrlCreateLabel("[ ; ]", 153, 153, 18, 16)
			$multiplesrce = GUICtrlCreateRadio("", 154, 170, 14, 14)
			GUICtrlSetTip($multiplesrce, "multiple Idrh 'Source' for adding groups   -  separator is [;]" & @CRLF & @CRLF & "[Shift] key pressed with button [OK ] => for removing groups..." & @CRLF & @CRLF & "work with: (Idrh, common name, SID key)", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelrole = GUICtrlCreateLabel("Add Role 'Srce' ?", 190, 200, 90, 12)
			$idcheckboxrole = GUICtrlCreateRadio("", 174, 200, 14, 14)
			GUICtrlSetTip($idcheckboxrole, "Adding role, work with: (Idrh only !)", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labellbpai = GUICtrlCreateLabel("LBPAI/Iliade ?", 40, 200, 90, 12)
			$idcheckboxlbpai = GUICtrlCreateRadio("", 24, 200, 14, 14)
			If $isdct = 1 Then
				GUICtrlSetTip($idcheckboxlbpai, "Add specifics [Category groups] for: LBPAI/Iliade -  [SHIFT] key for removing  , work with: (Idrh only !)" & @CRLF & "server: 752474SN-FI01\_Iliade$", "", 0, 1)
			Else
				GUICtrlSetTip($idcheckboxlbpai, "Add specifics [Category groups] for: LBPAI/Iliade , work with: (Idrh only !)", "", 0, 1)
			EndIf
			GUICtrlSetState(-1, $gui_unchecked)
			$labeldirective = GUICtrlCreateLabel("Directive 'Srce' ?", 40, 230, 90, 12)
			$idcheckboxdirective = GUICtrlCreateRadio("", 24, 230, 14, 14)
			GUICtrlSetTip($idcheckboxdirective, "Apply a Directive, work with: (Idrh only !)", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelairgroupremove = GUICtrlCreateLabel("remove group 'Srce' ?", 40, 170, 110, 12)
			$idcheckboxgroupremove = GUICtrlCreateRadio("", 24, 170, 14, 14)
			GUICtrlSetTip($idcheckboxgroupremove, "Removing group(s) or role(s), work with: (Idrh only !)", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelprolonge = GUICtrlCreateLabel("Extend date 'Srce' ?", 190, 230, 100, 12)
			$idcheckboxprolonge = GUICtrlCreateRadio("", 174, 230, 14, 14)
			GUICtrlSetTip($idcheckboxprolonge, "Extend date, work with: (Idrh only !)" & @CRLF & @CRLF & "[shift] key pressed wiyh button [OK]: no expiration for 'x' accounts only !", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelcreate = GUICtrlCreateLabel("Create 'Srce' ?", 190, 260, 90, 12)
			$idcheckboxcreate = GUICtrlCreateRadio("", 174, 260, 14, 14)
			GUICtrlSetTip($idcheckboxcreate, "Create new user, work with: (Idrh only !)" & @CRLF & @CRLF & "[SHIFT] key held for renaming User Source", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelmove = GUICtrlCreateLabel("Moving OU 'Srce' ?", 190, 290, 90, 12)
			$idcheckboxmove = GUICtrlCreateRadio("", 174, 290, 14, 14)
			GUICtrlSetTip($idcheckboxmove, "Move a user to another OU, work with: (Idrh only !)" & @CRLF & @CRLF & "checkbox with no Idrh / work with text_file for massive move" & @CRLF & " ex: pabc123|BPLY" & @CRLF & @CRLF & "Massive LBPAI/Iliade category" & @CRLF & " ex: pabc123|MIPS|Standard;commun full users;Chargé de Développement" & @CRLF & "	lbpai? or iliade? to see available category from [ANR] Srce", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labeldrive = GUICtrlCreateLabel("add Drive 'Srce' ?", 190, 320, 90, 12)
			$idcheckboxdrive = GUICtrlCreateRadio("", 174, 320, 14, 14)
			GUICtrlSetTip($idcheckboxdrive, "Adding a new drive from user, work with: (Idrh only !)" & @CRLF & @CRLF & "multiple idrh 'srce' with separator ';'  for list of users 'srce' drives / DCT only", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labeldriveremove = GUICtrlCreateLabel("remove Drive 'Srce' ?", 190, 350, 110, 12)
			$idcheckboxdriveremove = GUICtrlCreateRadio("", 174, 350, 14, 14)
			GUICtrlSetTip($idcheckboxdriveremove, "Remove a drive from user, work with: (Idrh only !)", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labeldesactivesrce = GUICtrlCreateLabel("Disable 'Srce' ?", 190, 380, 90, 12)
			$idcheckboxdesactivesrce = GUICtrlCreateRadio("", 174, 380, 14, 14)
			GUICtrlSetTip($idcheckboxdesactivesrce, "Disable a user, work with: (Idrh only !)" & @CRLF & @CRLF & "[SHIFT] key held for deleting account", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labeldescription = GUICtrlCreateLabel("Description 'Srce' ?", 190, 410, 90, 12)
			$idcheckboxdescription = GUICtrlCreateRadio("", 174, 410, 14, 14)
			GUICtrlSetTip($idcheckboxdescription, "change description for user Srce, work with: (Idrh only !)", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			$labelscan = GUICtrlCreateLabel("Comment 'Srce' ?", 190, 440, 75, 12)
			$idcheckboxscan = GUICtrlCreateRadio("", 174, 440, 14, 14)
			GUICtrlSetTip($idcheckboxscan, "change 'Comment' Mes scan for user Srce, work with: (Idrh only !)", "", 0, 1)
			GUICtrlSetState(-1, $gui_unchecked)
			Global $idok2 = GUICtrlCreateButton("OK", 55, 480, 225, 65)
			GUICtrlSetTip($idok2, "Press OK to valid", "", 0, 1)
			Global $idabout = GUICtrlCreateButton("About", 1, 1, 35, 20)
			Global $buttonlocalgroup = GUICtrlCreateButton("Local Groups", 274, 1, 75, 20)
			GUICtrlSetTip($buttonlocalgroup, "Manage Local Group" & @CRLF & "compare,list,add or remove locals groups...", "", 0, 1)
			Global $buttoncomputer = GUICtrlCreateButton("Computer", 105, 1, 60, 20)
			GUICtrlSetTip($buttoncomputer, "Add/remove 'Group' to Computer", "", 0, 1)
			Global $buttonbitlocker = GUICtrlCreateButton("Bitlocker", 41, 1, 60, 20)
			GUICtrlSetTip($buttonbitlocker, "'Bitlocker for Computer'", "", 0, 1)
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

Local $idmsg
		While 1
			$idmsg = GUIGetMsg()
			Switch $idmsg
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
					$t = MsgBox(4, "Computer ?", "[Oui]: Ajoute/enlève 'Groupe' => Computer" & @CRLF & @CRLF & "[Non]: ne fait rien")
					If $t = 6 Then
						computergroup()
					EndIf
				Case $idok2
					$idcheckboxcreate = BitAND(GUICtrlRead($idcheckboxcreate), $gui_checked)
					$idcheckboxprolonge = BitAND(GUICtrlRead($idcheckboxprolonge), $gui_checked)
					$idcheckboxmove = BitAND(GUICtrlRead($idcheckboxmove), $gui_checked)
					$idcheckboxdrive = BitAND(GUICtrlRead($idcheckboxdrive), $gui_checked)
					$userexist = _ad_objectexists($idrh1)
					$idcheckboxdirective = BitAND(GUICtrlRead($idcheckboxdirective), $gui_checked)
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
					If $idcheckboxmove = 1 AND $isdct = 1 AND StringLen($idrh1) = 0 Then
						massive_move()
						Return 0
					EndIf
					If $idcheckboxprolonge = 1 AND StringLen($idrh1) = 0 Then
						massive_date()
						Return 0
					EndIf
					If $idcheckboxdrive = 1 AND $isdct = 1 AND StringInStr($idrh1, ";") Then
						massive_drive()
						Return 0
					EndIf
					If StringCompare($idrh1, $idrh2) = 0 AND $idrh1 <> "" Then
						MsgBox(0, "Info !", "User 'Srce' is same as User 'Dest' !" & @CRLF & "Aborting...", 7)
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
											MsgBox(0, "Warning !", $idgroup2 & " , lenth<7" & @CRLF & "aborting...", 7)
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
											Global $groupdesc = InputBox("Set 'Description' group !", $idgroup2 & @CRLF & "Please describe it, if necessary...", "")
											$ivalue = _ad_modifyattribute($idgroup2, "description", $groupdesc, 2)
											MsgBox(64, "Info !", "Group '" & $idgroup2 & "' successfully created" & @CRLF & $sou & @CRLF & @CRLF & "description is set to:" & @CRLF & "'" & $groupdesc & "'")
											GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Group '" & $idgroup2 & "' successfully created to: " & "  " & $sou & "    -   " & "description is set to:" & "  " & "'" & $groupdesc & "'" & @CRLF, 1)
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Group '" & $idgroup2 & "' successfully created to: " & "  " & $sou & "    -   " & "description is set to:" & "  " & "'" & $groupdesc & "'" & @CRLF
										ElseIf @error = 1 Then
											MsgBox(64, "Info !", "Group '" & $idgroup2 & "' already exists in AD !")
										ElseIf @error = 2 Then
											MsgBox(64, "Warning !", "this OU= '" & $sou & "' does not exist in AD !")
										Else
											MsgBox(64, "Warning !", "Return code '" & @error & "' from AD..." & @CRLF & " (access denied ? or groupname lenth is too long !) for creating group")
										EndIf
										ToolTip("", 5, 5)
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
												GUIDelete($hgui)
												MsgBox(0, "Warning !", $idgroup[$z] & " , lenth<7" & @CRLF & "nothing done... aborting !", 7)
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
												$result = $result & "Group '" & $idgroup2[$z] & "' successfully created	" & $sou & "	/ description is set to:	" & "'" & $idgroup2 & "'" & @CRLF
											ElseIf @error = 1 Then
												$result = $result & "Group '" & $idgroup2[$z] & "' already exists in AD !" & @CRLF
											ElseIf @error = 2 Then
												$result = $result & "Warning !  " & "this OU= '" & $sou & "' does not exist in AD !" & @CRLF
											Else
												$result = $result & "Warning !  " & "Return code '" & @error & "' from AD... " & " (access denied ? or groupname lenth is too long !) for creating group: '" & $idgroup2[$z] & "'" & @CRLF
											EndIf
										Next
										ToolTip("", 5, 5)
										$result = "		Creation des nouveaux groupes en masse dans:  " & $sou & @CRLF & @CRLF & $result
										$result2 = StringTrimRight($result2, 1)
										$result = $result & @CRLF & @CRLF & $result2
										GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $result & @CRLF, 1)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $result & @CRLF
										Return 0
									Else
										MsgBox(0, "Warning !", "new group to create is Not Empty !", 7)
										ToolTip("", 5, 5)
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
										GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idgroup2 & @CRLF & $amembersaddsam & @CRLF, 1)
										If $amembersaddsam <> "" Then
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idgroup2 & @CRLF & "Le groupe contient [ " & $win[0] & " ] membres." & @CRLF & $amembersaddsam & @CRLF & @CRLF
										Else
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idgroup2 & @CRLF & "Le groupe n'a aucun membre ! [ " & $win[0] & " ]." & @CRLF & $amembersaddsam & @CRLF & @CRLF
										EndIf
									Else
										$win = "Le Groupe:  " & $idgroup2 & "  n'existe pas !"
										GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idgroup2 & @CRLF & $win & @CRLF, 1)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $idgroup2 & @CRLF & $win & @CRLF
									EndIf
								EndIf
								Return 0
						    Case $idcheckboxcommonname = 1
	  if $domainname="" Then
		   MsgBox(0,"Info !","no Active Directory found ! unable to scan..." & @crlf & $domainname,15)
		   ToolTip("",5,5,"")
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
									MsgBox(64, "Info !", "If user 'Srce' is set to 'disabled' or 'locked', then it will search for thoses users in AD..." & @CRLF & @CRLF & "And Also... 'SID' will search for all category (group,person)," & @CRLF & "'SIDh' will search for all SID history for all category (group,person...)" & @CRLF & "'lbpai?' or 'iliade?' to see all groups categorys for massive import...with section 'move user' and no idrh 'srce' !")
									ToolTip("", 5, 5)
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
											InputBox("SID key !", "SID key for: " & $displayname & @CRLF & $idgroup & @CRLF & $testbinaryh, $sidkeyg1)
											ClipPut($sidkeyg1)
										Else
											$result2 = "SID key for: " & $displayname & "  " & $idgroup & "     " & $sidkeyg1 & @CRLF & "SIDhistory : " & $testbinaryh & @CRLF
											ClipPut($result2)
											GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $result2 & @CRLF & @CRLF & "-- SID keys:" & @CRLF & $sidkeyg1 & @CRLF & "--" & @CRLF & @CRLF & "			=====	=====	=====	=====	=====			" & @CRLF & @CRLF, 1)
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $result2 & @CRLF & @CRLF & "-- SID keys:" & @CRLF & $sidkeyg1 & @CRLF & "--" & @CRLF
											ClipPut($sidkeyg1)
										EndIf
										ToolTip("", 5, 5)
										Return 0
									ElseIf StringInStr($idgroup, "-") AND StringLen($idgroup) >= 20 AND $idrh1 = "" Then
										$sidkeymatch = _ad_getobjectsinou("", "(&(objectSID=" & $idgroup & "))", 2, "sAMAccountName")
										If @error > 0 AND NOT IsArray($idgroup) Then
											MsgBox(64, "Info !", "SID key not found, err: " & @error & @CRLF & "  search for SID key:  " & $idgroup)
											$historik = $historik & "SID key not found, err: " & @error & @CRLF & "  search for SID key:  " & $idgroup
											ToolTip("", 5, 5)
											Return 0
										ElseIf NOT IsArray($idgroup) Then
											ToolTip("", 5, 5)
											InputBox("search SID key", "SAM Account Name:  " & $displayname & @CRLF & $sidkeymatch[1], $idgroup)
											ToolTip("", 5, 5)
											Return 0
										EndIf
										ToolTip("", 5, 5)
										Return 0
									EndIf
								Else
									ToolTip("", 5, 5)
								EndIf
								If StringInStr($idgroup, ";") AND $idgroup <> "" AND $idrh1 = "" Then
									$idgroup = StringSplit($idgroup, ";")
									If NOT IsArray($idgroup) Then
										MsgBox(0, "Warning ! '7841'", "no SIDkeys or Objects in this Matrix !", 7)
										ToolTip("", 5, 5)
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
														$result2 = $result2 & "SIDkey found:  " & $sidkeyg1 & "  for userID:  " & $idgroup[$z] & "  " & $displayname & $testbinaryh & @CRLF
													EndIf
													ToolTip("", 5, 5)
												ElseIf NOT IsArray($sidkeyg) Then
													$result2 = $result2 & "SIDkey not found for UserID:  " & $idgroup[$z] & @CRLF
												EndIf
											EndIf
										EndIf
										If $idgroup[$z] <> "" AND $idrh1 = "" AND (StringRegExp($idgroup[$z], "^S-\d-\d+-(\d+-){1,14}\d+$") OR StringRegExp($idgroup[$z], "^S-\d-(\d+-){1,14}\d+$")) Then
											$sidkeymatch = _ad_getobjectsinou("", "(&(objectSID=" & $idgroup[$z] & "))", 2, "sAMAccountName")
											If @error > 0 Then
												$result2 = $result2 & "Account  'Idrh' not found for SIDkey :  " & $idgroup[$z] & @CRLF
												ToolTip("", 5, 5)
											Else
												ToolTip("", 5, 5)
												$result2 = $result2 & "Account  'Idrh' found for SIDkey :  " & $sidkeymatch[1] & "  " & $idgroup[$z] & @CRLF
												ToolTip("", 5, 5)
											EndIf
											ToolTip("", 5, 5)
										EndIf
										ToolTip("", 5, 5)
									Next
									ToolTip("", 5, 5)
									If $result2 <> "" Then
										ClipPut($result2)
										GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $result2 & @CRLF & @CRLF & "			=====	=====	=====	=====	=====			" & @CRLF & @CRLF, 1)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $result2
									 EndIf
									 ToolTip("", 5, 5)
									Return 0
								EndIf
								$idrh1 = GUICtrlRead($buttonidrh1)
								If StringInStr($idrh1, "locked") Then
									ToolTip("Searching for locked users in AD...", 5, 5)
									$alocked = _ad_getobjectslocked()
									If @error > 0 Then
										MsgBox(64, "Info !", "No locked user accounts have been found, err: " & @error)
									Else
										ToolTip("", 5, 5)
										__arraydisplay($alocked, "Locked Users Accounts")
									EndIf
									ToolTip("", 5, 5)
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "searching for locked users accounts"
									Return 0
								EndIf
								$idrh1 = GUICtrlRead($buttonidrh1)
								If StringInStr($idrh1, "disabled") Then
									ToolTip("Searching for disabled users in AD...", 5, 5)
									$adisabled = _ad_getobjectsdisabled()
									If @error > 0 Then
										MsgBox(64, "Info !", "No disabled user accounts have been found, err: " & @error)
									Else
										ToolTip("", 5, 5)
										__arraydisplay($adisabled, "Disabled Users Accounts")
									EndIf
									ToolTip("", 5, 5)
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "searching for disabled users accounts"
									Return 0
								EndIf
								$idrh1 = GUICtrlRead($buttonidrh1)
								If StringInStr(($idrh1), "SID") AND NOT StringInStr($idrh1, "SIDh") AND StringLen($idrh1) = 3 Then
									ToolTip("Searching for SID in AD... (all category: group,person etc...)", 5, 5)
									$sid = _ad_getobjectsinou("", "(&(objectSID=*))", 2, "sAMAccountName,cn,displayName,objectSid,sidhistory")
									If @error > 0 Then
										MsgBox(64, "Info !", "No SID have been found, err: " & @error)
									Else
										ToolTip("", 5, 5)
										ClipPut(_arraytostring($sid, ";"))
										__arraydisplay($sid, "SID Accounts")
									EndIf
									ToolTip("", 5, 5)
									Return 0
								EndIf
								$idrh1 = GUICtrlRead($buttonidrh1)
								If StringInStr(($idrh1), "SIDh") AND StringLen($idrh1) = 4 Then
									ToolTip("Searching for SID history in AD...(all category: group,person etc...)", 5, 5)
									$sidh = _ad_getobjectsinou("", "(&(SIDhistory=*))", 2, "sAMAccountName,cn,displayName,objectSid,sidhistory")
									If @error > 0 Then
										MsgBox(64, "Info !", "No SID history have been found, err: " & @error)
									Else
										ToolTip("", 5, 5)
										ClipPut(_arraytostring($sidh, ";"))
										__arraydisplay($sidh, "SID history Accounts")
									EndIf
									ToolTip("", 5, 5)
									Return 0
								EndIf
								$idrh1 = GUICtrlRead($buttonidrh1)
								If StringInStr($idrh1, "iliade?") OR StringInStr($idrh1, "lbpai?") Then
									$groupslbpailiade = "Chargé de Développement" & @CRLF & "Direction des Ressources Humaines" & @CRLF & "Direction adminstrative et financière" & @CRLF & "Direction Technique" & @CRLF & "Direction juridique et risque" & @CRLF & "Direction des systèmes d'information" & @CRLF & "Direction Marketing et Commerciale" & @CRLF & "Standard" & @CRLF & "commun full users" & @CRLF & "Direction générale" & @CRLF & @CRLF & "Autres categories:" & @CRLF & "----------------------------" & @CRLF & "BDD_DAF" & @CRLF & "Bibliotheque_DP" & @CRLF & "Comité de Direction" & @CRLF & "Communication Interne" & @CRLF & "Contrats" & @CRLF & "DFRC" & @CRLF & "Direction Financiere" & @CRLF & "Direction Systèmes Informatiques" & @CRLF & "DRA" & @CRLF & "Echange_DAF-DT" & @CRLF & "Echange_DM-DT" & @CRLF & "Echange_DRA-DT" & @CRLF & "Gourvernance" & @CRLF & "Risques et Conformité" & @CRLF & "Services Generaux" & @CRLF & "Supports de présentation"
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & @CRLF & "		LBPAI/Iliade [Categories]" & @CRLF & @CRLF & $groupslbpailiade & @CRLF & @CRLF
									Return 0
								EndIf
								ToolTip("searching by common name : " & $idrh1 & " please wait... 'Ambiguous Name Resolution'", 5, 5)
								Global $sou = ""
								Global $sldapfilter = "(&(objectCategory=Person)(ANR=" & $idrh1 & " ))"
								If $filtreperso = 1 Then
									$sdatatoretrieve = InputBox("data to retrieve", "Ex: cn,sAMAccountName, initials,mail..", "Displayname,cn,SAMaccountName")
								Else
									Global $sdatatoretrieve = "sAMAccountName,cn,DisplayName,initials,title,mail,telephoneNumber,Mobile,physicalDeliveryOfficeName,comment,Laposte00-NetworkDrive"
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
									Else
										Global $sdatatoretrieve = "sAMAccountName,cn,DisplayName,initials,title,mail,telephoneNumber,Mobile,physicalDeliveryOfficeName,comment"
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
									MsgBox(64, "Warning !", "No users found with filter= " & $idrh1 & @CRLF & @CRLF & "  ... Now trying in ObjectClass= Groups  !", 20)
									Local $vargroup = _ad_getobjectsinou("", "(&(objectClass=group)(name=*" & $idrh1 & "*))", 2, "samaccountname")
									If IsArray($vargroup) Then
										$hwnd = GUICreate("List of groups matching filter= " & $idrh1, 1000, 60)
										$combo = GUICtrlCreateCombo("", 10, 10, 980, 60)
										For $d = 1 To UBound($vargroup) - 1
											GUICtrlSetData($combo, $vargroup[$d] & "  ==>  " & "[" & _ad_getobjectattribute($vargroup[$d], "distinguishedName") & "]", $vargroup[1] & "  ==>  " & "[" & _ad_getobjectattribute($vargroup[$d], "distinguishedName") & "]")
										Next
										GUISetState()
										Do
										Until GUIGetMsg() = $gui_event_close
										GUIDelete()
										ClipPut(_arraytostring($vargroup))
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & " searching for common name 'ANR'" & @CRLF & _arraytostring($vargroup)
										ToolTip("", 5, 5)
										Return 0
									 Else
										ToolTip("", 5, 5,"")
										MsgBox(64, "Warning !", "No group found with filter= " & $idrh1)
									 EndIf
									 ToolTip("", 5, 5,"")
									Return 0
								EndIf
								Global $adatatoretrieve = StringSplit($sdatatoretrieve, ",", $str_nocount)
								For $i = 0 To UBound($auserids, 2) - 1
									$auserids[0][$i] = $adatatoretrieve[$i]
								Next
								ToolTip("", 5, 5)
								Return 0
							Case $multiplesrce = 1
								$defautdc = $defautdcinit
								If _ispressed("10", $hdll) Then
									$test = 1
									ToolTip("multiple Idrh selection and removing selected groups...", 5, 5)
								Else
									$test = 0
									ToolTip("multiple Idrh selection and adding selected groups...", 5, 5)
								EndIf
								$multipleidrh = GUICtrlRead($buttonidrh1)
								If $multipleidrh = "" Then
									ToolTip("", 5, 5)
									MsgBox(0, "Info !", "aborting ! no users defined by ';' in 'Source' ... " & $multipleidrh, 7)
									Return 0
								EndIf
								If $isdct = 1 Then
									$unite = InputBox("default OU ? 2 or 4 chr$", "ex: MILY, MITE, BPLY, GAUB, MI ... virtuos" & @CRLF & "?? : scan all OUs...", "MI")
									If StringLen($unite) < 2 OR StringLen($unite) > 4 AND NOT ($unite = "virtuos" OR $unite = "??") Then
										ToolTip("", 5, 5)
										MsgBox(0, "Info !", "aborting ! no default OU defined... " & $unite, 7)
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
									MsgBox(64, $adroot, "No Groups found ! ")
									ToolTip("", 5, 5)
									Return 0
								Else
									$filterou = InputBox("filtering groups ?", "empty = all groups , or type: cbk, dpi ... ", "")
									Global $aselected[1]
									If $test = 1 Then
										Global $hgui2 = GUICreate("Select groups to remove  and [ Apply ] !  (" & $filterou & ")", 1000, 500)
									Else
										Global $hgui2 = GUICreate("Select groups for adding and [ Apply ] !  (" & $filterou & ")", 1000, 500)
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
									MsgBox(0, "Info !", "aborting, no groups found in: " & $unite, 7)
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
													$result = "		Adding   multiple users to selected groups ...	" & @CRLF
												Else
													$result = "		removing multiple users to selected groups ...	" & @CRLF
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
																	$result = $result & "Target OU not found !"
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
																	$result = $result & "Target OU not found !"
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
												GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $result & @CRLF & @CRLF & "			=====	=====	=====	=====	=====			" & @CRLF & @CRLF, 1)
												$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $result
												ExitLoop
											Else
												MsgBox(4160, "warning !", "nothing selected... try again or close window")
											EndIf
										Case $gui_event_close
											$idcheckboxgroup = 2
											GUIDelete($hgui2)
											ExitLoop
									EndSwitch
								WEnd
								ToolTip("", 5, 5)
								Return 0
							Case StringUpper($idrh1) = "" OR $userexist = 0 AND $multiplesrce = 0 AND $idcheckboxcreate = 0
								$defautdc = $defautdcinit
								MsgBox(0, "warning !", "User Idrh Srce '" & $idrh1 & "' = <empty> or bad user Idrh !", 7)
								Return 0
							Case $idcheckboxcreate = 1 AND $userexist = 0 OR ($idcheckboxcreate = 1 AND _ispressed("10", $hdll))
								$defautdc = $defautdcinit
								If _ispressed("10", $hdll) Then
									If _ad_objectexists($idrh1) = 0 Then
										MsgBox(0, "Warning !", "Id: " & $idrh1 & " not found...", 15)
										Return 0
									EndIf
									Global $ireply = MsgBox(308, "Rename AD Account ?", "Do you want to rename  '" & $idrh1 & "'  ?")
									If $ireply <> 6 Then
										Return 0
									EndIf
									Global $prenomidrh1 = ""
									Global $nomidrh1 = ""
									Global $samidrh1 = ""
									Global $formr = GUICreate("Renaming user:  [ " & $idrh1 & " ]  -  " & _ad_getobjectattribute($idrh1, "givenName") & " " & _ad_getobjectattribute($idrh1, "sn"), 814, 180)
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
												GUIDelete($formr)
												Return 0
											Case $bok
												$samidrh1 = GUICtrlRead($iobject)
												$prenomidrh1 = GUICtrlRead($inewnamegn)
												$nomidrh1 = GUICtrlRead($inewnamesn)
												$descriptionidrh1 = GUICtrlRead($inewdescription)
												GUIDelete($formr)
												$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "   renommer un compte utilisateur AD:  " & $idrh1 & "  commande validée !"
												ExitLoop
										EndSwitch
									WEnd
									If _ad_objectexists($samidrh1) = 0 OR ($samidrh1 = $idrh1) Then
										Global $ivalue = _ad_renameobject($idrh1, $prenomidrh1 & " " & $nomidrh1)
										If $ivalue = 1 Then
											MsgBox(64, "Renaming user...", "User [" & $idrh1 & "] renamed to..." & @CRLF & @CRLF & "[" & $samidrh1 & "] -  " & $prenomidrh1 & " " & $nomidrh1 & "  -  '" & $descriptionidrh1 & "'" & @CRLF & @CRLF & "initially was:" & @CRLF & @CRLF & "[" & StringUpper($idrh1) & "] -  " & $initialprenom & " " & $initialnom & "  -  '" & $initialdesc & "'")
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "   Renaming user..." & "  User [" & $idrh1 & "] renamed to..." & @CRLF & @CRLF & "[" & $samidrh1 & "] -  " & $prenomidrh1 & " " & $nomidrh1 & "  -  '" & $descriptionidrh1 & "'" & @CRLF & @CRLF & "compte initial:" & @CRLF & @CRLF & "[" & StringUpper($idrh1) & "] -  " & $initialprenom & " " & $initialnom & "  -  '" & $initialdesc & "'"
										ElseIf @error = 1 Then
											MsgBox(64, "Renaming user...", "Object '" & $samidrh1 & "' does not exist")
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "   utilisateur  " & $samidrh1 & "  inexistant !"
										Else
											MsgBox(64, "Renaming user...", "Return code '" & @error & "' from Active Directory")
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "   code erreur: " & @error & "  accès refusé ?!"
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
										MsgBox(0, "Warning !", "'" & $idrh1 & "'" & " can't be renamed to '" & $samidrh1 & "' !" & @CRLF & "This ID is already used ...", 15)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "   Warning !  " & "  '" & $idrh1 & "'" & " can't be renamed to '" & $samidrh1 & "' !" & "   " & "This ID is already used !"
									EndIf
									Return 0
								Else
									If $isdct = 1 Then
										$unite = InputBox("default OU ?", "ex: MILY, MITE, BPLY, GAUB ... " & @CRLF & "?? : scan all OUs...", "")
										If @error Then
											MsgBox(0, "Info !", "Aborting... create user !")
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  annulation creation nouvel utilisateur !"
											Return 0
										EndIf
										If StringLen($unite) < 2 OR StringLen($unite) > 4 AND ($unite <> "??") Then
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
										$passwordtoset = InputBox("initial Pwd for users: Cnah", "Initial Pwd set for: " & $idrh1, "Cnah" & @YEAR)
									ElseIf StringUpper($domainname) = "COURRIER" Then
										$passwordtoset = InputBox("initial Password", "Initial Pwd set for: " & $idrh1, "Abcd1234*")
									Else
										$passwordtoset = InputBox("initial Password", "Initial Pwd set for: " & $idrh1, "W!nd0ws10")
									EndIf
									If $passwordtoset = "" Then
										$passwordtoset = "W!nd0ws10"
									EndIf
									Global $ivalue = _ad_createuser($sou, $idrh1, $prenomidrh1 & " " & $nomidrh1)
									If $ivalue = 1 Then
										MsgBox(64, "", "User '" & $idrh1 & "' in OU '" & $sou & "' successfully created" & @CRLF & @CRLF & "mot de passe initial (à changer): " & $passwordtoset)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & "' in OU '" & $sou & "' à été crée, " & "le mot de passe initial (à changer): " & $passwordtoset & @CRLF & @CRLF
										If StringUpper($domainname) = "COURRIER" Then
											$oudep = InputBox("code regate", " code regate à 6 chiffres...", "")
											_ad_addusertogroup("GLD-Z-UtilEtab-Etab" & $oudep, $idrh1)
											_ad_addusertogroup("GGD-Z-DansEtab-Etab" & $oudep, $idrh1)
										EndIf
										If $isdct = 1 Then
											$historik = $historik & "Autopilote, groupe rajouté: rg-pitr_cm_dir-tertiaire_standard_std" & @CRLF & @CRLF
										EndIf
										ToolTip("working...", 5, 5, "")
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
										Else
											_ad_setpassword($idrh1, $passwordtoset, 1)
											$pwdreset = @CRLF & @CRLF & "Compte Srce: " & $idrh1 & " , reinit. mot de passe à:  " & $passwordtoset & "  (Le mot de passe devra etre modifié à l'ouverture de session !)"
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $pwdreset
										EndIf
									ElseIf @error = 1 Then
										MsgBox(64, "", "User '" & $idrh1 & "' already exists")
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & " , existe deja !"
									ElseIf @error = 2 Then
										MsgBox(64, "", "OU '" & $sou & "' does not exist")
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & " , n'existe pas !"
									ElseIf @error = 3 Then
										MsgBox(64, "", "Value for CN (e.g. Lastname Firstname) is missing")
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & " , le nom ou le prenom n'a pas été saisi !"
									ElseIf @error = 4 Then
										MsgBox(64, "", "Value for $sAD_User is missing")
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & " , manque une valeur !"
									Else
										MsgBox(64, "", "Return code '" & @error & "' from Active Directory")
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & " , code erreur " & @error & " !"
									EndIf
									$userexist = _ad_objectexists($idrh1)
									If $userexist = 0 Then
										MsgBox(0, "Warning !", "La creation à échouée car ce  Nom/prénom  existe sous un autre 'Idrh'" & @CRLF & "...ou droits dans cette OU !", 15)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  La création à echouée car l'utilisateur existe deja sous un autre 'idrh' !"
										Global $rechercheanrc = ""
										$rechercheanrc = InputBox("Recherche par 'ANR'", "Creation du compte échoué !" & @CRLF & "recherche par 'Nom' ... ", "")
										Global $sou = ""
										Global $sldapfilter = "(&(objectCategory=Person)(ANR=" & $rechercheanrc & " ))"
										Global $sdatatoretrieve = "sAMAccountName,cn,displayName,title,distinguishedName,comment,Laposte00-NetworkDrive"
										Global $auserids = _ad_getobjectsinou($sou, $sldapfilter, 2, $sdatatoretrieve, "displayName")
										If @error Then
											ToolTip("", 5, 5)
											MsgBox(64, "Warning !", "No users found ! " & $rechercheanrc)
											Return 0
										EndIf
										Global $adatatoretrieve = StringSplit($sdatatoretrieve, ",", $str_nocount)
										For $i = 0 To UBound($auserids, 2) - 1
											$auserids[0][$i] = $adatatoretrieve[$i]
										Next
										ToolTip("", 5, 5)
										ClipPut(_arraytostring($auserids, ";"))
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & " noms trouvés par recherche ANR... " & @CRLF & _arraytostring($auserids, ";")
										__arraydisplay($auserids, "common usernames found ! for keyworld,  '" & $rechercheanrc & "'")
										Return 0
									EndIf
								EndIf
								If $isdct = 1 Then
									$t = MsgBox(4, "Apply Directive ?", "Rajouter une directive après la création du compte  " & $idrh1 & "  ?")
									If $t = 6 Then
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  Rajouter une 'Directive' pour  " & $idrh1 & @CRLF & @CRLF
										directives()
									Else
									EndIf
								EndIf
							Case $idcheckboxcreate = 1 AND $userexist = 1
								$defautdc = $defautdcinit
								MsgBox(0, "warning !", " Idrh Srce [" & $idrh1 & "] already exists in AD !", 7)
								$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  L'idrh " & $idrh1 & "  existe deja !"
								Return 0
								$userexist = _ad_objectexists($idrh1)
							Case $idcheckboxprolonge = 1 AND StringInStr(StringLeft($idrh1, 1), "x")
								$defautdc = $defautdcinit
								Global $date_expire = ""
								$date_expire = _ad_getobjectproperties($idrh1, "accountExpires")
								$date_expire = $date_expire[1][1]
								If $date_expire = "0000/00/00 00:00:00" OR $date_expire = "1601/01/01 00:00:00" Then
									MsgBox(0, "warning !", "Le compte '" & $idrh1 & "' , n'expire jamais...", 7)
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "warning !  , " & "Le compte '" & $idrh1 & "' , n'expire jamais..."
								EndIf
								If $date_expire = "0000/00/00 00:00:00" OR $date_expire = "1601/01/01 00:00:00" Then
									$t = MsgBox(4, "compte AD en 'x' ", "Le compte " & $idrh1 & " est un compte en 'x' , souhaitez-vous mettre une date d'expiration ? (Oui)" & @CRLF & "   ou (NON), le compte n'expire jamais")
									If $t = 7 Then
										MsgBox(0, "no expiration date !", "[" & $idrh1 & "] : no expire date ...", 7)
										$date_expire = "01/01/1970 00:00:00"
										$ivalue = _ad_setaccountexpire($idrh1, $date_expire)
									Else
										$date_expire = _dateadd("d", -1, $date_expire)
										$hguicalendar = GUICreate("compte " & $idrh1, 300, 100, 350, 120)
										Local $iddate = GUICtrlCreateDate($date_expire, 30, 30, 100, 20, $dts_shortdateformat)
										Local $idbtn = GUICtrlCreateButton("OK", 120, 65, 60, 20)
										GUISetState(@SW_SHOW)
										While 1
											Switch GUIGetMsg()
												Case $gui_event_close
													GUIDelete($hguicalendar)
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
										MsgBox(0, "no expiration date !", "[" & $idrh1 & "] : no expire date ...", 7)
										$date_expire = "01/01/1970 00:00:00"
										$ivalue = _ad_setaccountexpire($idrh1, $date_expire)
									Else
										$date_expire = _dateadd("d", -1, $date_expire)
										$hguicalendar = GUICreate("compte " & $idrh1, 300, 100, 350, 120)
										Local $iddate = GUICtrlCreateDate($date_expire, 30, 30, 100, 20, $dts_shortdateformat)
										Local $idbtn = GUICtrlCreateButton("OK", 120, 65, 60, 20)
										GUISetState(@SW_SHOW)
										While 1
											Switch GUIGetMsg()
												Case $gui_event_close
													GUIDelete($hguicalendar)
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
									$ou = InputBox("default OU ?", "ex: MILY, MITE, BPLY, GAUB ... " & @CRLF & @CRLF & "   " & $idrh1 & " , actual OU:  [ " & $ousourcedir & " ]" & @CRLF & @CRLF & "?? : scan all OUs...", "")
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
									MsgBox(64, "Active Directory ", $idrh1 & "' successfully moved to '" & $ou & "'")
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & " déplacé vers " & $ou
									ToolTip("synchronising AD", 5, 5)
									Sleep(3000)
									ToolTip("", 5, 5)
								ElseIf @error = 1 Then
									MsgBox(64, "Active Directory ", "Target OU '" & $ou & "' does not exist")
								ElseIf @error = 2 Then
									MsgBox(64, "Active Directory ", "Object '" & $idrh1 & "' does not exist")
								ElseIf @error = 3 Then
									MsgBox(64, "Active Directory ", "Object '" & $idrh1 & "' already in OU '" & $ou & "'")
								ElseIf @error >= 4 OR @error = -2147352567 Then
									MsgBox(64, "Active Directory ", "You have no permissions to move user " & $idrh1)
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & "non deplacé vers " & $ou & ",  accès refusé !"
								Else
									MsgBox(0, "warning !", "User Idrh Srce '" & $idrh1 & "already exists in AD !", 7)
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
											$lettre = StringMid($chaine, $k, 1)
											$test = StringRegExp($drivesidrh1[$j], $lettre & ";\\", 3)
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
									If $driveismanual = "" Then
										MsgBox(0, "Info !", "Empty drive ! aborting...", 7)
										Return 0
									EndIf
									$manual2 = $driveismanual
									If StringRegExp($manual2, "Mes scan;") AND $presencemesscan = 1 Then
										MsgBox(0, "Warning !", "'Mes scan' already used ! aborting..." & @CRLF & "adding: " & $driveismanual & @CRLF & "dispo: " & $chaine3, 7)
										Return 0
									EndIf
									For $k = 1 To StringLen($chaine5)
										$chaine4 = StringMid($chaine5, $k, 1)
										$chaine4 = $chaine4 & ";"
										If StringInStr(StringUpper(StringLeft($manual2, 2)), $chaine4) Then
											MsgBox(0, "Warning !", "Drive letter already used ! aborting..." & @CRLF & $driveismanual & @CRLF & "dispo: " & $chaine3, 7)
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
									_arraydisplay($drivesidrh1, "LaPoste00-NetworkDrive modified for: [" & $idrh1 & "]")
									If $t = 7 Then
										ExitLoop
									Else
										$t = MsgBox(256 + 32 + 4, "Add new Drive ?", "Souhaitez-vous rajouter un autre lecteur ?")
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
									MsgBox(0, "info !", "[user does not exist]", 7)
								Else
									MsgBox(0, "info !", "[Return error code " & @error & "] from Active Directory", 7)
								EndIf
								ToolTip("synchronising AD", 5, 5)
								Sleep(3000)
								ToolTip("", 5, 5)
							Case $idcheckboxgroupremove = 1 AND StringLen($idrh1) <> 0 AND $userexist = 1
								$defautdc = $defautdcinit
								Global $groupsidrh1 = _ad_getusergroups($idrh1)
								If @error > 0 Then
									MsgBox(0, "Info !", "user " & $idrh1 & " n'a aucun groupe ...", 7)
									Return 0
								EndIf
								_arraysort($groupsidrh1, 0, 1)
								Global $aselected[1]
								Global $hgui2 = GUICreate("Select/unselect groups to REMOVE from '" & $idrh1 & "' , then validate with the button [ Apply ] !  ", 1000, 500)
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
									MsgBox(0, "Info !", "aborting, no Groups found in: " & $idrh1, 7)
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
															$actiongroups = $actiongroups & $sitems[$z] & " [target OU ?] ; "
														ElseIf @error = 3 AND $test = 1 Then
															$actiongroups = $actiongroups & $sitems[$z] & " [pas membre !] ; "
														ElseIf @error = 3 AND $test = 0 Then
															$actiongroups = $actiongroups & $sitems[$z] & " [déjà membre !] ; "
														ElseIf @error >= 4 OR @error = -2147352567 Then
															$actiongroups = $actiongroups & $sitems[$z] & " [accès refusé !] ; "
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
												MsgBox(4160, "warning !", "nothing selected... try again or close window")
											EndIf
										Case $gui_event_close
											GUIDelete($hgui2)
											Return 0
											ExitLoop
									EndSwitch
								WEnd
								GUIDelete($hgui2)
							Case $idcheckboxdesactivesrce = 1 AND StringLen($idrh1) <> 0 AND $userexist = 1
								$defautdc = $defautdcinit
								If _ispressed("10", $hdll) Then
									$t = MsgBox(4, "Supprimer le compte ?", "Do you want to delete account [ " & $idrh1 & " ] ?")
									If $t = 6 Then
										Global $ivalue = _ad_deleteobject($idrh1, _ad_getobjectclass($idrh1))
										If $ivalue = 1 Then
											MsgBox(64, "Info !", "User  '" & $idrh1 & "'  successfully deleted !")
											Return 0
										ElseIf @error = 1 Then
											MsgBox(64, "warning !", "User '" & $idrh1 & "' does not exist !")
										Else
											MsgBox(64, "warning !", "Return code '" & @error & "' from AD !")
										EndIf
									Else
										MsgBox(0, "Deleting Aborted !", "Aborting Deleting account :  " & $idrh1, 7)
									EndIf
								Else
									Global $ivalue = _ad_disableobject($idrh1)
									If $ivalue = 1 Then
										MsgBox(64, "Info !", "User  '" & $idrh1 & "'  successfully disabled !")
										Return 0
									ElseIf @error = 1 Then
										MsgBox(64, "warning !", "User '" & $idrh1 & "' does not exist !")
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
									Return 0
								EndIf
								If $descriptsrce <> "" Then
									_ad_modifyattribute($idrh1, "description", $descriptsrce, 2)
									MsgBox(0, "Info !", "Source " & $idrh1 & ", Description set to:" & @CRLF & @CRLF & $descriptsrce)
									Return 0
								Else
									_ad_modifyattribute($idrh1, "description", $descriptsrce, 2)
									MsgBox(0, "Info !", "Source " & $idrh1 & ", Description is now cleared !")
									Return 0
								EndIf
								Return 0
							Case $idcheckboxdriveremove = 1 AND StringLen($idrh1) <> 0 AND $userexist = 1
								$defautdc = $defautdcinit
								Global $drivesidrh1 = _ad_getobjectproperties($idrh1, "LaPoste00-NetworkDrive")
								_arraydelete($drivesidrh1, 0)
								If UBound($drivesidrh1) = 0 Then
									MsgBox(0, "Info !", "aborting, no Drives found for: " & $idrh1, 7)
									Return 0
								EndIf
								$drivesidrh1 = _arraytostring($drivesidrh1, "|")
								$drivesidrh1 = StringReplace($drivesidrh1, "LaPoste00-NetworkDrive|", "")
								$drivesidrh1 = StringSplit($drivesidrh1, @CRLF)
								Global $aselected[1]
								Global $hgui2 = GUICreate("Select drives to REMOVE from '" & $idrh1 & "' , then validate with the button [ Apply ] !  ", 1000, 500)
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
									MsgBox(0, "Info !", "aborting, no Drives found for: " & $idrh1, 7)
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
												$actiongroups = " Removed drives: " & $actiongroups
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
											Return 0
											ExitLoop
									EndSwitch
								WEnd
							Case $idcheckboxscan = 1 AND StringLen($idrh1) <> 0 AND $userexist = 1
								$defautdc = $defautdcinit
								$descriptsrce = _ad_getobjectattribute($idrh1, "Comment")
								$descriptsrce = InputBox("Comment " & $idrh1, " ", $descriptsrce)
								If @error = 1 Then
									MsgBox(0, "Info !", " ...aborting !" & @CRLF & @CRLF & $descriptsrce)
									Return 0
								Else
									_ad_modifyattribute($idrh1, "Comment", $descriptsrce, 2)
								EndIf
							Case $idcheckboxdirective = 1 AND StringLen($idrh1) <> 0 AND $userexist = 1
								$defautdc = $defautdcinit
								directives()
							Case $idcheckboxlbpai = 1 AND StringLen($idrh1) <> 0 AND $userexist = 1
								$defautdc = $defautdcinit
								lbpai_illiade()
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
													ElseIf @error = 1 Then
														$idgroup2result = $idgroup2result & "Group '" & $idgroup2[$z] & "' n'existe pas !" & @CRLF
														$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  :	" & $idgroup2[$z] & "	groupe n'existe pas !"
													ElseIf @error = 2 Then
														MsgBox(64, "", "Idrh source '" & $idrh1 & "' inexistant !")
													ElseIf @error = 3 Then
														$idgroup2result = $idgroup2result & $idrh1 & " n'est deja pas membre de '" & $idgroup2[$z] & "'" & @CRLF
														$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "	:  " & $idrh1 & ",  n'est deja pas membre de :  " & $idgroup2[$z]
													Else
														$idgroup2result = $idgroup2result & $idrh1 & " [X] '" & $idgroup2[$z] & "' " & "Return code '" & @error & "' from Active Directory" & @CRLF
														$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "	:	code erreur:  " & @error & " !  (accès refusé...)"
													EndIf
												Case $actiongrp = 1
													If $ivalue = 1 AND $actiongrp = 1 Then
														$idgroup2result = $idgroup2result & $idrh1 & " [+] OK,   '" & $idgroup2[$z] & "'" & " ;  [" & _ad_getobjectattribute($idgroup2[$z], "distinguishedName") & "]" & @CRLF
														$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  :" & $idrh1 & "	  '" & $idgroup2[$z] & "'" & "[+] OK  ;  [" & _ad_getobjectattribute($idgroup2[$z], "distinguishedName") & "]" & @CRLF
													ElseIf @error = 1 Then
														$idgroup2result = $idgroup2result & "Group '" & $idgroup2[$z] & "' n'existe pas !" & @CRLF
														$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  :	" & $idgroup2[$z] & "	groupe n'existe pas !"
													ElseIf @error = 2 Then
														MsgBox(64, "", "Idrh source '" & $idrh1 & "' inexistant !")
													ElseIf @error = 3 Then
														$idgroup2result = $idgroup2result & $idrh1 & " est deja membre de :  '" & $idgroup2[$z] & "'" & @CRLF
														$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & ",  est deja membre de :  " & $idgroup2[$z]
													Else
														$idgroup2result = $idgroup2result & $idrh1 & " [X] '" & $idgroup2[$z] & "' " & "Return code '" & @error & "' from Active Directory" & @CRLF
														$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  :  code erreur:  " & @error & " !  (accès refusé...)"
													EndIf
											EndSelect
										EndIf
									Next
									$actiongroups = $idgroup2result
									If $idgroup2result <> "" Then
										If $actiongrp = 0 Then
											$actiongroups = "Removing Multiple Groups from: " & $idrh1 & @CRLF & "--------------------------" & @CRLF & $actiongroups
										Else
											$actiongroups = "Adding Multiple Groups to: " & $idrh1 & @CRLF & "------------------------" & @CRLF & $actiongroups
										EndIf
									EndIf
								Else
									MsgBox(0, "Info !", "aborting, no Groups to add to: " & $idrh1, 7)
								EndIf
							 EndSelect

					#EndRegion Select Case

If $idcheckboxprolonge = 1 AND NOT StringInStr(StringLeft($idrh1, 1), "x") Then
						$t = MsgBox(4, "Expiration date ? ", "Le compte [" & $idrh1 & "] n'est pas un compte en 'x' , souhaitez-vous mettre une date d'expiration ?" & @CRLF & "     ex:  pour les CF -> oui !")
						If $t = 6 Then
							Global $date_expire = ""
							$date_expire = _ad_getobjectproperties($idrh1, "accountExpires")
							$date_expire = $date_expire[1][1]
							If $date_expire = "0000/00/00 00:00:00" OR $date_expire = "1601/01/01 00:00:00" Then
								$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "warning !  , " & "Le compte '" & $idrh1 & "' , n'expire jamais..."
							EndIf
							$date_expire = _dateadd("d", -1, $date_expire)
							$hguicalendar = GUICreate("compte " & $idrh1, 300, 100, 350, 120)
							Local $iddate = GUICtrlCreateDate($date_expire, 30, 30, 100, 20, $dts_shortdateformat)
							Local $idbtn = GUICtrlCreateButton("OK", 120, 65, 60, 20)
							GUISetState(@SW_SHOW)
							While 1
								Switch GUIGetMsg()
									Case $gui_event_close
										GUIDelete($hguicalendar)
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
							MsgBox(0, "no expiration date !", "[" & $idrh1 & "] : no expire date ...", 7)
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
									MsgBox(0, "Info !", "Airwatch ...aborting !" & @CRLF & @CRLF & $mail)
									Return 0
								EndIf
								_ad_modifyattribute($idrh1, "mail", $mail, 2)
								If $isdct = 0 Then
									MsgBox(0, "Airwatch DCT ?", "AW is only for domain DCT" & @CRLF & "  mail is set to: " & $mail & @CRLF & "  for user: " & $idrh1)
								Else
									Global $ousourcedir = _ad_getobjectattribute($idrh1, "distinguishedName")
									$ousourcedir = StringSplit($ousourcedir, ",")
									$ousourcedir = $ousourcedir[3]
									$ousourcedir = StringTrimLeft($ousourcedir, 3)
									If StringInStr($ousourcedir, "GAUB") = 1 Then
										$bpnn = InputBox("OU ? (BPLY, BPGR, etc...)", "Actual user OU= " & $ousourcedir & " (for: " & $idrh1 & ")" & @CRLF & "by default is BPNA, you can change it !", "BPNA")
									Else
										$bpnn = InputBox("OU ? (BPLY, BPGR, etc...)", "Actual user OU= " & $ousourcedir & " (for: " & $idrh1 & ")", $ousourcedir)
									EndIf
									If @error = 1 Then
										MsgBox(0, "Info !", "Airwatch ...aborting !" & @CRLF & @CRLF & $mail)
										Return 0
									EndIf
									If $bpe = 0 Then
										_ad_addusertogroup("RG-" & $bpnn & "_AW Commun Manager", $idrh1)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  rajout @mail et groupe Airwatch, pour: " & $idrh1 & "  " & "RG-" & $bpnn & "_AW Commun Manager"
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
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  rajout @mail et groupe Airwatch, pour: " & $idrh1 & "  " & $scomboreadaw
										GUIDelete($hguiaw)
									EndIf
								EndIf
							EndIf
						#EndRegion AW
						$idcheckboxgroup = BitAND(GUICtrlRead($idcheckboxgroup), $gui_checked)
						$idcheckboxrole = BitAND(GUICtrlRead($idcheckboxrole), $gui_checked)
						If $idcheckboxgroup = 1 AND $idcheckboxrole = 0 Then
							If $isdct = 1 Then
								$unite = InputBox("default OU ? 2 or 4 chr$", "ex: MILY, MITE, BPLY, GAUB, MI ... virtuos" & @CRLF & "?? : scan all OUs...", "MI")
								If StringLen($unite) < 2 OR StringLen($unite) > 4 AND NOT ($unite = "virtuos" OR $unite = "??") Then
									ExitLoop
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
								MsgBox(64, $adroot, "No Groups found ! ")
								Return 0
							Else
								$filterou = InputBox("filtering groups ?", "empty = all groups , or type: cbk, dpi ... ", "")
								Global $aselected[1]
								Global $hgui2 = GUICreate("Select/unselect groups then validate with the button [ Apply ] !  (" & $filterou & ")", 1000, 500)
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
								MsgBox(0, "Info !", "aborting, no groups found in: " & $unite, 7)
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
														$actiongroups = $actiongroups & $sitems[$z] & " [target OU ?] ; "
													ElseIf @error = 3 Then
														$actiongroups = $actiongroups & $sitems[$z] & " [deja membre !] ; "
													ElseIf @error = 3 Then
														$actiongroups = $actiongroups & $sitems[$z] & " [pas membre !] ; "
													ElseIf @error >= 4 OR @error = -2147352567 Then
														$actiongroups = $actiongroups & $sitems[$z] & " [accès refusé !] ; "
													EndIf
												EndIf
											Next
											$actiongroups = " Groupes [+]: " & $actiongroups
											GUIDelete($hgui2)
											ExitLoop
										Else
											MsgBox(4160, "warning !", "nothing selected... try again or close window")
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
							ToolTip("Awaiting 3 more secondes, for synchronizing AD ...", 5, 5, "")
							Sleep(3000)
							ToolTip("", 5, 5, "")
						EndIf
						If $idcheckboxrole = 1 AND $idcheckboxgroup = 0 Then
							If $isdct = 0 Then
								MsgBox(0, "Warning !", "Roles only available in domain: DCT.Adt.Local", 7)
								Return 0
							EndIf
							$unite = InputBox("default OU ? ", "ex: MILY, MITE, BPLY, GAUB, ... " & @CRLF & "?? : scan all OUs...", "GAUB")
							If StringLen($unite) < 2 OR StringLen($unite) > 4 AND NOT ($unite = "virtuos" OR $unite = "??") Then
								ExitLoop
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
								Return 0
							Else
								$filterou = InputBox("filtering Roles ?", "empty = all Roles , or type: Bancaire ... ", "")
								Global $aselected[1]
								Global $hgui2 = GUICreate("Select/unselect Role then validate with the button [ Apply ] !  (" & $filterou & ")", 1000, 500)
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
								MsgBox(0, "Info !", "aborting, no Roles found in: " & $unite, 7)
								Return 0
							EndIf
							Global $hbutton = GUICtrlCreateButton("[ Apply ]", 20, 3, 80, 25)
							GUISetState()
							While 1
								Switch GUIGetMsg()
									Case $hbutton
										$aselected = _guictrllistbox_getselitems($hlistbox)
										If $aselected[0] > 1 Then
											MsgBox(4160, "warning !", "multiple 'Roles' selected ! choose just one Role or close window")
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
																$actiongroups = $actiongroups & $sitems[$z] & " [target OU ?] ; "
															ElseIf @error = 3 Then
																$actiongroups = $actiongroups & $sitems[$z] & " [deja membre !] ; "
															ElseIf @error = 3 Then
																$actiongroups = $actiongroups & $sitems[$z] & " [pas membre !] ; "
															ElseIf @error >= 4 OR @error = -2147352567 Then
																$actiongroups = $actiongroups & $sitems[$z] & " [accès refusé !] ; "
															EndIf
														EndIf
													Next
													$actiongroups = "  Role [+]: " & $actiongroups
													$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  roles rajoutés à " & $idrh1 & " :" & @CRLF & $actiongroups
													GUIDelete($hgui2)
													ExitLoop
												Else
													MsgBox(4160, "warning !", "nothing selected... try again or close window")
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
							ToolTip("Awaiting 3 more secondes, for synchronizing AD ...", 5, 5, "")
							Sleep(3000)
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
								If StringInStr(StringLeft($idrh1, 1), "x") Then
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
									MsgBox(0, "", "Aborting reset pwd !", 3)
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
									Else
										$pwdreset = @CRLF & @CRLF & "  Compte Srce: " & $idrh1 & @CRLF & "     -> reinit. mdp à:  " & $password & "   (Le mot de passe ne devra pas etre changé à l'ouverture de session !)"
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $pwdreset
									EndIf
								Else
									MsgBox(0, "warning !", "Reseting password : Access denied , for user:  " & $idrh1)
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "" & "warning !, " & "Reset mot de passe windows, access refusé pour:  " & $idrh1
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
									$outputidrh1 = "	   Directive applied:   " & $scomboreaddirectives & "	,  OU d'origine= " & $ousourcedir & @CRLF
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & $outputidrh1
								Else
									$outputidrh1 = "	   Directive Not applied !   " & $scomboreaddirectives & "	,  OU d'origine= " & $ousourcedir & @CRLF
									$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & $outputidrh1
								EndIf
							Else
								$outputidrh1 = ""
							EndIf
							$outputidrh1 = $outputidrh1 & "  Object:  " & _ad_getobjectattribute($idrh1, "distinguishedName") & "  | Expiration ?:  " & $accountexpire & @CRLF & "									Proprietaire:     " & $proprietaire & @CRLF & "									compte crée  le:  " & $createdsrce & @CRLF & "									mdp  changé  le:  " & $lastchanged & @CRLF & @CRLF & "  email :     " & $mail & $pwdreset & @CRLF & $pwdinfo & @CRLF & $description & @CRLF & "  Drives of user: " & $idrh1 & @CRLF & $drivesidrh1 & @CRLF & @CRLF & "  Comment 'Mes scan' of user: " & $idrh1 & @CRLF & $commentidrh1 & @CRLF & @CRLF & "  Printers for user: " & $idrh1 & @CRLF & $printersidrh1 & @CRLF & @CRLF & "  Groups for user: " & $idrh1 & @CRLF & $groupidrh_final
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
								$actiongroupsar = " Removed groups: " & $actiongroupsar
								$outputidrh1 = $outputidrh1 & @CRLF & @CRLF & "	Modifs: " & $actiongroupsar & @CRLF & @CRLF
								$lastiter = $lastiter & @CRLF & @CRLF & "	Modifs: " & $actiongroupsar & @CRLF & @CRLF
							EndIf
							If $listmanualdrives <> "" Then
								$outputidrh1 = $outputidrh1 & @CRLF & @CRLF & "	Modifs: adding new drive:  " & $listmanualdrives & @CRLF & @CRLF
								$lastiter = $lastiter & @CRLF & @CRLF & "	Modifs: adding new drive:  " & $listmanualdrives & @CRLF & @CRLF
							EndIf
							If $idcheckboxgroup <> 2 Then
								ClipPut($outputidrh1)
								$historik = $historik & @CRLF & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $outputidrh1
								GUICtrlSetData($aff, @CRLF & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $outputidrh1 & @CRLF & @CRLF & "			=====	=====	=====	=====	=====			" & @CRLF & @CRLF, 1)
							EndIf
							$stateidrh1 = _ad_isobjectdisabled($idrh1)
							If $stateidrh1 = 1 Then
								$t = MsgBox(4, "", "Souhaitez-vous réactiver le compte source: " & $idrh1 & " ?")
								If $t = 6 Then
									$result = _ad_enableobject($idrh1)
									If $result = 1 Then
										ToolTip("Compte " & $idrh1 & " réactivé !", 5, 5, "")
										GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Compte " & $idrh1 & " réactivé !", 1)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & ", compte réactivé !"
										Sleep(2500)
										ToolTip("", 5, 5, "")
									EndIf
								ElseIf $t = 7 Then
								EndIf
							EndIf
							$userexist2 = _ad_objectexists($idrh2)
							If StringUpper($idrh2) = "" OR StringLen($idrh2) = 0 OR $userexist2 = 0 Then
								Sleep(2500)
								ToolTip("", 5, 5, "")
							Else
								$checkbuttondrives = BitAND(GUICtrlRead($idcheckboxdrives), $gui_checked)
								$checkbuttongroups = BitAND(GUICtrlRead($idcheckboxgroups), $gui_checked)
								If $checkbuttondrives = 1 Then
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
												MsgBox(0, "Warning !", "  White space found in drive (from Idrh Srce): " & $idrh1 & @CRLF & $drivewspace & @CRLF & "  ..removed WS !", 15)
												GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  White space found in drive (from Idrh Srce): " & $idrh1 & @CRLF & $drivewspace & "  ..removed WS !" & @CRLF & @CRLF, 1)
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
										GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "No drives to copy from Srce:  " & $idrh1 & " !" & @CRLF & @CRLF, 1)
									EndIf
									If UBound($cpydrives1) <> 0 AND $cpydrives1[1] <> "" Then
										_arraydisplay($cpydrives1, "cpy drives from '" & $idrh1 & "'   to   '" & $idrh2 & "'")
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
										MsgBox(0, "warning !", "no Drives for user srce: " & $idrh1 & " , stopping process...")
										GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "no Drives for user srce: " & $idrh1 & " , stopping process..." & @CRLF & @CRLF, 1)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "no Drives for user srce: " & $idrh1 & " , stopping process..."
									ElseIf UBound($cpydrives3) = 0 Then
										MsgBox(0, "warning !", "All Drives letters are already used by user dest. " & $idrh2 & " !" & " , stopping process..." & @CRLF & "already used:" & @CRLF & $doublons)
										GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "All Drives letters are already used by user dest. " & $idrh2 & " !" & " , stopping process..." & @CRLF & "already used:" & @CRLF & $doublons & @CRLF & @CRLF, 1)
										$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "All Drives letters are already used by user dest. " & $idrh2 & " !" & " , stopping process..." & @CRLF & "already used:" & @CRLF & $doublons
									Else
										Global $ivalue = _ad_modifyattribute($idrh2, "LaPoste00-NetworkDrive", $cpydrives3, 3)
										If $ivalue = 1 Then
											If $doublons = "" Then
												MsgBox(0, "Info cpy drives", "successfully added drives to " & $idrh2 & "  from " & $idrh1 & @CRLF & @CRLF & $cpydrives1bis, 7)
												GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "successfully added drives to " & $idrh2 & "  from " & $idrh1 & @CRLF & @CRLF & $cpydrives1bis & @CRLF & @CRLF, 1)
												$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "successfully added drives to " & $idrh2 & "  from " & $idrh1 & @CRLF & @CRLF & $cpydrives1bis
											Else
												MsgBox(0, "Warning cpy drives", "not added some drives to " & $idrh2 & "  from " & $idrh1, 7)
												GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "not added some drives to " & $idrh2 & "  from " & $idrh1 & @CRLF & @CRLF, 1)
												$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":  not added some drives to " & $idrh2 & "  from " & $idrh1 & @CRLF
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
												$result = "Warning !  :  " & "cpy drives from : " & $idrh1 & " to " & $idrh2 & @CRLF & @CRLF & "Drives added:  :)" & @CRLF & $cpydrives1bis & @CRLF & @CRLF & "Drives letter already used by  '" & $idrh2 & "'  and can't be added..." & @CRLF & "-You have to manually adding drives by changing drives Letters for each of them below...  :(" & $doublons & @CRLF & @CRLF & "'Drives list'  before cpy for user dest.  '" & $idrh2 & "' :" & @CRLF & $drivesidrh2 & @CRLF & @CRLF & "'Drives list' after  cpy for user dest.   '" & $idrh2 & "' :" & @CRLF & $drivesidrh2bis & @CRLF
												ClipPut($result)
												GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $result & @CRLF & @CRLF & "			=====	=====	=====	=====	=====			" & @CRLF & @CRLF, 1)
												$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & $result
											EndIf
										ElseIf @error = 1 Then
											MsgBox(0, "info cpy drives", "user does not exist " & $idrh2)
											GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "user does not exist " & $idrh2 & @CRLF & @CRLF, 1)
										Else
											MsgBox(0, "info cpy drives", "[Return code " & @error & "] from Active Directory for user " & $idrh2 & @CRLF & "Access denied ...")
											GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "[Return code " & @error & "] from Active Directory for user " & $idrh2 & "  Access denied ..." & @CRLF & @CRLF, 1)
										EndIf
									EndIf
								Else
								EndIf
								If $idcheckscancpy = 1 Then
									$commentidrh2 = $commentidrh1
									If $commentidrh2 = "" Then
										MsgBox(0, "Warning !", "user 'Srce': " & $idrh1 & ", comment 'Mes scan' undefined !" & @CRLF & "Aborting copy Comment 'Me scan' from 'Srce' to user 'Dest': " & $idrh2 & " !", 15)
										GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "user 'Srce': " & $idrh1 & ", comment 'Mes scan' undefined !" & @CRLF & "Aborting copy Comment 'Me scan' from 'Srce' to user 'Dest': " & $idrh2 & " !" & @CRLF & @CRLF, 1)
									Else
										$commentidrh2 = StringReplace($commentidrh2, $idrh1, $idrh2)
										_ad_modifyattribute($idrh2, "Comment", "", 1)
										Global $ivalue = _ad_modifyattribute($idrh2, "Comment", $commentidrh2, 2)
										If $ivalue = 1 Then
											MsgBox(0, "info cpy Comment 'Mes scan'", "successfully added comment to " & $idrh2 & "  from " & $idrh1, 15)
											GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "successfully added comment to " & $idrh2 & "  from " & $idrh1 & @CRLF & @CRLF, 1)
										ElseIf @error = 1 Then
											MsgBox(0, "info cpy Comment 'Mes scan'", "[user does not exist] " & $idrh2, 15)
											GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "[user does not exist] " & $idrh2 & @CRLF & @CRLF, 1)
										Else
											MsgBox(0, "info cpy Comment 'Mes scan'", "[Return code " & @error & "]  from Active Directory for  " & $idrh2 & @CRLF & "Access denied ?...", 15)
											GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "[Return code " & @error & "]  from Active Directory for  " & $idrh2 & @CRLF & "Access denied ?..." & @CRLF & @CRLF, 1)
										EndIf
									EndIf
								Else
								EndIf
								If $checkbuttongroups = 1 Then
									_arraydisplay($groupidrh1_add, "list of groups selected for adding to user " & $idrh2)
									$idrh2 = _ad_samaccountnametofqdn($idrh2)
									$mail = _ad_getobjectattribute($idrh2, "mail")
									For $k = 0 To UBound($groupidrh1_add) - 1
										$ivalue = _ad_addusertogroup($groupidrh1_add[$k], $idrh2)
										If $ivalue = 1 Then
											ToolTip("User '" & $idrh2 & "' successfully assigned to group '" & $groupidrh1_add[$k] & "'", 5, 5, "")
											GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "User '" & $idrh2 & "' successfully assigned to group '" & $groupidrh1_add[$k] & "'", 1)
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "User '" & $idrh2 & "' successfully assigned to group '" & $groupidrh1_add[$k] & "'"
										ElseIf @error = 1 Then
											MsgBox(64, "Active Directory", "Group '" & $groupidrh1_add[$k] & "' does not exist")
											GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Active Directory", "Group '" & $groupidrh1_add[$k] & "' does not exist", 1)
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":  Group '" & $groupidrh1_add[$k] & "' n'existe pas !"
										ElseIf @error = 2 Then
											MsgBox(64, "Active Directory", "User '" & $idrh2 & "' does not exist")
											GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "User '" & $idrh2 & "' does not exist", 1)
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":   User '" & $idrh2 & "' n'existe pas !"
										ElseIf @error = 3 Then
											ToolTip("User '" & $idrh2 & "' already memberof group '" & $groupidrh1_add[$k] & "'", 5, 5, "")
											GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "User '" & $idrh2 & "' already memberof group '" & $groupidrh1_add[$k] & "'", 1)
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":  User '" & $idrh2 & "' already memberof group '" & $groupidrh1_add[$k] & "'"
										Else
											MsgBox(64, "Active Directory", "Return code '" & @error & "' from Active Directory for adding group " & $groupidrh1_add[$k] & " to " & $idrh2 & @CRLF & "Access denied ...")
											GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "Return code '" & @error & "' from Active Directory for adding group " & $groupidrh1_add[$k] & " to " & $idrh2 & @CRLF & "Access denied ...", 1)
											$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":  " & $groupidrh1_add[$k] & "  =>  " & $idrh2 & ", access refusé !"
										EndIf
										If StringInStr($groupidrh1_add[$k], "_AW Commun Manager") AND $mail = "" Then
											_ad_removeuserfromgroup($groupidrh1_add[$k], $idrh2)
											$historik = $historik & @CRLF & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":  " & $groupidrh1_add[$k] & "  =>  " & $idrh2 & ", removed :  no email Description is set !" & @CRLF
										EndIf
										If StringInStr($groupidrh1_add[$k], "_AW BPE") AND $mail = "" Then
											_ad_removeuserfromgroup($groupidrh1_add[$k], $idrh2)
											$historik = $historik & @CRLF & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":  " & $groupidrh1_add[$k] & "  =>  " & $idrh2 & ", removed :  no email Description is set !" & @CRLF
										EndIf
										If StringInStr($groupidrh1_add[$k], "_AW BPE Restreint") AND $mail = "" Then
											_ad_removeuserfromgroup($groupidrh1_add[$k], $idrh2)
											$historik = $historik & @CRLF & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":  " & $groupidrh1_add[$k] & "  =>  " & $idrh2 & ", removed :  no email Description is set !" & @CRLF
										EndIf
									Next
								Else
								EndIf
							EndIf
							ToolTip("", 5, 5, "")
							Return 0
						Else
							MsgBox(0, "Warning !", "userID:  '" & $idrh1 & "'  does not exist !", 5)
							GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "userID:  '" & $idrh1 & "'  does not exist !", 1)
							$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & ":  utilisateur,  " & $idrh1 & ", n'existe pas !"
							Return 0
						EndIf
					EndIf
				Case $gui_event_close
					_terminate()
					Return 0
				Case $idabout
					about()
					$t = MsgBox(0, "About AD Tools", "AD Tools - Nicolas RISTOVSKI" & @CRLF & @CRLF & $lastdatecompile)
					_Exit()
			EndSwitch
		WEnd
	EndFunc

#EndRegion

#Region Localgroups

	Func readlocalgroup2()
		FileInstall("addusers.exe", ".\")
		FileSetAttrib("addusers.exe", "+h")
		Global $cpname
		Global $readcombo
		Global $listmembers
		Global $userexist
		Global $hguilg = GUICreate("AD Tools [Add / Remove Members]-[Local Group]", 669, 433, 196, 128)
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
					FileDelete(@WorkingDir & "\addusers.exe")
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
									GUICtrlSetData($listmemberslg, $listatraiter[$z] & @CRLF, 1)
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
									GUICtrlSetData($listmemberslg, $listatraiter[$z] & @CRLF, 1)
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
										GUICtrlSetData($listmemberslg, "'" & $copygrps[$z] & "   [    adding    ]  " & GUICtrlRead($inputcomputer2) & "  from:  " & GUICtrlRead($inputcomputer) & @CRLF, 1)
									Else
										RunWait(@ComSpec & " /c " & "net localgroup " & Chr(34) & $copygrps[$z] & Chr(34) & " " & GUICtrlRead($inputcomputer2) & " /DELETE", @WorkingDir, @SW_HIDE)
										GUICtrlSetData($listmemberslg, "'" & $copygrps[$z] & "   [   deleting   ]  " & GUICtrlRead($inputcomputer2) & "  from:  " & GUICtrlRead($inputcomputer) & @CRLF, 1)
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
											GUICtrlSetData($listmemberslg, $listatraiter[$z] & @CRLF, 1)
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
								GUICtrlSetData($listmemberslg, "'" & $copygrps[$z] & "   [   deleting   ]  " & GUICtrlRead($inputcomputer2) & @CRLF, 1)
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
									GUICtrlSetData($listmemberslg, $listatraiter[$z] & @CRLF, 1)
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
		FileInstall("addusers.exe", ".\")
		FileSetAttrib("addusers.exe", "+h")
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
								GUICtrlSetData($affiche, $sfolder & " => " & $line & @CRLF, 1)
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
						GUICtrlSetData($affiche, "" & $exist & @CRLF, 1)
						If _ad_objectexists(GUICtrlRead($users)) Then
							$exist = "'S'  [ " & GUICtrlRead($users) & "  exist ? => YES ]"
						Else
							$exist = "'S'  [ " & GUICtrlRead($users) & "  exist ? => NO  ]"
						EndIf
						GUICtrlSetData($affiche, "" & $exist & @CRLF, 1)
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
													GUICtrlSetData($affiche, "'" & $scomboreadgl & "'  [  adding   ]  " & $user & "" & $exist & @CRLF, 1)
												Else
													RunWait(@ComSpec & " /c " & "net localgroup " & Chr(34) & $scomboreadgl & Chr(34) & " " & $user & " /DELETE", @WorkingDir, @SW_HIDE)
													GUICtrlSetData($affiche, "'" & $scomboreadgl & "'  [  removing ]  " & $user & "" & $exist & @CRLF, 1)
												EndIf
												GUIDelete($hguigl)
												ExitLoop
											EndIf
										Case $gui_event_close
											GUIDelete($hguigl)
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
			GUICtrlSetData($affiche, $readlinelg & @CRLF, 1)
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
		GUICtrlSetData($aff, @CRLF & "Bitlocker Error ! No computers found in this OU: " & $sout, 1)
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
			GUICtrlSetData($aff, @CRLF & "Bitlocker Error ! No computers found in this OU: " & $sout, 1)
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
									GUICtrlSetData($aff, @CRLF & "Bitlocker Recovery:  " & "@error: " & @error & ", @extended: " & @extended & "  No rights or Bitlocker not found for: " & $sad_ou & " !  Aborting...", 1)
								Else
									GUICtrlSetData($aff, @CRLF & $sad_ou & @CRLF & _arraytostring($aresult), 1)
									$aresult = _ad_getobjectproperties($aresult[1])
									If @error <> 0 Then
										GUICtrlSetData($aff, "Bitlocker Recovery:  " & "@error: " & @error & ", @extended: " & @extended & "  No rights or Bitlocker not found !   " & "  Aborting...", 1)
									Else
										GUICtrlSetData($aff, @CRLF & $sad_ou & @CRLF & _arraytostring($aresult), 1)
									EndIf
								EndIf
							EndIf
						Next
						ToolTip("", 5, 5)
						ExitLoop
					Else
						MsgBox(4160, "warning !", "nothing selected... try again or close window")
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
		GUICtrlSetData($aff, @CRLF & "SCCM PITR /DCT ! No computers found in this OU: " & $sout, 1)
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
					MsgBox(4160, "warning !", "nothing selected... try again or close window")
				EndIf
			Case $gui_event_close
				GUIDelete($hgui2)
				ExitLoop
		EndSwitch
	WEnd
	If IsArray($sitems) = 0 Then
		MsgBox(4160, "warning !", "no computer(s) selected !... abort !")
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

GUICTRLSETDATA($aff,  @crlf & "  Time: " & @hour & ":" & @min & ":" & @sec & "=>   Initial computer groups: " & @crlf & $result  & "  Time: " & @hour & ":" & @min & ":" & @sec & "   Initial computer groups   <=" & @crlf,1)
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
					MsgBox(4160, "warning !", "nothing selected... try again or close window")
				EndIf
			Case $gui_event_close
				$idcheckboxgroup = 2
				GUIDelete($hgui2)
				ExitLoop
		EndSwitch
	 WEnd

	 ; boucle for $i (computers[$i])
	 GUIDelete($hgui2)
					For $i = 1 To $sitems[0]
						$computers = $sitems[$i]
						   if $computers<>"" then
							  $result=$result & "> " & $computers & @CRLF
						   EndIf
						$userexist = _ad_objectexists($computers)
						If $userexist = 1 Then
							For $z = 1 To $sitemss[0]
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
					GUICtrlSetData($aff, @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $result & @CRLF & "			=====	=====	=====	=====	=====			" & @CRLF, 1)
					$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $result

	 ; fin boucle for $i (computers[$i])

	ToolTip("", 5, 5)
	;GUIDelete($hgui2)
	Return 0
EndFunc

#Region Directives Metiers
;/////////////////////////////////////////////////////  Directives Metiers domaine DCT uniquement  /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Func directives() ;routine principale des Directives Metiers
	$defautdc = $defautdcinit

	#Region IHM Directives
		If $isdct = 0 Then
			MsgBox(0, "warning !", "'Directives' only working for domain: DCT.Adt.Local", 7)
			Return 0
		EndIf
		Global $ousourcedir = _ad_getobjectattribute($idrh1, "distinguishedName")
		$ousourcedir = StringSplit($ousourcedir, ",")
		$ousourcedir = $ousourcedir[3]
		$ousourcedir = StringTrimLeft($ousourcedir, 3)
		$hguidir = GUICreate("Select Directive and close window...", 650, 80)
	#EndRegion

	#Region Liste des Directives
		If $directives = "" Then
			liste_directives()
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
					MsgBox(0, "Warning !", "inexistant Directive... aborting !", 7)
					GUIDelete($hguidir)
					ExitLoop
				EndIf
				GUIDelete($hguidir)
				If $scomboreaddirectives = "[ Zero Directives ]" Then
					$t = MsgBox(4, "Remove Directive ?", "Do you want to remove any Directive found for: " & $idrh1 & " ...?")
				Else
					$t = MsgBox(4, "Apply Directive ?", "Do you want to lauch process for: " & @CRLF & $scomboreaddirectives)
				EndIf
				If $t = 6 Then
					SplashTextOn("", "Applying Directive:  " & $scomboreaddirectives & " !", 1100, 100, -1, -1, 1, -1, 13, 600)
					Sleep(1200)
					SplashOff()
					$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  Valider une 'Directive'"
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
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									$scomboreaddirectives = " [ Zero Directives ] = Enleve toutes les Directives connues !"
									$dirapplied = 1
								EndIf
								If $scomboreaddirectives = "_SM_Sites Tertiaires RLP" OR $scomboreaddirectives = "_SM_Site_Tertiaire_Enseigne" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "_SM_Sites Tertiaires RLP", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Sites Tertiaires RLP", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_Sites Tertiaires RLP" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_Sites Tertiaires RLP" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_Sites Tertiaires RLP" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_GUICHET (VirtuOS)" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE ... " & @CRLF & "_SM_GUICHET", "GAUB")
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_GUICHET", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_GUICHET" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_GUICHET" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_GUICHET" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
									EndIf
									If $dirapplied = 1 Then
										_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
										_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)
										_ad_addusertogroup("USR_BP_GUICHET_ESPACE_CO", $idrh1)
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "19 _SM_Conseil Bancaire" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									_arraysort($groupsidrh1, 0, 1)
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE ... " & @CRLF & "_SM_Conseil Bancaire", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Conseil Bancaire", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_Conseil Bancaire" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_Conseil Bancaire" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_Conseil Bancaire" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
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
								If $scomboreaddirectives = "_SM_Conseil Bancaire W10" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									_arraysort($groupsidrh1, 0, 1)
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE ... " & @CRLF & "_SM_Conseil Bancaire W10 (sur BPnn)", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_DET", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_DET" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_DET" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_DET" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
									EndIf
									If $dirapplied = 1 Then
										_ad_addusertogroup("RG-PITR_CM_COBA_Agent", $idrh1)
										_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
										_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Conseil Bancaire" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									_arraysort($groupsidrh1, 0, 1)
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE ... " & @CRLF & "_SM_Conseil Bancaire (sur BPnn)", $ousourcedir)
									move_user()
									$t = MsgBox(4, "choix W10/W7", "_SM_Conseil Bancaire , poste W10=(Oui)?" & @CRLF & "sinon poste W7=(Non)")
									If $t = 6 Then
										Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_DET", $idrh1)
									ElseIf $t = 7 Then
										Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Conseil Bancaire", $idrh1)
									EndIf
									If $ivalue = 1 Then
										$dirapplied = 1
										If $t = 6 Then
											MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_DET" & "'")
										EndIf
										If $t = 7 Then
											MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_Conseil Bancaire" & "'")
										EndIf
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_DET" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_DET" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
									EndIf
									If $dirapplied = 1 Then
										If $t = 6 Then
											_ad_addusertogroup("RG-PITR_CM_COBA_Agent", $idrh1)
										EndIf
										_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
										_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Siege_Enseigne" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "_SM_Siege_Enseigne", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Siege_Enseigne", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_Siege_Enseigne" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_Siege_Enseigne" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_Siege_Enseigne" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Siege Banque" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "_SM_Siege Banque", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Siege Banque", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_Siege Banque" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_Siege Banque" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_Siege Banque" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Siege Banque_LBPGP" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "_SM_Siege Banque_LBPGP", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Siege Banque_LBPGP", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_Siege Banque_LBPGP" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_Siege Banque_LBPGP" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_Siege Banque_LBPGP" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Siege Banque_LBPAI" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "_SM_Siege Banque_LBPAI", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Siege Banque_LBPAI", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_Siege Banque_LBPAI" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_Siege Banque_LBPAI" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_Siege Banque_LBPAI" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Siege Banque_FondEcranFixe" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "_SM_Siege Banque_FondEcranFixe", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Siege Banque_FondEcranFixe", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_Siege Banque_FondEcranFixe" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_Siege Banque_FondEcranFixe" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_Siege Banque_FondEcranFixe" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
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
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, MITO ... " & @CRLF & "_SM_DISFE", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_DISFE", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_DISFE" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_DISFE" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_DISFE" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Siege Banque_LBPF" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "_SM_Siege Banque_LBPF", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Siege Banque_LBPF", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_Siege Banque_LBPF" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_Siege Banque_LBPF" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_Siege Banque_LBPF" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_DET" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "_SM_DET", "GAUB")
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_DET", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_DET" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_DET" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_DET" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
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
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "SM_Centres Financiers et Centres Nationaux", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Centres Financiers et Centres Nationaux", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_Centres Financiers et Centres Nationaux" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_Centres Financiers et Centres Nationaux" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_Centres Financiers et Centres Nationaux" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Utilisateurs_Postes_SIA" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "SM_Utilisateurs_Postes_SIA", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Utilisateurs_Postes_SIA", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_Utilisateurs_Postes_SIA" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_Utilisateurs_Postes_SIA" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_Utilisateurs_Postes_SIA" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_RF_INDET" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "SM_RF_INDET", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_RF_INDET", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_RF_INDET" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_RF_INDET" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_RF_INDET" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
									EndIf
									If $dirapplied = 1 Then
										Global $groupsidrh1 = _ad_getusergroups($idrh1)
										affiche_remove_groupes()
									EndIf
								EndIf
								If $scomboreaddirectives = "_SM_Management Commercial Unique" Then
									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									If @error > 0 Then
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & "SM_Management Commercial Unique", $ousourcedir)
									move_user()
									Global $ivalue = _ad_addusertogroup("RG-" & $ou & "_SM_Management Commercial Unique", $idrh1)
									If $ivalue = 1 Then
										$dirapplied = 1
										MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & "_SM_Management Commercial Unique" & "'")
									ElseIf @error = 1 Then
										MsgBox(64, "", "Group '" & "RG-" & $ou & "_SM_Management Commercial Unique" & "' does not exist")
									ElseIf @error = 2 Then
										MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
									ElseIf @error = 3 Then
										MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & "_SM_Management Commercial Unique" & "'")
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
											MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
										$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
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
										MsgBox(0, "Info !", "user " & $idrh1 & "has no groups...", 7)
										Return 0
									EndIf
									auto_remove()
									Global $restoredirectives = $actiongroups
									$restoredirectives = StringReplace($restoredirectives, " ; ", "|")
									$actiongroups = " Removed obsolete Directives groups: " & $actiongroups
									If StringInStr($scomboreaddirectives, ";") Then
										$dira = StringSplit($scomboreaddirectives, ";")
										$direaid = _stringexplode($dira[1], "]", -2)
										$dirid = $direaid[1]
										$dirpitr = $dira[2]
									EndIf
									$ou = InputBox("default OU ?", "ex: BPLY, BPRE, CFLY ... " & @CRLF & $dirid, $ousourcedir)
									move_user()
								#EndRegion mask

								Global $ivalue = _ad_addusertogroup("RG-" & $ou & $dirid, $idrh1)
								If $ivalue = 1 Then
									$dirapplied = 1
									MsgBox(64, "", "User '" & $idrh1 & "' successfully assigned to group '" & "RG-" & $ou & $dirid & "'")
								ElseIf @error = 1 Then
									MsgBox(64, "", "Group '" & "RG-" & $ou & $dirid & "' does not exist")
								ElseIf @error = 2 Then
									MsgBox(64, "", "User '" & $idrh1 & "' does not exist")
								ElseIf @error = 3 Then
									MsgBox(64, "", "User '" & $idrh1 & "' is already a member of group '" & "RG-" & $ou & $dirid & "'")
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
										MsgBox(64, "Active Directory ", $idrh1 & "' successfully replaced to '" & $ousourcedir & "'" & @CRLF & @CRLF & "Directive Not Applied...")
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
									$actiongroups = StringReplace($actiongroups, " Removed obsolete Directives groups: ", " Replaced original Directives groups: ")
								EndIf

							    If $dirapplied = 1 Then
									_ad_addusertogroup($dirpitr, $idrh1)

									#Region cas particuliers

									Select
											Case StringInStr($dirpitr, "RG-PITR_CM_DIR-BP_DS_REC_RE_RC") AND StringInStr($dirid, "_SM_DET")
												Global $regategaub = ""
												$regategaub = InputBox("Regate (6 chiffres) ?", "ex: 694100" & @CRLF & "RG-GAUB_[regate]_GESTION_GROUPE" & @CRLF & "SG-GAUB_[regate]_PRIVE" & @CRLF & "Oui => (REC/RE,DS et CECI)" & @CRLF & "Non => (CCPRO,GCPRO,GDCPRO,GESPRO,RCPART)", "")
												_ad_addusertogroup("RG-GAUB_" & $regategaub & "_GESTION_GROUPE", $idrh1)
												_ad_addusertogroup("SG-GAUB_" & $regategaub & "_PRIVE", $idrh1)
												_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
												_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)

											Case StringInStr($dirpitr, "RG-PITR_CM_DIR-BP_COCLI_COFI_Itinerant") AND StringInStr($dirid, "_SM_DET")
												_ad_addusertogroup("RG-PITR_CM_COBA_Agent", $idrh1)

										    Case StringInStr($dirpitr, "RG-PITR_CM_DIR-TERTIAIRE_MAISON_HABITAT") AND StringInStr($dirid, "_SM_Sites Tertiaires RLP")

											Case StringInStr($dirpitr, "RG-PITR_CM_DIR-BP_BPE_RCP") AND StringInStr($dirid, "_SM_DET")
												_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
												_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)

											Case StringInStr($dirpitr, "RG-PITR_CM_DIR-BP_CCPRO_GCPRO") AND StringInStr($dirid, "_SM_DET")
												_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
												_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)

											Case StringInStr($dirpitr, "SM_Siege Banque_FondEcranFixe")
												_ad_addusertogroup("SG-GAUB_ACCES_EVS", $idrh1)
												_ad_addusertogroup("USR_BP_GUICHET_GENE", $idrh1)

									EndSelect

									#EndRegion Cas Particuliers

									Global $groupsidrh1 = _ad_getusergroups($idrh1)
									affiche_remove_groupes()
								EndIf
							#EndRegion nouvelles Directives

			#EndRegion traitement des Directives

					 EndIf

				ElseIf $t = 7 Then
					SplashTextOn("", "Aborting " & $scomboreaddirectives & " !", 1100, 100, -1, -1, 1, -1, 13, 600)
					Sleep(1000)
					SplashOff()
					$scomboreaddirectives = ""
					$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  Annulé la validation 'Directive'"
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
   ;VirtuOS 28-01-2022 (Vir_01 à Vir_014)
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
   ;VirtuOS 28-01-2022 (Vir_01 à Vir_014)
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
		_arrayadd($array2, "22 RG-[EAID]_SM_Siege_Enseigne")
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
		_arrayadd($array2, "49 RG-[EAID]_SM_Siege_Enseigne")
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
		_arrayadd($array2, "68 RG-[EAID]_SM_Siege_Enseigne")
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
		_arrayadd($array2, "145 RG-[EAID]_SM_Siege_Enseigne")
		_arrayadd($array2, "146")
		_arrayadd($array2, "147")
		_arrayadd($array2, "148")
		_arrayadd($array2, "149")
		_arrayadd($array2, "150")
		_arrayadd($array2, "151")
		_arrayadd($array2, "152 RG-[EAID]_SM_Siege_Enseigne")
		_arrayadd($array2, "153")
		_arrayadd($array2, "154")
		_arrayadd($array2, "155 RG-[EAID]_SM_Siege_Enseigne")
		_arrayadd($array2, "156")
		_arrayadd($array2, "157 RG-[EAID]_SM_Siege_Enseigne")
		_arrayadd($array2, "158")
		_arrayadd($array2, "159")
		_arrayadd($array2, "160")
		_arrayadd($array2, "161")
		_arrayadd($array2, "162")
		_arrayadd($array2, "163 RG-[EAID]_SM_Siege_Enseigne")
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
		_arrayadd($array2, "186 RG-[EAID]_SM_Siege_Enseigne")
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
		_arrayadd($array2, "362 RG-[EAID]_SM_Siege_Enseigne")
		_arrayadd($array2, "363 RG-[EAID]_SM_Siege_Enseigne")
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
		_arrayadd($array2, "408 RG-[EAID]_SM_Siege_Enseigne")
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
		_arrayadd($array2, "739 RG-[EAID]_SM_Siege_Enseigne")
		_arrayadd($array2, "740 RG-[EAID]_SM_Siege_Enseigne")
		_arrayadd($array2, "741 RG-[EAID]_SM_Siege_Enseigne")
		_arrayadd($array2, "742 RG-[EAID]_SM_Siege_Enseigne")
		_arrayadd($array2, "743 RG-[EAID]_SM_Siege_Enseigne")
		_arrayadd($array2, "744 RG-[EAID]_SM_Siege_Enseigne")
		_arrayadd($array2, "745 RG-[EAID]_SM_Siege_Enseigne")
		_arrayadd($array2, "746 RG-[EAID]_SM_Siege_Enseigne")
		_arrayadd($array2, "747 RG-[EAID]_SM_Siege_Enseigne")
		_arrayadd($array2, "748 RG-[EAID]_SM_Siege_Enseigne")
		_arrayadd($array2, "749 RG-[EAID]_SM_Siege_Enseigne")
		_arrayadd($array2, "750 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_DIRECTEURS_ASS_DIR")
		_arrayadd($array2, "751 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_FILIERE RH_PSST")
		_arrayadd($array2, "752 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_STANDARD_STD")
		_arrayadd($array2, "753 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_COMMUNICATION")
		_arrayadd($array2, "754 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_3631_EXP_CLIENT")
		_arrayadd($array2, "755")
		_arrayadd($array2, "756 RG-[EAID]_SM_Sites Tertiaires RLP;RG-PITR_CM_DIR-TERTIAIRE_STANDARD_BANCAIRE")
		_arrayadd($array2, "757 RG-[EAID]_SM_Siege_Enseigne;RG-PITR_CM_DIR-TERTIAIRE_FIDUCIAIRE_TRANSPORT")
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
		_arrayadd($array2, "770 RG-[EAID]_SM_Siege_Enseigne;RG-PITR_CM_DIR-TERTIAIRE_EXPERT_FINANCE_COMPTA")
		_arrayadd($array2, "771 RG-[EAID]_SM_Siege_Enseigne;RG-PITR_CM_DIR-SIEGE_BANQUE_SG")
		_arrayadd($array2, "772 RG-[EAID]_SM_Siege_Enseigne;RG-PITR_CM_DIR-POLE_ASSURANCE_CNAH")
		_arrayadd($array2, "773 RG-[EAID]_SM_Siege_Enseigne;RG-PITR_CM_DIR-SIEGE_BANQUE_DEDT_STD")
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
		;VirtuOS 28-01-2022
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
		_arrayinsert($array, 0, "_SM_Siege_Enseigne")
		_arrayinsert($array, 0, "_SM_Conseil Bancaire W10")
		_arrayinsert($array, 0, "19 _SM_Conseil Bancaire")
		_arrayinsert($array, 0, "_SM_GUICHET (VirtuOS)")
		_arrayinsert($array, 0, "_SM_Management Commercial Unique")
		_arrayinsert($array, 0, "_SM_Sites Tertiaires RLP")
		_arrayinsert($array, 0, "[ Zero Directives ]")
	#EndRegion mise en forme $array2 => $array

Global $directives = _arraytostring($array, "|")
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
		MsgBox(64, "Active Directory ", $idrh1 & "' successfully moved to '" & $ou & "'")
		ToolTip("synchronising AD", 5, 5)
		Sleep(3000)
		ToolTip("", 5, 5)
	ElseIf @error = 1 Then
		MsgBox(64, "Active Directory ", "Target OU '" & $ou & "' does not exist")
	ElseIf @error = 2 Then
		MsgBox(64, "Active Directory ", "Object '" & $idrh1 & "' does not exist")
	ElseIf @error = 3 Then
		MsgBox(64, "Active Directory ", "Object '" & $idrh1 & "' already in OU '" & $ou & "'")
	ElseIf @error >= 4 OR @error = -2147352567 Then
		MsgBox(64, "Active Directory ", "You have no permissions to move user " & $idrh1)
	Else
		MsgBox(0, "warning !", "User Idrh Srce '" & $idrh1 & "already exists in AD !", 7)
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
		Global $hgui2 = GUICreate("Select obsolete Role to REMOVE from '" & $idrh1 & "' , then validate with the button [ Apply ] ! (or close window)", 1000, 500)
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
			MsgBox(0, "Info !", "aborting, no Groups found in: " & $idrh1, 7)
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
						MsgBox(4160, "warning !", "nothing selected... try again or close window")
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
			MsgBox(0, "warning !", "'LBPAI / Illiade' only working for domain: DCT.Adt.Local", 7)
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
			MsgBox(64, "Active Directory ", $idrh1 & "' successfully moved to '" & $ou & "'")
			$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & " déplacé vers " & $ou
			ToolTip("synchronising AD", 5, 5)
			Sleep(3000)
			ToolTip("", 5, 5)
		ElseIf @error = 1 Then
			MsgBox(64, "Active Directory ", "Target OU '" & $ou & "' does not exist")
			Return 0
		ElseIf @error = 2 Then
			MsgBox(64, "Active Directory ", "Object '" & $idrh1 & "' does not exist")
			Return 0
		ElseIf @error = 3 Then
			MsgBox(64, "Active Directory ", "Object '" & $idrh1 & "' already in OU '" & $ou & "'")
		ElseIf @error >= 4 OR @error = -2147352567 Then
			MsgBox(64, "Active Directory ", "You have no permissions to move user " & $idrh1)
			$historik = $historik & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $idrh1 & "non deplacé vers " & $ou & ",  accès refusé !"
			Return 0
		Else
			MsgBox(0, "warning !", "User Idrh Srce '" & $idrh1 & "already exists in AD !", 7)
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
				Global $hgui2 = GUICreate("Select LBPAI/Iliade Category to [+] for user '" & $idrh1 & "' , then validate with the button [ Apply ] ! (or close window to abort/exit)...", 1000, 500)
			Else
				Global $hgui2 = GUICreate("Select LBPAI/Iliade Category to [-] for user '" & $idrh1 & "' , then validate with the button [ Apply ] ! (or close window to abort/exit)...", 1000, 500)
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
							If $remove = 0 Then
								$actiongroups = " [ Category groups LBPAI/Iliade ] :  [+]  " & @CRLF & $actiongroups
							Else
								$actiongroups = " [ Category groups LBPAI/Iliade ] :  [-]  " & @CRLF & $actiongroups
							EndIf
							GUIDelete($hgui2)
							ToolTip("Synchronizing with AD...", 5, 5)
							Sleep(3000)
							ToolTip("", 5, 5)
							ExitLoop
						Else
							MsgBox(4160, "warning !", "nothing selected... try again or close window")
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
			MsgBox(0, "Info !", "massive move only available for domain DCT !", 14)
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
			Return 0
		Else
			$sfileopendialog = StringReplace($sfileopendialog, "|", @CRLF)
			MsgBox(0, "", "fichier choisi : " & $sfileopendialog)
		EndIf
		$f = $sfileopendialog
		$f = StringSplit($f, "\")
		$i = $f[0]
		$file = FileOpen($f[$i], 0)
		Global $groupfile = $f[$i]
		If $file = -1 Then
			MsgBox(0, "Error", "Unable to open file." & $f[$i])
			Return 0
		EndIf
		$file = "" & $f[$i]
		$sizeof = FileGetSize($file)
		If $sizeof = 0 Then
			MsgBox(0, $group & ".txt", "no users to move in file: " & $group & @CRLF & "aborting !")
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
					ToolTip("User move:  " & $sobject & "  |  Target OU=" & $stargetou, 5, 5)
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
						FileWriteLine($out, $user & " | LBPAI/Iliade Category [+]" & @CRLF & $actiongroups & @CRLF & @CRLF)
					EndIf
					If $remove = 1 Then
						FileWriteLine($out, $user & " | LBPAI/Iliade Category [-]" & @CRLF & $actiongroups & @CRLF & @CRLF)
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
		Return 0
	EndFunc

	Func massive_date()
		$defautdc = $defautdcinit
		Dim $table
		Local Const $smessage = "fichier cible: 'Liste users à changer la date d'expiration" & "*" & ".txt'  à traiter !  (Utilisateurs à traiter, format ligne: Pabc123|jj/mm/aaaa )"
		Local $sfileopendialog = FileOpenDialog($smessage, @ScriptDir & "\", "Txt (*.txt)", $fd_filemustexist, "*" & ".txt")
		If @error Then
			MsgBox(0, "", "Aucun fichier choisi !")
			Return 0
		Else
			$sfileopendialog = StringReplace($sfileopendialog, "|", @CRLF)
			MsgBox(0, "", "fichier choisi : " & $sfileopendialog)
		EndIf
		$f = $sfileopendialog
		$f = StringSplit($f, "\")
		$i = $f[0]
		$file = FileOpen($f[$i], 0)
		Global $groupfile = $f[$i]
		If $file = -1 Then
			MsgBox(0, "Error", "Unable to open file." & $f[$i])
			Return 0
		EndIf
		$file = "" & $f[$i]
		$sizeof = FileGetSize($file)
		If $sizeof = 0 Then
			MsgBox(0, $group & ".txt", "no users to extend Date in file: " & $group & @CRLF & "aborting !")
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
					Return 0
				EndIf
			Else
				MsgBox(0, "warning !", "bad format txt file..." & @CRLF & "abort...")
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
		Return 0
	EndFunc

	Func massive_drive()
		$defautdc = $defautdcinit
		If $isdct = 0 Then
			MsgBox(0, "Warning !", "Only available for domain DCT !", 14)
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
						$outdrive = $outdrive & "  user  [ " & $idrh1[$z] & " ] ?" & @CRLF & @CRLF
					EndIf
				EndIf
			Next
			$historik = $historik & @CRLF & @CRLF & "  Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & "pasted in clipboard..." & @CRLF & @CRLF
			ClipPut($outdrive)
			GUICtrlSetData($aff, $historik, 1)
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
				Local $readad2 = "first call scanning all OUs"
			Else
				Local $readad2 = "already scanned all OUs"
			EndIf
			Global $hmain = GUICreate("Active Direcory OU Treeview - [" & $readad2 & "]", 743, 683, -1, -1)
			Global $htree = GUICtrlCreateTreeView(6, 6, 600, 666, -1, $ws_ex_clientedge)
			Global $bexpand = GUICtrlCreateButton("Expand", 624, 56, 97, 33)
			Global $bcollapse = GUICtrlCreateButton("Collapse", 624, 104, 97, 33)
			Global $bselect = GUICtrlCreateButton("Select OU", 624, 152, 97, 33)
		#EndRegion ### END Koda GUI section ###

		Global $atreeview = _ad_getoutreeview($sout, $htree, False, "", "%", 2)
		If @error <> 0 Then MsgBox(16, "Active Direcory OU Treeview", "Error creating list of OUs starting with '" & $sout & "'." & @CRLF & "Error returned by function _AD_GetALLOUs: @error = " & @error & ", @extended =  " & @extended)
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
			ToolTip("Scanning all OUs in Active Directory...", 5, 5)
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
;	FMUSICMEM_StopSong($hFMod)
;	FMUSICMEM_FreeSong($hFMod)
;	FSOUNDMEM_Close()
EndFunc

#EndRegion n2

#Region routines decompression

Func _WinAPI_Base64Decode2($sB64String)
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

func _Inferno()
   EndFunc