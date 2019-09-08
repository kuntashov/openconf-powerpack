;===========================================================================
;
;	Copyright (c) 2004,2005 OpenConf Community, http://openconf.itland.ru
;   
;	NSIS-������ ��� ������ ������������ OpenConf'�, �������� � ��������
;
;	������:
;    
;	��������� �������� aka a13x <kuntashov@yandex.ru> icq#338758861  
;
;===========================================================================

	!include "Sections.nsh"

	!include "parameters.nsh"

	!ifndef OC_VerFile
		!define OC_VerFile "_setup"
	!endif

	!ifndef OC_VerDisplay
		!define OC_VerDisplay ""
	!endif

;===========================================================================

	;�������� ���� ������������
	OutFile "${OC_ReleaseDir}\oc${OC_VerFile}.exe"
	;������������ ������� ������
	SetCompressor lzma

	;���� ���������
	InstType "������"
	InstType "�����������"
	InstType "�����������"

	;���������� ��������� �� ���������
	InstallDir "$PROGRAMFILES\1Cv77\BIN"
	InstallDirRegKey HKLM "Software\OpenConf\Install" ""

;===========================================================================

	;���������� Modern UI
	!include "MUI.nsh"

;===========================================================================

	Name "OpenConf"
	Caption "��������� OpenConf ${OC_VerDisplay} ��� 1�:����������� 7.7"

;===========================================================================
;���������� ����������

	;;Var MUI_TEMP
	Var STARTMENU_FOLDER
	Var hwnd
	Var UPDATE

;===========================================================================
;��������������� ������� � �������

	!include "utils.nsh"

;===========================================================================
;���������

!define MUI_ABORTWARNING

;FIXME ���� ��������� ��������� ���������� ���� � ������� ��� OC � ����������
;!define MUI_HEADERIMAGE
;!define MUI_WELCOMEFINISHPAGE_BITMAP "${OC_HeaderBitmap}"

;FIXME ��������
!define MUI_COMPONENTSPAGE_NODESC
;!define MUI_COMPONENTSPAGE_SMALLDESC

;������ ��������
!define MUI_WELCOMEPAGE_TITLE "��� ������������ ������ ��������� OpenConf ${OC_VerDisplay}!"
;; TODO ��������� ����� �� �������� ����� (���� �����)
!define MUI_WELCOMEPAGE_TEXT \
		"���� ������ ������� ��� ���������� OpenConf ${OC_VerDisplay} \
		��� 1�:����������� 7.7.\r\n\r\n  \
		OpenConf - ��� ����������� �������, ������������� ����������� ���������, ����������� \
		����������� ����������� ������������� 1�:�����������.\r\n\r\n  \
		OpenConf ������������� �������� - ��������������� ���������, ����������� \
		���������������� �������� ������ �� ���������������� � ������� �������� �� \
		����� ������, �������������� ���������� Active Scripting (��������, VBScript \
		� JScript) � �������� �� ������ �������� ������ (�++, Delphi � �.�.) � \
		�������������� ���������� COM.\r\n\r\n  $_CLICK"
!insertmacro MUI_PAGE_WELCOME

;��������
!ifdef OC_LicenseFile
	!insertmacro MUI_PAGE_LICENSE "${OC_LicenseFile}"
!endif

;����������� �������� ������ �������� ���������
Page custom MyDirectoryPageEnter MyDirectoryPageLeave

!insertmacro MUI_PAGE_COMPONENTS

;�������� ������ ����������� ������ (� ���� ����)
!ifdef OC_StartMenuPage
	!define MUI_STARTMENUPAGE_REGISTRY_ROOT "HKCU"
	!define MUI_STARTMENUPAGE_REGISTRY_KEY "Software\1C\1Cv7\7.7\OpenConf"
	!define MUI_STARTMENUPAGE_REGISTRY_VALUENAME "Start Menu Folder"
	!insertmacro MUI_PAGE_STARTMENU Application $STARTMENU_FOLDER
!endif

;�������� ��������� ���������
!insertmacro MUI_PAGE_INSTFILES

;��������� �������� ("������")
!define MUI_FINISHPAGE_LINK "�������� ���-���� ������� �������� ������������!"
!define MUI_FINISHPAGE_LINK_LOCATION "http://openconf.itland.ru/"
!define MUI_FINISHPAGE_NOREBOOTSUPPORT ;��� ����� ���� ������������
!insertmacro MUI_PAGE_FINISH

!ifdef OC_Uninstall
	;�������������
	!insertmacro MUI_UNPAGE_CONFIRM
	!insertmacro MUI_UNPAGE_INSTFILES
!endif

;===========================================================================

;Russian only, sorry
!insertmacro MUI_LANGUAGE "Russian"

;===========================================================================
;�������� ������

Section "OpenConf ${OC_VerDisplay} (�����������)" SecOpenConf

	SectionIn 1 2 3 RO

	SetOutPath "$INSTDIR"

	!insertmacro OC_STATUS "��������� OpenConf..."

	;;������� bin\config � bin\config\scripts
	CreateDirectory "$INSTDIR\config"
	CreateDirectory "$INSTDIR\config\docs"
	CreateDirectory "$INSTDIR\config\scripts"

	;;�������� ������������ 1�'������ config.dll � bin\config,
	;;���� ������ ��� �� ����� ����������
	StrCmp $UPDATE "1" core_files 0
		!insertmacro OC_STATUS "��������� OpenConf | ����������� config.dll..."
		CopyFiles "$INSTDIR\config.dll" "$INSTDIR\config"

	core_files:
		!insertmacro OC_STATUS "��������� OpenConf | ���������� ������..."
		File "${OC_BinDir}\config.dll"
		;; TODO ��������� � �������������� ������ �� tlb � �������
		File "${OC_BinDir}\config.tlb"
		File "${OC_BinDir}\OpenConf.chm"

	;; ���� �� ������� ������������ ���, ��� ������� �� OpenConf'� 
	;; [� �������� docs] ���, �� �� �����-�� ����, ��� � ������, ��� ����
	;; [� �������� BIN, ��� ���� ������� OpenConf], �� �������� � docs 
	;; ���� �� ����� �� ���
	CreateShortCut "$INSTDIR\config\docs\OpenConf.lnk" "$INSTDIR\OpenConf.chm"

	;; readme- � history- �����
	SetOutPath "$INSTDIR\config\docs"
	File "${OC_ResDir}\readme.txt"
	File "${OC_ConfigDir}\docs\OpenConf.history.txt"
    File "${OC_ConfigDir}\docs\WhatsNew.txt"

	;��� �������� ��������� �������� � �������
	WriteRegStr HKLM "Software\1C\1Cv7\7.7\OpenConf\Install" "" $INSTDIR

	!ifdef OC_Uninstall
		; ������� �����������
		WriteUninstaller "$INSTDIR\config\Uninstall.exe"
	!endif

	!ifdef OC_StartMenuPage
		!insertmacro OC_STATUS "��������� OpenConf | �������� ����������� ������ � �������..."
		!insertmacro MUI_STARTMENU_WRITE_BEGIN Application
		; ������
    	CreateDirectory "$SMPROGRAMS\$STARTMENU_FOLDER"
		CreateDirectory "$SMPROGRAMS\$STARTMENU_FOLDER\������������"
		; ��� ������� �� ��������� � readme-�����
		CreateShortCut	"$SMPROGRAMS\$STARTMENU_FOLDER\������� �� OpenConf.lnk" \
						"$INSTDIR\OpenConf.chm"
		CreateShortCut	"$SMPROGRAMS\$STARTMENU_FOLDER\ReadMe.lnk" \
						"$INSTDIR\config\docs\readme.txt"
		CreateShortCut	"$SMPROGRAMS\$STARTMENU_FOLDER\������� ��������� OpenConf.lnk" \
						"$INSTDIR\config\docs\OpenConf.history.txt"
		CreateShortCut	"$SMPROGRAMS\$STARTMENU_FOLDER\��� ������.lnk" \
						"$INSTDIR\config\docs\WhatsNew.txt"
		!ifdef OC_Uninstall
			;��� �������������
			CreateShortCut	"$SMPROGRAMS\$STARTMENU_FOLDER\�������� OpenConf.lnk" \
							"$INSTDIR\config\Uninstall.exe"
		!endif
    	!insertmacro MUI_STARTMENU_WRITE_END	
	!endif

SectionEnd ;; OpenConf

;===========================================================================
;�������������� ������

;; �������
!ifdef OC_Plugins
	!include "${OC_Plugins}"
!endif

;; �������
!ifdef OC_Scripts
	!include "${OC_Scripts}"
!endif

;; ������������ ���������� � ������ ������
!ifdef OC_AddOns
	!include "${OC_AddOns}"
!endif

;; �������������� ������������ (�������, ���������� etc)
!ifdef OC_Docs
	!include "${OC_Docs}"
!endif

;; �������������� ���������� (��� ��������)
!ifdef OC_System
	!include "${OC_System}"
!endif

;===========================================================================
;������ ��������������
!ifdef OC_Uninstall
Section "Uninstall"

	!insertmacro OC_STATUS "�������� OpenConf..."

	;; ���������, � �������� �� ���? (�� ������, ���� ����� ���-��
	;; ��� �������� ������������� ��������'� �������)
	GetDllVersion "$INSTDIR\config.dll" $R0 $R1
	IntOp $R2 $R0 / 0x00010000
	IntOp $R3 $R0 & 0x0000FFFF
	IntOp $R4 $R1 / 0x00010000
	IntOp $R5 $R1 & 0x0000FFFF	

	;; ���� ������ ������.��� ���� 7�, �� �������, ��� ��� ��������
	IntCmpU $R2 7 delete_files 0 delete_files

		!insertmacro OC_STATUS "�������� OpenConf | �������� ������..."
		Delete "$INSTDIR\config.dll"
		Delete "$INSTDIR\config.tlb"

		!insertmacro OC_STATUS "�������� OpenConf | ����������� config.dll..."
		CopyFiles "$INSTDIR\config\config.dll" "$INSTDIR"
		Delete "$INSTDIR\config\config.dll"
  
	delete_files:

	!insertmacro OC_STATUS "�������� OpenConf | �������� ������..."
	;; ������ ��������� ����� (�������, ����� � ��� �����������)
	Delete "$INSTDIR\OpenConf.chm"
	Delete "$INSTDIR\config\docs\OpenConf.history.txt"
	Delete "$INSTDIR\config\docs\readme.txt"
	Delete "$INSTDIR\config\docs\OpenConf.lnk"
	Delete "$INSTDIR\config\docs\WhatsNew.txt"
	Delete "$INSTDIR\config\Uninstall.exe"

	;; ���� ����������� ����������� ������ � ���� ����, �� ���� ���� �������
	!ifdef OC_StartMenuPage		
		!insertmacro OC_STATUS "�������� OpenConf | �������� ������� � ����������� ������..."
		Delete	"$SMPROGRAMS\$STARTMENU_FOLDER\������� ��������� OpenConf.lnk"
		Delete	"$SMPROGRAMS\$STARTMENU_FOLDER\ReadMe.lnk"
		Delete	"$SMPROGRAMS\$STARTMENU_FOLDER\������� �� OpenConf.lnk"
		Delete	"$SMPROGRAMS\$STARTMENU_FOLDER\�������� OpenConf.lnk"	
		Delete	"$SMPROGRAMS\$STARTMENU_FOLDER\��� ������.lnk"
		;; �������� ������� ���� ������ (���� ������ ��� ��������� �����)
		RMDir "$SMPROGRAMS\$STARTMENU_FOLDER\������������"
		RMDir "$SMPROGRAMS\$STARTMENU_FOLDER"		
	!endif

	;; ������� ������� (���� ������ ����� ��������� �����)
	RMDir "$INSTDIR\config\docs"
	RMDir "$INSTDIR\config\scripts"
	RMDir "$INSTDIR\config"

	!insertmacro OC_STATUS "�������� OpenConf | �������� ������ � �������..."

	DeleteRegKey /ifempty HKLM "Software\1C\1Cv7\7.7\OpenConf\Install"
	DeleteRegKey /ifempty HKLM "Software\1C\1Cv7\7.7\OpenConf"

	!insertmacro OC_STATUS "������"

SectionEnd ;; Uninstall
!endif

;===========================================================================
;����������� �������

Function .onInit
	!insertmacro MUI_INSTALLOPTIONS_EXTRACT_AS "OpenConf.ini" "OpenConf.ini"
FunctionEnd

;===========================================================================

Function un.onInit
	;; ������ ���������� � ���� ��������� �� �������
	;; FIXME ����� �� ��������� ���������� �����������?
	ReadRegStr $INSTDIR HKLM "Software\1C\1Cv7\7.7\OpenConf\Install" ""
	!ifdef OC_StartMenuPage	
		!insertmacro MUI_STARTMENU_GETFOLDER Application $STARTMENU_FOLDER
	!endif
FunctionEnd

;===========================================================================
;; �������� ������������ ������� ������ �������� ���������

LangString SEC_DIR_TITLE ${LANG_RUSSIAN} "����� �������� ��������� OpenConf"
LangString SEC_DIR_SUBTITLE ${LANG_RUSSIAN} "�������� ����������� 1�:�����������, � \
	������� ������� ���������� OpenConf ��� ������� ������� ��������� �������"

Function MyDirectoryPageEnter
	
	push $R0
	push $R1
	push $R2
	push $R3

	StrCpy $R0 ""
	StrCpy $R1 "0"
	StrCpy $R2 ""
	StrCpy $R3 ""

	;; ����� ����, �� ������ ��������� �� ������ "�����" �� ��������� ��������?
	ReadINIStr $R0 "$PLUGINSDIR\OpenConf.ini" "Field 7" "State"
	StrCmp $R0 "" 0 init_page

	;; �������� ��������� ������ �� ��������� 1� �� �������
	read_next_key:
		EnumRegKey $R0 HKLM "SOFTWARE\1C\1Cv7\7.7" "$R1"
		IntOp $R1 $R1 + 1
		;; ��������� ������� �� ������ �������� � ������
		StrCmp $R0 "Components" read_next_key ;; ��������� ���������
		StrCmp $R0 "OpenConf" read_next_key ;; ���������, ����������� �����������
		StrCmp $R0 "" end_loop
		StrCmp $R2 "" 0 +4
		StrCpy $R2 $R0
		StrCpy $R3 $R0
		goto +2
		StrCpy $R2 "$R2|$R0"
		goto read_next_key
	end_loop:

	WriteINIStr "$PLUGINSDIR\OpenConf.ini" "Field 4" "ListItems" $R2
         
	;; ��������� ����������� �� ���������
	WriteINIStr "$PLUGINSDIR\OpenConf.ini" "Field 4" "State" $R3

	;; ����� �������� ��������� ������� - �������� �� ���������
	ReadRegStr $R3 HKLM "SOFTWARE\1C\1Cv7\7.7\$R3" "1CPath"
	push $R3
	call GetParent
	pop $R3
	WriteINIStr "$PLUGINSDIR\OpenConf.ini" "Field 7" "State" $R3

	pop $R3
	pop $R2
	pop $R1
	pop $R0

	init_page:	

	!insertmacro MUI_HEADER_TEXT "$(SEC_DIR_TITLE)" "$(SEC_DIR_SUBTITLE)"
	!insertmacro MUI_INSTALLOPTIONS_INITDIALOG "OpenConf.ini"

	pop $hwnd ; ��������� ����� ����, �� ��� ����� �����

	!insertmacro MUI_INSTALLOPTIONS_SHOW

FunctionEnd ;; MyDirectoryPageEnter

;===========================================================================

Function MyDirectoryPageLeave

	ReadINIStr $0 "$PLUGINSDIR\OpenConf.ini" "Settings" "State"
	StrCmp $0 0 onValidate
	StrCmp $0 1 onRadioButtonClick
	StrCmp $0 2 onRadioButtonClick
	StrCmp $0 4 onListChange
	StrCmp $0 7 onDirRequestChange
	Abort

	onListChange:
	Abort

	onRadioButtonClick:
		IntOp $0 $0 - 1
		GetDlgItem $1 $hwnd 1206 ; DirRequest  (1200 + Field 7 - 1)
		EnableWindow $1 $0
		GetDlgItem $1 $hwnd 1207 ; ... ������
		EnableWindow $1 $0
		IntOp $0 1 - $0
		GetDlgItem $1 $hwnd 1203 ; DropList  (1200 + Field 4 - 1 )
		EnableWindow $1 $0

		;; ���������� ��������� ��������� �� �����, ����� ����������� �� ��������� ��������
		;; �� ������ "�����" ��� ��������� ����� ��� ��, ��� ����� ���, ��� �� �� ��� �������

		;; FIXME ������ ������ ��������! ������������ ���������� ��������� ����� (StrCmp � goto)!

		StrCmp $0 1 0 +8

		WriteINIStr "$PLUGINSDIR\OpenConf.ini" "Field 1" "State" "1"
		WriteINIStr "$PLUGINSDIR\OpenConf.ini" "Field 2" "State" ""
		WriteINIStr "$PLUGINSDIR\OpenConf.ini" "Field 4" "Flags" "NOTIFY"
		WriteINIStr "$PLUGINSDIR\OpenConf.ini" "Field 7" "Flags" "DISABLED|NOTIFY"

		goto +6
	
		WriteINIStr "$PLUGINSDIR\OpenConf.ini" "Field 1" "State" ""
		WriteINIStr "$PLUGINSDIR\OpenConf.ini" "Field 2" "State" "1"
		WriteINIStr "$PLUGINSDIR\OpenConf.ini" "Field 4" "Flags" "DISABLED"
		WriteINIStr "$PLUGINSDIR\OpenConf.ini" "Field 7" "Flags" "NOTIFY"

	  Abort

	onDirRequestChange:
	Abort

	onValidate:
		ReadINIStr $0 "$PLUGINSDIR\OpenConf.ini" "Field 1" "State"
		StrCmp $0 1 0 read_reqdir_field

		ReadINIStr $0 "$PLUGINSDIR\OpenConf.ini" "Field 4" "State"
		ReadRegStr $0 HKLM "SOFTWARE\1C\1Cv7\7.7\$0" "1CPath"
		push $0
		call GetParent
		pop $0
		goto validate_dir

	read_reqdir_field:
		ReadINIStr $0 "$PLUGINSDIR\OpenConf.ini" "Field 7" "State"

	validate_dir:
		;; ���������, �������� �� �������� ���������� ��������� ��������� 1�
		IfFileExists "$0\1cv7*.exe" +3 0
		MessageBox MB_OK|MB_ICONEXCLAMATION "��������� ������� �� �������� ��������� ��������� 1�!"
		Abort

		;; ������� ���������
		StrCpy $INSTDIR $0

		;; ��������� ������ 1���
		;; �� ��������� ��������� �� �����, �.�. �� ����� ��� ������ ���,
		;; � �� ������.��� - ������ ��� ��������, ��� �������� ��� ����������

		GetDllVersion "$0\basic.dll" $R0 $R1
		IntOp $R2 $R0 / 0x00010000
		IntOp $R3 $R0 & 0x0000FFFF
		IntOp $R4 $R1 / 0x00010000
		IntOp $R5 $R1 & 0x0000FFFF
		IntCmpU $R2 7  +1 +3 +3
		IntCmpU $R3 70 +1 +2 +1
		IntCmpU $R5 14 version_ok +1 version_ok
		MessageBox MB_OK|MB_ICONEXCLAMATION "��� ������ ��������� ��������� 1� 7.7 �� ���� 14-�� ������ (7.70.014)!"
		Abort

    version_ok:
		;; ��������� ������ ��������� (� ����� � ��, ���������� �� ��)
		GetDllVersion "$0\config.dll" $R0 $R1
		IntOp $R2 $R0 / 0x00010000
		IntOp $R3 $R0 & 0x0000FFFF
		IntOp $R4 $R1 / 0x00010000
		IntOp $R5 $R1 & 0x0000FFFF

		GetDLLVersionLocal "${OC_ConfigDir}\config.dll" $0 $1
		IntOp $2 $0 / 0x00010000
		IntOp $3 $0 & 0x0000FFFF
		IntOp $4 $1 / 0x00010000
		IntOp $5 $1 & 0x0000FFFF

		StrCpy $UPDATE "0"
		;; ���� ������ ������.��� ���� 7�, �� �������, ��� ��� ��������
		IntCmpU $R2 "7" end 0 end
		;; �������� ����������, ��������, ���� �� ��� ���������
		IntCmpU $R2 $2 0 update update_warning
		IntCmpU $R3 $3 0 update update_warning
		IntCmpU $R4 $4 0 update update_warning
		IntCmpU $R5 $5 update_warning 0 update_warning

	update:
		StrCpy $UPDATE "1"
		goto end

	update_warning:
		;; FIXME ???
		MessageBox MB_OK|MB_ICONEXCLAMATION "���������� ��������� ��������� ����� �� ��� ����� ������� ������!"	 
		!insertmacro UnselectSection "0" ;; ������������ ����������������� ����

	end:

FunctionEnd ;; MyDirectoryPageLeave

;===========================================================================
;; EOF
