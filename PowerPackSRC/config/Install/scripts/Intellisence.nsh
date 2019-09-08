;===========================================================================
; �������� ��������� ������� Intellisence.vbs

Section "Intellisence.vbs"

	SectionIn 1 2

	!insertmacro OC_STATUS "��������� �������� | Intellisence.vbs"

	SetOutPath "$INSTDIR\config\scripts\Intellisence"
	File "${OC_ConfigDir}\scripts\Intellisence\Intellisence.vbs"

	SetOutPath "$INSTDIR\config\docs"
	File "${OC_ConfigDir}\docs\Intellisence.readme.txt"
	!insertmacro OC_ADD_DOCFILE_TO_STARTMENU "Intellisence.readme.txt" "Intellisence - �������"

	CreateDirectory "$INSTDIR\config\Intell"
	SetOutPath "$INSTDIR\config\Intell"

	IfFileExists "$INSTDIR\config\telepat.dll" 0 no_telepat
		File /oname="intell.ini" "${OC_ConfigDir}\Intell\intell.ini"
		goto copy_ints
	no_telepat:
		File "${OC_ConfigDir}\Intell\intell_no_telepat.ini"

	copy_ints:
		!insertmacro OC_STATUS "��������� �������� | *.ints ����� ��� Intellisence.vbs"
		File "${OC_ConfigDir}\Intell\*.ints"

SectionEnd ;;Intellisence.vbs
