;===========================================================================
; �������� ��������� ������� RunAllIntellisence.vbs

Section "���������� Intellisence � Dots"

	SectionIn 1 2

	!insertmacro OC_STATUS "��������� �������� | ���������� Intellisence � Dots..."

	SetOutPath "$INSTDIR\config\scripts\Intellisence"
	File "${OC_ConfigDir}\scripts\Intellisence\RunAllIntellisence.vbs"

SectionEnd