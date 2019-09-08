;===========================================================================
; �������� ��������� ������� ���������� (������ ������ aka Phoenix)

Section "���������� 1.3.3.3"

    SectionIn 1 2

	!insertmacro OC_STATUS "��������� �������� | ����������..."

    SetOutPath "$INSTDIR\config"
    File "${OC_ConfigDir}\SubSystems.dll"
    ;;File "${OC_ConfigDir}\SubSystems.tlb"
    File "${OC_ConfigDir}\system\ClipboardHistory.exe"

    ;; �������� ��� �������������, ��� ��������� "������", �������
	;; ���� ���� � ����������� ��� ����������, �� ����� ��� ��������
	IfFileExists "$INSTDIR\config\SubSystemsData.mdb" +2 0
		File "${OC_ConfigDir}\SubSystemsData.mdb"

    SetOutPath "$INSTDIR\config\scripts"
    File "${OC_ConfigDir}\scripts\SubSystemsManage.vbs"

	SetOutPath "$INSTDIR\config\docs"
	File "${OC_ConfigDir}\docs\SubSystems.readme.doc"
	!insertmacro OC_ADD_DOCFILE_TO_STARTMENU "SubSystems.readme.doc" "���������� - �������"
	File "${OC_ConfigDir}\docs\ClipboardHistory.readme.txt"
	!insertmacro OC_ADD_DOCFILE_TO_STARTMENU "ClipboardHistory.readme.txt" "ClipboardHistory - ��������"

SectionEnd