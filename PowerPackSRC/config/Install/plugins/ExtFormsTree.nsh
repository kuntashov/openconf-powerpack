;===========================================================================
; �������� ��������� ������� ExtFormsTree (������ ������ aka Phoenix)

Section "ExtForms Tree 1.2.1.2"

    SectionIn 1 2

	!insertmacro OC_STATUS "��������� �������� | ExtFormsTree..."

    SetOutPath "$INSTDIR\config"
    File "${OC_ConfigDir}\ExtFormsTree.dll"
    File "${OC_ConfigDir}\ExtFormsTree.tlb"

    ;; �������� ��� �������������, ��� ��������� "������", �������
	;; ���� ���� � ����������� ��� ����������, �� ����� ��� ��������
	IfFileExists "$INSTDIR\config\ExtFormsTree.txt" +2 0
		File "${OC_ConfigDir}\ExtFormsTree.txt"

    SetOutPath "$INSTDIR\config\scripts"
    File "${OC_ConfigDir}\scripts\ExtFormsTreeManage.vbs"

	SetOutPath "$INSTDIR\config\docs"
	File "${OC_ConfigDir}\docs\ExtFormsTree.readme.txt"
	!insertmacro OC_ADD_DOCFILE_TO_STARTMENU "ExtFormsTree.readme.txt" "ExtFormsTree - ��������"

SectionEnd