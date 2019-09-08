;===========================================================================
; �������� ������������� ������� ���������� (������ ������ aka Phoenix)

Section un."����������"

	!insertmacro OC_STATUS "�������� �������� | ���������� ..."

    Delete "$INSTDIR\config\SubSystems.dll"
    Delete "$INSTDIR\config\SubSystems.tlb"
    Delete "$INSTDIR\config\ClipboardHistory.exe"

    Delete "$INSTDIR\config\scripts\SubSystemsManage.vbs"

	;; ����� � ������������� ����������� �������������� ������� �� �����, ������ � ����������
	!insertmacro OC_USERFRIENDLY_DELETE "$INSTDIR\config\SubSystemsData.mdb"

	;; ���� SubSystems.mdb ���-���� ��� ������, ������ � ������������� *.ldb-����
	IfFileExists "$INSTDIR\config\SubSystemsData.mdb" +2 0
		Delete "$INSTDIR\config\SubSystemsData.ldb"

	Delete "$INSTDIR\config\docs\SubSystems.readme.doc"
	!insertmacro OC_DEL_DOCFILE_FROM_STARTMENU "���������� - �������"

	Delete "$INSTDIR\config\docs\ClipboardHistory.readme.txt"
	!insertmacro OC_DEL_DOCFILE_FROM_STARTMENU "ClipboardHistory - ��������"

SectionEnd