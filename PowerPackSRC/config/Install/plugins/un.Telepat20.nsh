;===========================================================================
; �������� ������������� ������� ������� 2.0 (��������� �������)

SubSection un."������� 2.0r"

	Section un."TelepatBeforeUnInstall"
		!insertmacro OC_STATUS "�������� �������� | ������� 2.0..."
	SectionEnd

	Section un."������, ������������"
	    Delete "$INSTDIR\config\telepat.dll"
	    Delete "$INSTDIR\config\scripts\telepat.vbs"
		Delete "$INSTDIR\config\docs\Telepat.history.txt"
		Delete "$INSTDIR\config\docs\Telepat.chm"
		Delete "$INSTDIR\telepat.icl"
		!insertmacro OC_USERFRIENDLY_DELETE "$INSTDIR\config\1cv7srct.st"
		!insertmacro OC_DEL_DOCFILE_FROM_STARTMENU "������� 2.0 - ������� ���������"
		!insertmacro OC_DEL_DOCFILE_FROM_STARTMENU "������� 2.0 - �������"
	SectionEnd

	Section un."������� xml2tls"
		Delete "$INSTDIR\config\system\xml2tls\1cpplang.xml"
		Delete "$INSTDIR\config\system\xml2tls\readme.txt"
		Delete "$INSTDIR\config\system\xml2tls\xml2tls.exe"
		!insertmacro OC_DEL_FILE_FROM_STARTMENU "������������\������� 2.0 - ������� xml2ts"
		RMDir "$INSTDIR\config\system\xml2tls"
	SectionEnd

SubSectionEnd