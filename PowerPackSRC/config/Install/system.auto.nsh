;===========================================================================
;		������������� ������������� � ������� �������
; 			perl tools\gen_nsh.pl system.nsh.pl ..\system
; 		��� ��������� ��������� ��� ��������� ����� ��������!
; 		��������� ������� ������� � ������ system.nsh.pl.
;===========================================================================

;===========================================================================
;; �������� ��������� �����������

Section "���������� (COM-dll, WSC)"
	
	SectionIn 1 2 3
	
	!insertmacro OC_STATUS "��������� ����������� | ����������� ������..."
	
	;; ������� �� �����������
	SetOutPath "$INSTDIR\config\docs"
	File "${OC_ConfigDir}\docs\System.readme.txt"
	!insertmacro OC_ADD_DOCFILE_TO_STARTMENU "System.readme.txt" "���������� - ��������"

	SetOutPath "$INSTDIR\config\system"

	;; ������� fecho.exe ���������� ��������
	File "${OC_ConfigDir}\system\fecho.exe"

	;; COM-dll
	File "..\system\ArtWin.dll"
 	File "..\system\dynwrap.dll"
 	File "..\system\macrosenum.dll"
 	File "..\system\SelectDialog.dll"
 	File "..\system\SelectValue.dll"
 	File "..\system\svcsvc.dll"
 	File "..\system\TLBINF32.DLL"
 	File "..\system\WshExtra.dll"


	;; ���������� (*.wsc)
	File "..\system\1S.StatusIB.wsc"
 	File "..\system\Collections.wsc"
 	File "..\system\CommonServices.wsc"
 	File "..\system\OpenConf.RegistryIniFile.wsc"
 	File "..\system\Registry.wsc"
 	File "..\system\SyntaxAnalysis.wsc"


	;; ������� ��� ����������� � ������ ����������� �����������
	File "..\system\regfiles.js"
	File "..\system\regall.bat"
	File "..\system\unregall.bat"

	;; Dynamic Wrapper ��� �� ��������� Win9x
	SetOutPath "$INSTDIR\config\system\DynWin9x"
	File "${OC_ConfigDir}\system\DynWin9x\dynwrap.dll"
	File "${OC_ConfigDir}\system\DynWin9x\readme.txt"

	ClearErrors
	
	!insertmacro OC_STATUS "��������� ����������� | ����������� �����������..."
	
	;; ��������� ������ ����������� �����������
	ExecWait 'wscript.exe //nologo "$INSTDIR\config\system\regfiles.js" /I /S /L'
	
	;; ���� � ���������� ���������� ������� ��� �������� != 0, �� �������� �� ������
	IfErrors 0 end
		MessageBox MB_OK|MB_ICONEXCLAMATION \
			"��������� �� ����������� (*.wsc- ��� *.dll-�����)$\r$\n\
			�� ���� ���������������� ���������.$\r$\n\
			����� ���������� �� ���� ���������������� ����� ������ �� �����$\r$\n\
			$INSTDIR\config\system\regfiles.log.$\r$\n\
			��������� ���������� �� ���������� ���� ������ ����� ����� � �����$\r$\n\
			$INSTDIR\config\docs\System.readme.txt"
			
	end:
	
	!insertmacro OC_STATUS "��������� ����������� | ������"
	
SectionEnd ;; "����������"

;===========================================================================
;; �������� ������������� �����������

Section "un.���������� (COM-dll, WSC)"

	!insertmacro OC_STATUS "�������� ����������� | ������ �����������..."

	;; ��������� ������ ������ ����������� �����������
	IfFileExists "$INSTDIR\config\system\regfiles.js" 0 Delete_Files
		ExecWait 'wscript.exe //nologo "$INSTDIR\config\system\regfiles.js" /U /S'
		
	Delete_Files:

	!insertmacro OC_STATUS "�������� ����������� | �������� ������..."
	
	;; ������� ��������� �������
	Delete "$INSTDIR\config\system\regfiles.log"
	Delete "$INSTDIR\config\system\regfiles.js"
	Delete "$INSTDIR\config\system\regall.bat"
	Delete "$INSTDIR\config\system\unregall.bat"

	;; ������� fecho.exe ���������� ��������
	Delete "$INSTDIR\config\system\fecho.exe"
	
	;; COM-dll
	Delete "$INSTDIR\config\system\ArtWin.dll"
 	Delete "$INSTDIR\config\system\dynwrap.dll"
 	Delete "$INSTDIR\config\system\macrosenum.dll"
 	Delete "$INSTDIR\config\system\SelectDialog.dll"
 	Delete "$INSTDIR\config\system\SelectValue.dll"
 	Delete "$INSTDIR\config\system\svcsvc.dll"
 	Delete "$INSTDIR\config\system\TLBINF32.DLL"
 	Delete "$INSTDIR\config\system\WshExtra.dll"


	;; ���������� (*.wsc)
	Delete "$INSTDIR\config\system\1S.StatusIB.wsc"
 	Delete "$INSTDIR\config\system\Collections.wsc"
 	Delete "$INSTDIR\config\system\CommonServices.wsc"
 	Delete "$INSTDIR\config\system\OpenConf.RegistryIniFile.wsc"
 	Delete "$INSTDIR\config\system\Registry.wsc"
 	Delete "$INSTDIR\config\system\SyntaxAnalysis.wsc"


	Delete "$INSTDIR\config\system\DynWin9x\dynwrap.dll"
	Delete "$INSTDIR\config\system\DynWin9x\readme.txt"
	RMDir "$INSTDIR\config\system\DynWin9x"

	;; ������� �������
	Delete "$INSTDIR\config\docs\System.readme.txt"
	!insertmacro OC_DEL_DOCFILE_FROM_STARTMENU "���������� - ��������"

	;; ���� ���������� system ������, ������� ��
	RMDir "$INSTDIR\config\system"	

	!insertmacro OC_STATUS "�������� ����������� | ������"
	
SectionEnd ;; un."����������"
