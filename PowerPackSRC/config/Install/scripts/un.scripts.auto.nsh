;===========================================================================
;		������������� ������������� � ������� �������
; 			perl tools\scripts.nsh.pl -u -d ..\scripts -f scripts\filters.txt -i scripts
; 		��� ��������� ��������� ��� ��������� ����� ��������!
; 		��������� ������� ������� � ���� tools\scripts.nsh.pl.
;===========================================================================


SubSection un."�"
	Section "un.AfterUnInstall�"
		RMDir "$INSTDIR\config\scripts\�"
	SectionEnd
SubSectionEnd ;; un.�
SubSection un."��������������"
	!include "scripts\un.author.nsh"
	Section "un.Brackets.vbs"
		!insertmacro OC_STATUS "�������� �������� | Brackets.vbs ..."
		Delete "$INSTDIR\config\scripts\��������������\Brackets.vbs"
	SectionEnd ;; un.Brackets.vbs
	Section "un.mb_utils.js"
		!insertmacro OC_STATUS "�������� �������� | mb_utils.js ..."
		Delete "$INSTDIR\config\scripts\��������������\mb_utils.js"
	SectionEnd ;; un.mb_utils.js
	Section "un.MultiString.js"
		!insertmacro OC_STATUS "�������� �������� | MultiString.js ..."
		Delete "$INSTDIR\config\scripts\��������������\MultiString.js"
	SectionEnd ;; un.MultiString.js
	Section "un.refactorer.js"
		!insertmacro OC_STATUS "�������� �������� | refactorer.js ..."
		Delete "$INSTDIR\config\scripts\��������������\refactorer.js"
	SectionEnd ;; un.refactorer.js
	Section "un.RTrimModule.vbs"
		!insertmacro OC_STATUS "�������� �������� | RTrimModule.vbs ..."
		Delete "$INSTDIR\config\scripts\��������������\RTrimModule.vbs"
	SectionEnd ;; un.RTrimModule.vbs
	Section "un.������ ����.vbs"
		!insertmacro OC_STATUS "�������� �������� | ������ ����.vbs ..."
		Delete "$INSTDIR\config\scripts\��������������\������ ����.vbs"
	SectionEnd ;; un.������ ����.vbs
	Section "un.���������� ������ � ����� ������.vbs"
		!insertmacro OC_STATUS "�������� �������� | ���������� ������ � ����� ������.vbs ..."
		Delete "$INSTDIR\config\scripts\��������������\���������� ������ � ����� ������.vbs"
	SectionEnd ;; un.���������� ������ � ����� ������.vbs
	Section "un.�������������� ������.vbs"
		!insertmacro OC_STATUS "�������� �������� | �������������� ������.vbs ..."
		Delete "$INSTDIR\config\scripts\��������������\�������������� ������.vbs"
	SectionEnd ;; un.�������������� ������.vbs
	Section "un.AfterUnInstall��������������"
		RMDir "$INSTDIR\config\scripts\��������������"
	SectionEnd
SubSectionEnd ;; un.��������������
SubSection un."������"
	Section "un.1C++.vbs"
		!insertmacro OC_STATUS "�������� �������� | 1C++.vbs ..."
		Delete "$INSTDIR\config\scripts\������\1C++.vbs"
	SectionEnd ;; un.1C++.vbs
	Section "un.DocPath.js"
		!insertmacro OC_STATUS "�������� �������� | DocPath.js ..."
		Delete "$INSTDIR\config\scripts\������\DocPath.js"
	SectionEnd ;; un.DocPath.js
	Section "un.ParseCmdLineInConfig.vbs"
		!insertmacro OC_STATUS "�������� �������� | ParseCmdLineInConfig.vbs ..."
		Delete "$INSTDIR\config\scripts\������\ParseCmdLineInConfig.vbs"
	SectionEnd ;; un.ParseCmdLineInConfig.vbs
	Section "un.SetWinCaption.vbs"
		!insertmacro OC_STATUS "�������� �������� | SetWinCaption.vbs ..."
		Delete "$INSTDIR\config\scripts\������\SetWinCaption.vbs"
	SectionEnd ;; un.SetWinCaption.vbs
	Section "un.TelepatSettings.js"
		!insertmacro OC_STATUS "�������� �������� | TelepatSettings.js ..."
		Delete "$INSTDIR\config\scripts\������\TelepatSettings.js"
	SectionEnd ;; un.TelepatSettings.js
	Section "un.TurboMD_Artur.vbs"
		!insertmacro OC_STATUS "�������� �������� | TurboMD_Artur.vbs ..."
		Delete "$INSTDIR\config\scripts\������\TurboMD_Artur.vbs"
	SectionEnd ;; un.TurboMD_Artur.vbs
	Section "un.���� ���� ��������.vbs"
		!insertmacro OC_STATUS "�������� �������� | ���� ���� ��������.vbs ..."
		Delete "$INSTDIR\config\scripts\������\���� ���� ��������.vbs"
	SectionEnd ;; un.���� ���� ��������.vbs
	!include "scripts\un.���� �������� �� �����.nsh"
	Section "un.���������� ��������.vbs"
		!insertmacro OC_STATUS "�������� �������� | ���������� ��������.vbs ..."
		Delete "$INSTDIR\config\scripts\������\���������� ��������.vbs"
	SectionEnd ;; un.���������� ��������.vbs
	Section "un.���������� �������� ����.vbs"
		!insertmacro OC_STATUS "�������� �������� | ���������� �������� ����.vbs ..."
		Delete "$INSTDIR\config\scripts\������\���������� �������� ����.vbs"
	SectionEnd ;; un.���������� �������� ����.vbs
	Section "un.������������������1C.vbs"
		!insertmacro OC_STATUS "�������� �������� | ������������������1C.vbs ..."
		Delete "$INSTDIR\config\scripts\������\������������������1C.vbs"
	SectionEnd ;; un.������������������1C.vbs
	Section "un.AfterUnInstall������"
		RMDir "$INSTDIR\config\scripts\������"
	SectionEnd
SubSectionEnd ;; un.������
SubSection un."��������������"
	Section "un.������������ ��������� �������.vbs"
		!insertmacro OC_STATUS "�������� �������� | ������������ ��������� �������.vbs ..."
		Delete "$INSTDIR\config\scripts\��������������\������������ ��������� �������.vbs"
	SectionEnd ;; un.������������ ��������� �������.vbs
	Section "un.�������� ����.js"
		!insertmacro OC_STATUS "�������� �������� | �������� ����.js ..."
		Delete "$INSTDIR\config\scripts\��������������\�������� ����.js"
	SectionEnd ;; un.�������� ����.js
	Section "un.������� ������ �� �����.vbs"
		!insertmacro OC_STATUS "�������� �������� | ������� ������ �� �����.vbs ..."
		Delete "$INSTDIR\config\scripts\��������������\������� ������ �� �����.vbs"
	SectionEnd ;; un.������� ������ �� �����.vbs
	Section "un.������� ��������� � ������ �� �����.vbs"
		!insertmacro OC_STATUS "�������� �������� | ������� ��������� � ������ �� �����.vbs ..."
		Delete "$INSTDIR\config\scripts\��������������\������� ��������� � ������ �� �����.vbs"
	SectionEnd ;; un.������� ��������� � ������ �� �����.vbs
	Section "un.������� ����.js"
		!insertmacro OC_STATUS "�������� �������� | ������� ����.js ..."
		Delete "$INSTDIR\config\scripts\��������������\������� ����.js"
	SectionEnd ;; un.������� ����.js
	Section "un.AfterUnInstall��������������"
		RMDir "$INSTDIR\config\scripts\��������������"
	SectionEnd
SubSectionEnd ;; un.��������������
SubSection un."�����"
	Section "un.������� ���� ���������.vbs"
		!insertmacro OC_STATUS "�������� �������� | ������� ���� ���������.vbs ..."
		Delete "$INSTDIR\config\scripts\�����\������� ���� ���������.vbs"
	SectionEnd ;; un.������� ���� ���������.vbs
	Section "un.����������.js"
		!insertmacro OC_STATUS "�������� �������� | ����������.js ..."
		Delete "$INSTDIR\config\scripts\�����\����������.js"
	SectionEnd ;; un.����������.js
	Section "un.��������.vbs"
		!insertmacro OC_STATUS "�������� �������� | ��������.vbs ..."
		Delete "$INSTDIR\config\scripts\�����\��������.vbs"
	SectionEnd ;; un.��������.vbs
	Section "un.AfterUnInstall�����"
		RMDir "$INSTDIR\config\scripts\�����"
	SectionEnd
SubSectionEnd ;; un.�����
SubSection un."���������"
	Section "un.FindText.vbs"
		!insertmacro OC_STATUS "�������� �������� | FindText.vbs ..."
		Delete "$INSTDIR\config\scripts\���������\FindText.vbs"
	SectionEnd ;; un.FindText.vbs
	Section "un.jumper.vbs"
		!insertmacro OC_STATUS "�������� �������� | jumper.vbs ..."
		Delete "$INSTDIR\config\scripts\���������\jumper.vbs"
	SectionEnd ;; un.jumper.vbs
	Section "un.NavigationTools.vbs"
		!insertmacro OC_STATUS "�������� �������� | NavigationTools.vbs ..."
		Delete "$INSTDIR\config\scripts\���������\NavigationTools.vbs"
	SectionEnd ;; un.NavigationTools.vbs
	!include "scripts\un.navigator.nsh"
	Section "un.ScriptMethodList.vbs"
		!insertmacro OC_STATUS "�������� �������� | ScriptMethodList.vbs ..."
		Delete "$INSTDIR\config\scripts\���������\ScriptMethodList.vbs"
	SectionEnd ;; un.ScriptMethodList.vbs
	Section "un.���������.vbs"
		!insertmacro OC_STATUS "�������� �������� | ���������.vbs ..."
		Delete "$INSTDIR\config\scripts\���������\���������.vbs"
	SectionEnd ;; un.���������.vbs
	Section "un.������� ���� �� ��������� ����������������.vbs"
		!insertmacro OC_STATUS "�������� �������� | ������� ���� �� ��������� ����������������.vbs ..."
		Delete "$INSTDIR\config\scripts\���������\������� ���� �� ��������� ����������������.vbs"
	SectionEnd ;; un.������� ���� �� ��������� ����������������.vbs
	Section "un.�������� �� ������.vbs"
		!insertmacro OC_STATUS "�������� �������� | �������� �� ������.vbs ..."
		Delete "$INSTDIR\config\scripts\���������\�������� �� ������.vbs"
	SectionEnd ;; un.�������� �� ������.vbs
	Section "un.AfterUnInstall���������"
		RMDir "$INSTDIR\config\scripts\���������"
	SectionEnd
SubSectionEnd ;; un.���������
SubSection un."������������"
	Section "un.������������ ���������.vbs"
		!insertmacro OC_STATUS "�������� �������� | ������������ ���������.vbs ..."
		Delete "$INSTDIR\config\scripts\������������\������������ ���������.vbs"
	SectionEnd ;; un.������������ ���������.vbs
	Section "un.������������ ����������.vbs"
		!insertmacro OC_STATUS "�������� �������� | ������������ ����������.vbs ..."
		Delete "$INSTDIR\config\scripts\������������\������������ ����������.vbs"
	SectionEnd ;; un.������������ ����������.vbs
	Section "un.������������ ��������.vbs"
		!insertmacro OC_STATUS "�������� �������� | ������������ ��������.vbs ..."
		Delete "$INSTDIR\config\scripts\������������\������������ ��������.vbs"
	SectionEnd ;; un.������������ ��������.vbs
	Section "un.������������ ���������������� ��������.vbs"
		!insertmacro OC_STATUS "�������� �������� | ������������ ���������������� ��������.vbs ..."
		Delete "$INSTDIR\config\scripts\������������\������������ ���������������� ��������.vbs"
	SectionEnd ;; un.������������ ���������������� ��������.vbs
	Section "un.������������ ������������.vbs"
		!insertmacro OC_STATUS "�������� �������� | ������������ ������������.vbs ..."
		Delete "$INSTDIR\config\scripts\������������\������������ ������������.vbs"
	SectionEnd ;; un.������������ ������������.vbs
	Section "un.������������ ��.vbs"
		!insertmacro OC_STATUS "�������� �������� | ������������ ��.vbs ..."
		Delete "$INSTDIR\config\scripts\������������\������������ ��.vbs"
	SectionEnd ;; un.������������ ��.vbs
	Section "un.������� ��������.vbs"
		!insertmacro OC_STATUS "�������� �������� | ������� ��������.vbs ..."
		Delete "$INSTDIR\config\scripts\������������\������� ��������.vbs"
	SectionEnd ;; un.������� ��������.vbs
	Section "un.AfterUnInstall������������"
		RMDir "$INSTDIR\config\scripts\������������"
	SectionEnd
SubSectionEnd ;; un.������������
SubSection un."������������������"
	Section "un.cvs.vbs"
		!insertmacro OC_STATUS "�������� �������� | cvs.vbs ..."
		Delete "$INSTDIR\config\scripts\������������������\cvs.vbs"
	SectionEnd ;; un.cvs.vbs
	Section "un.extforms.vbs"
		!insertmacro OC_STATUS "�������� �������� | extforms.vbs ..."
		Delete "$INSTDIR\config\scripts\������������������\extforms.vbs"
	SectionEnd ;; un.extforms.vbs
	Section "un.�������� ������.vbs"
		!insertmacro OC_STATUS "�������� �������� | �������� ������.vbs ..."
		Delete "$INSTDIR\config\scripts\������������������\�������� ������.vbs"
	SectionEnd ;; un.�������� ������.vbs
	Section "un.��������������.vbs"
		!insertmacro OC_STATUS "�������� �������� | ��������������.vbs ..."
		Delete "$INSTDIR\config\scripts\������������������\��������������.vbs"
	SectionEnd ;; un.��������������.vbs
	Section "un.AfterUnInstall������������������"
		RMDir "$INSTDIR\config\scripts\������������������"
	SectionEnd
SubSectionEnd ;; un.������������������
SubSection un."MD_Tasks"
	Section "un.autoload.vbs"
		!insertmacro OC_STATUS "�������� �������� | autoload.vbs ..."
		Delete "$INSTDIR\config\scripts\MD_Tasks\autoload.vbs"
	SectionEnd ;; un.autoload.vbs
	Section "un.ExtractAllReporstIntoExternalFiles.vbs"
		!insertmacro OC_STATUS "�������� �������� | ExtractAllReporstIntoExternalFiles.vbs ..."
		Delete "$INSTDIR\config\scripts\MD_Tasks\ExtractAllReporstIntoExternalFiles.vbs"
	SectionEnd ;; un.ExtractAllReporstIntoExternalFiles.vbs
	Section "un.AfterUnInstallMD_Tasks"
		RMDir "$INSTDIR\config\scripts\MD_Tasks"
	SectionEnd
SubSectionEnd ;; un.MD_Tasks
SubSection un."Intellisence"
	Section "un.als2xml.js"
		!insertmacro OC_STATUS "�������� �������� | als2xml.js ..."
		Delete "$INSTDIR\config\scripts\Intellisence\als2xml.js"
	SectionEnd ;; un.als2xml.js
	!include "scripts\un.Intellisence.nsh"
	Section "un.Intellisence_DocRef.vbs"
		!insertmacro OC_STATUS "�������� �������� | Intellisence_DocRef.vbs ..."
		Delete "$INSTDIR\config\scripts\Intellisence\Intellisence_DocRef.vbs"
	SectionEnd ;; un.Intellisence_DocRef.vbs
	!include "scripts\un.intsOLEGenerator.nsh"
	Section "un.VimComplete.js"
		!insertmacro OC_STATUS "�������� �������� | VimComplete.js ..."
		Delete "$INSTDIR\config\scripts\Intellisence\VimComplete.js"
	SectionEnd ;; un.VimComplete.js
	Section "un.AfterUnInstallIntellisence"
		RMDir "$INSTDIR\config\scripts\Intellisence"
	SectionEnd
SubSectionEnd ;; Intellisence
