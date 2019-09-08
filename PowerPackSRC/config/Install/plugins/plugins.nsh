;===========================================================================
;; �������� ���������/������������� ��������

;===========================================================================
;; ��������� ���������

SubSection /e "�������"

	Section -"PluginsBeforeInstall"
		!insertmacro OC_STATUS "��������� ��������..."
	SectionEnd
	
	!include "Telepat20.nsh"
	!include "HistoryPlugin.nsh"
	!include "SubSystems.nsh"
	!include "ExtFormsTree.nsh"
	!include "Oc2Vba.nsh"
	!include "Vis1Cpp.nsh"
	!include "ClassesWizard.nsh"
	!include "FDSubst.nsh"
	!include "Inspector2.nsh"
	!include "KbdProcs.nsh"

	Section -"PluginsAfterInstall"
		!insertmacro OC_STATUS "��������� �������� | ������"
	SectionEnd

SubSectionEnd

;===========================================================================
;; ��������� �������������

!ifdef OC_Uninstall
SubSection un."�������"

	Section un."PluginsBeforeUnInstall"
		!insertmacro OC_STATUS "�������� ��������..."
	SectionEnd

	!include "un.Telepat20.nsh"
	!include "un.HistoryPlugin.nsh"
	!include "un.SubSystems.nsh"
	!include "un.ExtFormsTree.nsh"
	!include "un.Oc2Vba.nsh"
	!include "un.Vis1Cpp.nsh"
	!include "un.ClassesWizard.nsh"
	!include "un.FDSubst.nsh"
	!include "un.Inspector2.nsh"
	!include "un.KbdProcs.nsh"

	Section un."AfterUninstallPlugins"
		;; ���� ���������� system ������, ������ �� �� ���������� 
		;; �������� ������������� ���� ��������
		RMDir "$INSTDIR\config\system"
		!insertmacro OC_STATUS "�������� �������� | ������"
	SectionEnd

SubSectionEnd
!endif ;; OC_Uninstall
