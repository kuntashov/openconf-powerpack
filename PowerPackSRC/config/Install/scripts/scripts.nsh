;===========================================================================
;; �������� ��������� ��������

SubSection /e "�������"

	!include "scripts\scripts.auto.nsh"

SubSectionEnd ;; �������

;===========================================================================
;; �������� ������������� ��������

!ifdef OC_Uninstall
SubSection /e un."�������"

	!include "scripts\un.scripts.auto.nsh"

SubSectionEnd ;; un.�������
!endif ;; OC_Uninstall