'
'         ��� e-mail: artbear@bashnet.ru
'         ��� ICQ: 265666057
'

Sub TogglePanel(PanelName)
	Windows.PanelVisible(PanelName)=Not Windows.PanelVisible(PanelName)
End Sub

' ������� ���� ���������, ����� �� ������ �������� ���������
Sub Configurator_AllPluginsInit()
	TogglePanel "���� ���������"
	Scripts.Unload SelfScript.Name
End Sub

TogglePanel "���� ���������"