/* ������ ����������� ��� ������ "������������", � �� ���� 
� ����� ������������ "������������" � "�����������������" */

function TogglePanelOff(PanelName) {
	if (Windows.PanelVisible(PanelName)) {
		Windows.PanelVisible(PanelName) = false;
	}
}

function Configurator::OnActivateWindow() {
	TogglePanelOff("������������");
	TogglePanelOff("������������");
	TogglePanelOff("�����������������");
}