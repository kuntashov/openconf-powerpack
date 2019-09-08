/* прячем ненавистную мне панель "Конструкторы", а за одно 
и редко используемые "Конструкторы" и "Администрирование" */

function TogglePanelOff(PanelName) {
	if (Windows.PanelVisible(PanelName)) {
		Windows.PanelVisible(PanelName) = false;
	}
}

function Configurator::OnActivateWindow() {
	TogglePanelOff("Конструкторы");
	TogglePanelOff("Конфигурация");
	TogglePanelOff("Администрирование");
}