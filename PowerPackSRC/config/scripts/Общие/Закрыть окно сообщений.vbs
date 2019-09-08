'
'         Мой e-mail: artbear@bashnet.ru
'         Мой ICQ: 265666057
'

Sub TogglePanel(PanelName)
	Windows.PanelVisible(PanelName)=Not Windows.PanelVisible(PanelName)
End Sub

' Закрыть окно сообщений, чтобы не читать ненужные сообщения
Sub Configurator_AllPluginsInit()
	TogglePanel "Окно сообщений"
	Scripts.Unload SelfScript.Name
End Sub

TogglePanel "Окно сообщений"