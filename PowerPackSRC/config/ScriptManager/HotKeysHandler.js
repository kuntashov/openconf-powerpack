//$NAME HotKeysHandler

var HotKeyEventProxy = 
{
	ScriptManager : null,

	init : function ()
	{
		//try {
			this.ScriptManager = new ActiveXObject("OpenConf.ScriptManager");
			this.ScriptManager.OnInit(this, Configurator);
		//} 
		/*catch (e) {
			Message("Ошибка инициализации менеджера скриптов!", mRedErr);
			Message("Описание ошибки: " + e.description, mRedErr);
			Message("Скрипты не загружены!");
			Scripts.UnLoad(SelfScript.Name);
		}*/
	},

	fireHotKeyEvent : function (hk)
	{
		this.ScriptManager.OnHotKey(hk);
	},
	
	shutdown : function ()
	{
		if (this.ScriptManager) {
			this.ScriptManager.Shutdown();
			this.ScriptManager = null;
		}
	},
	openHotKeysEditor : function ()
	{
		this.ScriptManager.OpenHotKeysEditor();
	}
}

/* инициализация */
HotKeyEventProxy.init();

function OpenHotKeysEditor()
{
	HotKeyEventProxy.openHotKeysEditor();
}

function Configurator::OnQuit()
{
	HotKeyEventProxy.shutdown();
}

//======================================================================================

eval((new ActiveXObject("Scripting.FileSystemObject")).OpenTextFile(BinDir+"config\\system\\ScriptManager\\hotkeys.js").ReadAll());
