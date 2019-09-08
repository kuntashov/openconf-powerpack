/*===========================================================================
Скрипт:  DocPath.js
Версия:  $Revision: 1.1 $
Автор:   Александр Кунташов
E-mail:  kuntashov at [yandex.ru|gmail.com]
ICQ UIN: 338758861
Описание: 
    Макросы для показа/копирования в буфер обмена полного пути 
    открытого документа (внешнего отчета, мокселя или текстового файла).
    
	Макросы:
	
    	ShowDocPath	- выводит путь текущего документа в окно сообщений
    	CopyDocPath	- копирует путь текущего документа в буфер обмена 
    	OpenParentDir - открывает каталог, в котором находится текущий 
						документ в эксплорере
    	CopyParentDir - копирует каталог, в котором находится текущий
						документ в буфер обмена
===========================================================================*/

var CoreObj = {
		
	version : "$Revision: 1.1 $".replace(/\s|\$/ig, "").replace(/^Revision:/, ""),

	ocs : null,

	init : function () 
	{
		try {
			this.ocs = new ActiveXObject("OpenConf.CommonServices");
			this.ocs.SetConfig(Configurator);
		}
		catch (e) {
			Message("Не могу создать объект OpenConf.CommonServices", mRedErr);
			Message(e.description, mRedErr);
			Message("Скрипт " + SelfScript.Name + " не загружен", mInformation);
			Scripts.UnLoad(SelfScript.Name); 		
		}	
	},
	
	show : function ()
	{
		var a = null;
		if (a = Windows.ActiveWnd) {
			if (a = a.Document) {
				Message(a.Path, mInformation);
			}
		}
	},

	copy : function ()
	{
		var a = null;
		if (a = Windows.ActiveWnd) {
			if (a = a.Document) {
				this.ocs.CopyToClipboard(a.Path);
			}
		}
	},

	openDir : function ()
	{
		var a = null;
		if (a = Windows.ActiveWnd) {
			if (a = a.Document) {
				var p = this.ocs.FSO.GetParentFolderName(a.Path);				
				this.ocs.WSH.Run("%SYSTEMROOT%\\explorer.exe " + p);
			}
		}
	},

	copyDir : function ()
	{
		var a = null;
		if (a = Windows.ActiveWnd) {
			if (a = a.Document) {
				var p = this.ocs.FSO.GetParentFolderName(a.Path);
				this.ocs.CopyToClipboard(p);
			}
		}
	}

}

function ShowDocPath() { CoreObj.show() }
function CopyDocPath() { CoreObj.copy() }

function OpenParentDir() { CoreObj.openDir() }
function CopyParentDir() { CoreObj.copyDir() }

CoreObj.init()
