var Handlers = {
	onClose			: UnloadSelf,
	onCancelClick	: CloseHtmlWindow,
	onSaveClick		: SaveHotKeys,
	onSelMacroClick	: SelectMacros,
	onDelMacroClick	: DeleteMacros,
	onCheckBoxClick : FillForm,
	onSelectKey		: FillForm,
	onLoad			: InitForm
}

var mList = null;

// XXX оформить отдельным скриптлетом?
function MacrosList(_)
{
	var list = new ActiveXObject("Scripting.Dictionary"); 
	
	this.getRawList = function ()
	{
		return list;
	}
	this.Count = function ()
	{
		return list.Count;
	}
	this.Load = function ()
	{
		var me = new ActiveXObject("Macrosenum.Enumerator");
		for (var i=0; i<Scripts.Count; i++) {
			var sName = Scripts.Name(i);
			if ((sName!=SelfScript.Name)&&(sName!="HotKeysHandler")) {
				var macroses = (new VBArray(me.EnumMacros(Scripts(i)))).toArray();
				if (macroses.length>0) {
					list.Add(sName, macroses);
				}
			}
		}
		me = null;
	}
	this.Reload = function ()
	{
		list.RemoveAll();
		this.Load();
	}
	this.Stringify = function (selScr)
	{
		var str="";
		var scripts = (new VBArray(list.Keys())).toArray();
		for(var i=0; i<scripts.length; i++) {
			var macroses = list.Item(scripts[i]);
			str += scripts[i] 
				+ ((selScr&&(selScr==scripts[i]))?"|e":"") + "\r\n"					
			for(var j=0; j<macroses.length; j++) {
				str += "\t" + macroses[j] + "\r\n";
			}
		}		
		return str;
	}
	this.Select = function (selScr)
	{
		var str = this.Stringify(selScr);
		var ret = SvcSvc.SelectInTree(str, "Выберите макрос", false, true);
		if (ret != "") {
			ret = ret.replace(/\\/,'::');
		}
		return ret;
	}
	this.Shutdown = function ()
	{
		list = null;
	}
}

function CloseHtmlWindow()
{
	Windows.ActiveWnd.Close();
}

function UnloadSelf(_)
{
	if (mList) mList.Shutdown();
	Scripts.Unload(SelfScript.Name);
}

function SelectMacros(hkForm, event)
{
	var txtName = event.srcElement.name.replace(/^btn/, 'txt');

	// позаботимся о том, чтобы при перевыборе макроса
	// в диалоге выбора была развернута ветка, соответствующая 
	// скрипту ранее выбранного макроса
	var curSel = hkForm[txtName].value;
	curSel = (curSel == "<Не задан>") ? undefined : (curSel.split(/::/))[0];

	if (!mList.Count()) {		
		mList.Load(); 
	}

	var ret = mList.Select(curSel);
	if (ret=="") return;

	hkForm[txtName].value = ret;
	SyncWithKeyMap(hkForm);
}

function DeleteMacros(hkForm, event)
{
	var txtName = event.srcElement.name.replace(/^btnX/, 'txtMacro');
	hkForm[txtName].value = '<Не задан>';
	SyncWithKeyMap(hkForm);
}

function SaveHotKeys(hkForm)
{
	if (!KMC.Save()) {
		Message("Ошибка сохранения кеймапа!", mRedError);
		if (KMC.LastError) {
			Message(KMC.LastError.description);
		}
		return;
	}
	with (new ActiveXObject("WScript.Shell")) {
		Popup("Изменения успешно сохранены и вступят \r\n"
			+ "в силу после перезапуска Конфигуратора!", 
			0,"Hot Keys Editor", 64)
	}
	CloseHtmlWindow();
}

function InitForm(hkForm, divShortCut)
{
	var select = '<select name="selKey" onchange="onSelectKey()">' + "\r\n";
	var VK = KMC.smKM.VK;
	for (var key in VK) {
		select	+= '<option value="' + VK[key] + '" />' + key + '</option>' + "\r\n";
	}
	for (var i=0; i<10; i++) {
		select	+= '<option value="' + i.toString().charCodeAt(0) + '" />' 
				+ i.toString() + '</option>' + "\r\n";
	}
	for (var code='A'.charCodeAt(0); code<='Z'.charCodeAt(0); code++) {
		select	+= '<option value="' + code + '">' 
				+ String.fromCharCode(code) + '</option>' + "\r\n";
	}
	select += '</select>'
	divShortCut.innerHTML += select;
	FillForm(hkForm);
}

function getShortCut(hkForm)
{
	var Ctrl	= hkForm.chbCtrl.checked;
	var Alt		= hkForm.chbAlt.checked;
	var Shift	= hkForm.chbShift.checked;

	var selIx	= hkForm.selKey.selectedIndex;
	var Key		= parseInt(hkForm.selKey.options[selIx].value);

	// без модификаторов Ctrl и Alt можно использовать
	// только функциональные клавиши (F1-F12)
	if (!((0x69<Key)&&(Key<0x7c))) {
		if (!(Ctrl || Alt)) {
			hkForm.chbAlt.checked = Alt = true;
		}
	}

	return (Ctrl?'Ctrl':'')	+ (Alt?'Alt':'')
		+ (Shift?'Shift':'') + toShortCut(Key<<16);
}

function FillForm(hkForm)
{
	var ShortCut = getShortCut(hkForm);
	var info = KMC.HotKeyInfo(ShortCut);
	
	if (!info) {
		for (var i=0; i<9; i++) {
			hkForm['txtMacro' + i].value = '<Не задан>';
		}
		return;
	}

	for (var i=0; i<9; i++) {
		var value = '<Не задан>';
		if (info[i]) {
			value = info[i].script + '::' + info[i].macros;
		}
		hkForm['txtMacro' + i].value = value;
	}
}

function SyncWithKeyMap(hkForm)
{
	var info = new Array(9);
	for (var i=0; i<9; i++) {
		var txtMacroX = hkForm['txtMacro' + i].value;
		if (txtMacroX != '<Не задан>') {
			var parts = txtMacroX.split(/::/);
			info[i] = { 'script' : parts[0], 'macros' : parts[1] }
		}
	}
	var ShortCut = getShortCut(hkForm);
	KMC.HotKeyInfo(ShortCut) = info;
}

function fromCode(code) { return KMC.smKM.fromCode(code) }
function toCode(char)	{ return KMC.smKM.toCode(char) }

function fromShortCut(shortCut) { return KMC.smKM.fromShortCut(shortCut) }
function toShortCut(scan)		{ return KMC.smKM.toShortCut(scan) }

function onDocumentComplete(d, u)
{
	try {				
		HtmlWindow.Document.Script.SetHandlers(Handlers);
		if (!KMC.Load()) {
			Message("Ошибка загрузки кеймапа", mRedErr);
			if (KMC.LastError) {
				Message(KMC.LastError.description, mRedError);
			}
			CloseHtmlWindow();
			return;
		}
		Windows.ActiveWnd.Maximized = true;
	}
	catch (e) {
		Message("Ошибка инициализации окна редактора хоткеев", mRedErr);
		Message(e.description, mRedErr);
	}
}

function OpenHotKeyEditor(_)
{
	var wnd = OpenOleForm("Shell.Explorer", "Hot Key Editor");
	SelfScript.AddNamedItem("HtmlWindow", wnd, false);
    eval('function HtmlWindow::DocumentComplete(d,u){ return onDocumentComplete(d, u) }');
	HtmlWindow.Navigate2(BinDir+"config\\system\\ScriptManager\\html\\HotKeyEditor.htm");	
}

function CreateObjectOrDie(progid, id) 
{
	try {
		var obj = new ActiveXObject(progid);
		if (id) {
			SelfScript.AddNamedItem(id, obj, false);		
		}
		return obj;
	}
	catch (e) {
		Message("Не могу создать объект " + progid, mRedErr);
		Message(e.description, mRedErr);
		Message("Скрипт " + SelfScript.Name + " не загружен", mInformation);
		Scripts.UnLoad(SelfScript.Name); 		
	}		
}

function Init(_)
{
	//CreateObjectOrDie("OpenConf.CommonServices", "CommonScripts");
	CreateObjectOrDie("SvcSvc.Service", "SvcSvc");
	CreateObjectOrDie("ScriptManager.KeyMapsContainer", "KMC");
	KMC.fPath = BinDir + "\\config\\system\\ScriptManager\\test.keymap.js";
	mList = new MacrosList();
	OpenHotKeyEditor();
}

Init();
