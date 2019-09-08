var TemplateJS = 'function %ShortCut%() { HotKeyEventProxy.fireHotKeyEvent("%ShortCut%") }';
var TemplateReg = '[HKEY_CURRENT_USER\\Software\\1C\\1Cv7\\7.7\\OpenConf\\HotKeys\\%Counter%]' + "\r\n"
				+ '"IsScript"=dword:00000001' + "\r\n"
				+ '"modul"="HotKeysHandler"' + "\r\n"
				+ '"macros"="%ShortCut%"' + "\r\n"
				+ '"code"=dword:%HexCharCode%' + "\r\n";

var ExclusionList = {}
var KM = new ActiveXObject("ScriptManager.KeyMap")
var VK = KM.VK

function usage()
{
	return "cscript //nologo hkgen.js file";
}

function Writer(jsFileName, regFileName)
{
	var fso = new ActiveXObject("Scripting.FileSystemObject");

	this.Counter = -1;

	this.jsFile = jsFileName;
	this.regFile = regFileName;

	this.js = this.jsFile?fso.OpenTextFile(jsFileName, 2, true):false;
	this.reg = this.regFile?fso.OpenTextFile(regFileName, 2, true):null; 

	if (this.reg) {
		this.reg.Write("REGEDIT4\r\n\r\n"
			+ "[HKEY_CURRENT_USER\\Software\\1C\\1Cv7\\7.7\\OpenConf\\HotKeys]\r\n\r\n");
	}
	
	this.write = function(Char, Code, Ctrl, Alt, Shift)
	{
		var ShortCut = "";
		var _Code = Code << 16;
		if (Ctrl)	{ ShortCut += "Ctrl";	_Code |= 0x8;	}
		if (Alt)	{ ShortCut += "Alt"; 	_Code |= 0x10;	}
		if (Shift)	{ ShortCut += "Shift";	_Code |= 0x4;	}
		ShortCut += Char;
		if (ExclusionList[ShortCut]) return;
		var HexCharCode = _Code.toString(16);
		this.Counter++;
		if (this.js) {
			this.js.WriteLine(
				TemplateJS.replace(/%ShortCut%/g, ShortCut));
		}
		if (this.reg) {
			this.reg.WriteLine(
				TemplateReg.replace(/%ShortCut%/g, ShortCut)
					.replace(/%Counter%/g, this.Counter).replace(/%HexCharCode%/g, HexCharCode));
		}
	}
	this.rem = function (str)
	{
		if (str) {
			if (this.js)  this.js.WriteLine("// " + str);
			if (this.reg) this.reg.WriteLine(";; " + str);
		}
	}
	this.shutdown = function ()
	{
		if (this.js)  this.js.Close();
		if (this.reg) this.reg.Close();
	}
}

function genShortCuts(wr, Key, Code, Excl)
{
	// ------------------------------ Ctrl   Alt    Shift ----------		
	if (!Excl[1]) wr.write(Key ,Code ,false ,false ,false); //
	if (!Excl[2]) wr.write(Key ,Code ,true  ,false ,false); // Ctrl
	if (!Excl[3]) wr.write(Key ,Code ,false ,true  ,false); // Alt
	if (!Excl[4]) wr.write(Key ,Code ,false ,false ,true ); // Shift
	if (!Excl[5]) wr.write(Key ,Code ,true  ,true  ,false); // Crtl Alt
	if (!Excl[6]) wr.write(Key ,Code ,true  ,false ,true ); // Ctrl Shift
	if (!Excl[7]) wr.write(Key ,Code ,false ,true  ,true ); // Alt Shift
	if (!Excl[8]) wr.write(Key ,Code ,true  ,true  ,true ); // Ctrl Alt Shift

}

function ExclList()
{
	var ret = Array(9);
	if (arguments.length > 8) return arguments;	
	for(var i=0; i<arguments.length; i++) {
		ret[i+1] = arguments[i];
	}
	return ret;
}

function main()
{
	if (!WScript.Arguments.Count()) {
		WScript.Echo(usage());
		WScript.Quit(1)
	}
	var file = WScript.Arguments(0)
	WScript.Echo('Generating file ' + file + ' ...')
	var wr = new Writer(file)
	for (var key in VK) {
		genShortCuts(
			wr, KM.toShortCut(VK[key]<<16), 
			VK[key], ExclList(!((0x69<key)&&(key<0x7c)))
		)
	}
	for (var i=0; i<10; i++) {
		var key = i.toString()
		genShortCuts(
			wr, key, key.charCodeAt(0), 
			ExclList(true,false,false,true)
		)	
	}
	for (var code='A'.charCodeAt(0); code<='Z'.charCodeAt(0); code++) {
		genShortCuts(
			wr, String.fromCharCode(code), 
			code, ExclList(true,false,false,true)
		)
	}
	wr.shutdown();
	WScript.Echo("Done!");
}

main();
