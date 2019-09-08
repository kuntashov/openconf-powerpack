RegWrite("HKCU\\TestHotKeys\\0\\code", 		0x740000,	"REG_DWORD") // F5
RegWrite("HKCU\\TestHotKeys\\0\\IsScript",	1,			"REG_DWORD")
RegWrite("HKCU\\TestHotKeys\\0\\modul",		"HotKeysHandler")
RegWrite("HKCU\\TestHotKeys\\0\\macros",	"F5")

RegWrite("HKCU\\TestHotKeys\\1\\code", 		0x340018,	"REG_DWORD") // Ctrl+Alt+4
RegWrite("HKCU\\TestHotKeys\\1\\IsScript",	1,			"REG_DWORD")
RegWrite("HKCU\\TestHotKeys\\1\\modul",		"HotKeysHandler")
RegWrite("HKCU\\TestHotKeys\\1\\macros",	"CtrlAlt4")

var CrLf = "\r\n"
var map = '//keymap file, version=1.0' + CrLf
		+ 'hk.F5 = new Array(9)' + CrLf
		+ 'hk.F5[0] = {' + CrLf
		+ '	"script" : "jsSMTest",' + CrLf 
		+ '	"macros" : "NoWindow"' + CrLf
		+ '}' + CrLf
		+ 'hk.F5[4] = {' + CrLf
		+ '	"script" : "jsSMTest",' + CrLf 
		+ '	"macros" : "OpenConfMDTab"' + CrLf
		+'}' + CrLf
		+ 'hk.CtrlAlt4 = new Array(9)' + CrLf
		+ 'hk.CtrlAlt4[2] = {' + CrLf
		+ '	"script" : "jsSMTest",' + CrLf 
		+ '	"macros" : "Interfaces"' + CrLf
		+'}' + CrLf;


var fPath = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName())
var f = fso.CreateTextFile(fPath, true)
f.Write(map);
f.Close();


KMC.rPath = "HKCU\\TestHotKeys"
KMC.fPath = fPath;
  
assign(KMC.Load(), true)

//fso.DeleteFile(fPath, true);

RegDelete("HKCU\\TestHotKeys\\1\\code")
RegDelete("HKCU\\TestHotKeys\\1\\IsScript")
RegDelete("HKCU\\TestHotKeys\\1\\modul")
RegDelete("HKCU\\TestHotKeys\\1\\macros")
RegDelete("HKCU\\TestHotKeys\\1\\")
RegDelete("HKCU\\TestHotKeys\\0\\code")
RegDelete("HKCU\\TestHotKeys\\0\\IsScript")
RegDelete("HKCU\\TestHotKeys\\0\\modul")
RegDelete("HKCU\\TestHotKeys\\0\\macros")
RegDelete("HKCU\\TestHotKeys\\0\\")
RegDelete("HKCU\\TestHotKeys\\")
