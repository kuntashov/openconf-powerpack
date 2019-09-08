RegWrite("HKCU\\TestHotKeys\\0\\code", 		0x740000,	"REG_DWORD") // F5
RegWrite("HKCU\\TestHotKeys\\0\\IsScript",	1,			"REG_DWORD")
RegWrite("HKCU\\TestHotKeys\\0\\modul",		"HotKeysHandler")
RegWrite("HKCU\\TestHotKeys\\0\\macros",	"F5")

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
		+'}' + CrLf;

var CtrlAlt4 = new Array(9)
CtrlAlt4[2] = { "script" : "jsSMTest", "macros" : "Interfaces" }

var fPath = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName())
var f = fso.CreateTextFile(fPath, true)
f.Write(map);
f.Close();

KMC.rPath = "HKCU\\TestHotKeys"
KMC.fPath = fPath;

assign(KMC.Load(), true)

KMC.HotKeyInfo("CtrlAlt4") = CtrlAlt4

var hki = KMC.ocKM.HotKeyInfo(0x340018)
assign((hki != null), true)

assign(hki.Code, 0x340018)
assign(hki.IsScript, 1)
assign(hki.Modul, "HotKeysHandler")
assign(hki.Macros, "CtrlAlt4")
 
//fso.DeleteFile(fPath)

RegDelete("HKCU\\TestHotKeys\\0\\code")
RegDelete("HKCU\\TestHotKeys\\0\\IsScript")
RegDelete("HKCU\\TestHotKeys\\0\\modul")
RegDelete("HKCU\\TestHotKeys\\0\\macros")
RegDelete("HKCU\\TestHotKeys\\0\\")
RegDelete("HKCU\\TestHotKeys\\")
