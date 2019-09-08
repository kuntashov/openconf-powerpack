RegWrite("HKCU\\TestHotKeys\\0\\code", 		0xbe0010,	"REG_DWORD")
RegWrite("HKCU\\TestHotKeys\\0\\IsScript",	1,			"REG_DWORD")
RegWrite("HKCU\\TestHotKeys\\0\\modul",		"TestModul")
RegWrite("HKCU\\TestHotKeys\\0\\macros",	"TestMacros")

RKM.Add(0xbe0010, 1, "", "") // просто чтобы создать объект HotKeyInfo
assign(RKM.Count(), 1)

var hki = RKM.HotKeyInfo(0xbe0010)

var reg = new ActiveXObject("kuntashov.Registry")
reg.CurrentKey = "HKCU\\TestHotKeys\\0"

hki.Load(reg)
assign(hki.Code, 0xbe0010)
assign(hki.IsScript, true)
assign(hki.Modul, "TestModul")
assign(hki.Macros, "TestMacros")

reg.Close()
reg = null

RegDelete("HKCU\\TestHotKeys\\0\\code")
RegDelete("HKCU\\TestHotKeys\\0\\IsScript")
RegDelete("HKCU\\TestHotKeys\\0\\modul")
RegDelete("HKCU\\TestHotKeys\\0\\macros")
RegDelete("HKCU\\TestHotKeys\\0\\")
RegDelete("HKCU\\TestHotKeys\\")