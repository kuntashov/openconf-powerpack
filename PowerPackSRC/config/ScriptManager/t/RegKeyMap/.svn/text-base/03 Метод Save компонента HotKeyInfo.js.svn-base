RegWrite("HKCU\\TestHotKeys\\0\\")

RKM.Add(0xbe0010, 1, "", "") // просто чтобы создать объект HotKeyInfo
assign(RKM.Count(), 1)

var hki = RKM.HotKeyInfo(0xbe0010)

hki.Code = 0xbe0010
hki.IsScript = true
hki.Modul = "TestModul"
hki.Macros = "TestMacros"

var reg = new ActiveXObject("kuntashov.Registry")
reg.OpenKey("HKCU\\TestHotKeys\\0")

hki.Save(reg)
assign(RegRead("HKCU\\TestHotKeys\\0\\code"),		0xbe0010)
assign(RegRead("HKCU\\TestHotKeys\\0\\IsScript"),	1)
assign(RegRead("HKCU\\TestHotKeys\\0\\modul"),		"TestModul")
assign(RegRead("HKCU\\TestHotKeys\\0\\macros"),		"TestMacros")

reg.Close()
reg = null

RegDelete("HKCU\\TestHotKeys\\0\\code")
RegDelete("HKCU\\TestHotKeys\\0\\IsScript")
RegDelete("HKCU\\TestHotKeys\\0\\modul")
RegDelete("HKCU\\TestHotKeys\\0\\macros")
RegDelete("HKCU\\TestHotKeys\\0\\")
RegDelete("HKCU\\TestHotKeys\\")