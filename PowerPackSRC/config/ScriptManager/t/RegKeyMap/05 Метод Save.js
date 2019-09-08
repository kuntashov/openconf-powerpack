RegWrite("HKCU\\TestHotKeys\\")

RKM.RootKey = "HKCU\\TestHotKeys"
RKM.Add(0xbe0010, true, "Modul1", "Macros1")
RKM.Add(0x720000, true, "Modul2", "Macros2")

RKM.Save()
close()

// сохранение порядка следования элементов не гарантируется,
// поэтому приходится прибегать к ухищрениям, дабы однажды вдруг
// тест не сломался только из-за того, что элементы вдруг будут
// записаны не в том порядке, что мы ожидаем

var codes = { 0xbe0010 : "0", 0x720000 : "1" }
if (RegRead("HKCU\\TestHotKeys\\0\\code")==0x720000) {
	 codes = { 0xbe0010 : "1", 0x720000 : "0" }
}

assign(RegRead("HKCU\\TestHotKeys\\"+codes[0xbe0010]+"\\code"),		0xbe0010)
assign(RegRead("HKCU\\TestHotKeys\\"+codes[0xbe0010]+"\\IsScript"),	1)
assign(RegRead("HKCU\\TestHotKeys\\"+codes[0xbe0010]+"\\modul"),	"Modul1")
assign(RegRead("HKCU\\TestHotKeys\\"+codes[0xbe0010]+"\\macros"),	"Macros1")

assign(RegRead("HKCU\\TestHotKeys\\"+codes[0x720000]+"\\code"),		0x720000)
assign(RegRead("HKCU\\TestHotKeys\\"+codes[0x720000]+"\\IsScript"),	1)
assign(RegRead("HKCU\\TestHotKeys\\"+codes[0x720000]+"\\modul"),	"Modul2")
assign(RegRead("HKCU\\TestHotKeys\\"+codes[0x720000]+"\\macros"),	"Macros2")

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