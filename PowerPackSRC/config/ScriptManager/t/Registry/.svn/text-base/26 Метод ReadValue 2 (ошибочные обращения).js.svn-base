RegWrite("HKCU\\TestKey\\Value", "Hello, World!")

R.CurrentKey = "HKCU\\TestKey"

// Тестируем ошибочные ситуации

try { 
	R.ReadValue(2) 
}
catch (e) { 
	assign(e.description, "Индекс за пределами значения: 2"); 
}

try { 
	R.ReadValue("Value2")
}
catch (e) { 
	assign(e.description, "Индекс за пределами значения: Value2"); 
}

RegDelete("HKCU\\TestKey\\Value")
RegDelete("HKCU\\TestKey\\")