RegWrite("HKCU\\TestKey\\Value", "Hello, World!")

R.CurrentKey = "HKCU\\TestKey"

// ��������� ��������� ��������

try { 
	R.ReadValue(2) 
}
catch (e) { 
	assign(e.description, "������ �� ��������� ��������: 2"); 
}

try { 
	R.ReadValue("Value2")
}
catch (e) { 
	assign(e.description, "������ �� ��������� ��������: Value2"); 
}

RegDelete("HKCU\\TestKey\\Value")
RegDelete("HKCU\\TestKey\\")