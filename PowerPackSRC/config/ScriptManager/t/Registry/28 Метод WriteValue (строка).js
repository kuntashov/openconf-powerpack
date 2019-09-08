RegWrite("HKCU\\TestKey\\", "1") 

R.CurrentKey = "HKCU\\TestKey"
assign(R.ValuesCount(), 1) // значение по умолчанию 

R.WriteValue("Hello", "Hello, World!")
assign(R.ValuesCount(), 2);
assign(R.ReadValue("Hello"), "Hello, World!")

RegDelete("HKCU\\TestKey\\Hello")
RegDelete("HKCU\\TestKey\\")