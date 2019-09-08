RegWrite("HKCU\\TestKey\\Value", "Hello, World!")

R.CurrentKey = "HKCU\\TestKey"
assign(R.ReadValue(0), "Hello, World!")
assign(R.ReadValue("Value"), "Hello, World!")

RegDelete("HKCU\\TestKey\\Value")
RegDelete("HKCU\\TestKey\\")