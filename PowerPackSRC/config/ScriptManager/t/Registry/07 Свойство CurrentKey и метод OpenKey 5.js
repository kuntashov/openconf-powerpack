// Несуществующий ключ
try { // to open
	R.CurrentKey = "HKEY_CLASSES_ROOT\\TheKeyThatDoesNotExists"
} 
catch (e) {
	assign(e.description, "Ключ не существует: HKEY_CLASSES_ROOT\\TheKeyThatDoesNotExists")
}