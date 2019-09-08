var short_names	= ['HKCR', 'HKCU', 'HKLM']
var long_names	= [ 'HKEY_CLASSES_ROOT', 'HKEY_CURRENT_USER', 'HKEY_LOCAL_MACHINE']

// "длинные" имена
for (var i=0; i<3; i++) {
	R.RootKey = long_names[i]
	assign(R.RootKey, long_names[i])
}

// "короткие" имена
for (var i=0; i<3; i++) {
	R.RootKey = short_names[i]
	assign(R.RootKey, long_names[i])
}

// тестируем обработку ошибок
try {
	R.RootKey = "erroneous_root_key_name"
	throw(new Error(0, "Ошибка не обработана"))
}
catch (e) {
	assign(e.description, "Неизвестное имя раздела: " + "erroneous_root_key_name")
}