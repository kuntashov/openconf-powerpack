// �������������� ����
try { // to open
	R.CurrentKey = "HKEY_CLASSES_ROOT\\TheKeyThatDoesNotExists"
} 
catch (e) {
	assign(e.description, "���� �� ����������: HKEY_CLASSES_ROOT\\TheKeyThatDoesNotExists")
}