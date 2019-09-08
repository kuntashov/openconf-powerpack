var msg = function (str) { WScript.Echo(str) }

var wsh = new ActiveXObject("WScript.Shell");
var RegRead 	= function (strName) 					{ return wsh.RegRead(strName) }
var RegWrite	= function (strName, strValue, Type)	{ wsh.RegWrite(strName, strValue, Type?Type:"REG_SZ") }
var WriteDWORD	= function (strName, strValue)			{ wsh.RegWrite(strName, strValue, "REG_DWORD") }
var RegDelete	= function (strName)					{ wsh.RegDelete(strName) }

var assign = function (a, b) { if (a!=b) throw(new Error(0, "Assignment failed: " + a + " != " + b))  }

msg('---------------------------------------------');
msg('  Testing OpenConf.RegistryKeyMap component  ');
msg('---------------------------------------------');

var Counter = 0;
var Failed = 0;

var mask = null;
if (WScript.Arguments.Count() > 0) {
	mask = new RegExp(WScript.Arguments(0));
}

var fso = new ActiveXObject("Scripting.FileSystemObject");
var dir = fso.GetFolder(".\\RegKeyMap");
var tests = new Enumerator(dir.Files);

for(; !tests.atEnd(); tests.moveNext()) {
	var test = tests.item();
	if (mask instanceof RegExp) {
		if (!mask.test(test.Name)) {
			continue;
		}
	}
	Counter++
	msg("Running test " + test.Name);
	try {
		var RKM = new ActiveXObject("OpenConf.RegistryKeyMap");
		var closed = false;
		var close = function () { if (!closed) { RKM.Close(); closed=true } }

		eval(test.OpenAsTextStream().ReadAll());

		close();
		msg("OK.");
	}
	catch (e) {
		Failed++;
		msg('FAILED. Reason: '+ e.description);
	}
}

msg('---------------------------------------------');
msg('Tests passed: ' + (Counter - Failed) + "/" + Counter);
msg('Tests failed: ' + Failed + "/" + Counter);
