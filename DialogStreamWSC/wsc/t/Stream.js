var sp = new ActiveXObject("OpenConf.StreamParser");

var t = ['{}', '{{}', '{"Test"}', '{"1", {"2",{"3"}},{"222","22222"}}'];

for (var i in t) {
    if (sp.Parse(t[i])) {
        WScript.Echo('Ok');
        WScript.Echo(sp.Stream.item(0).toString());
    }
    else {
        WScript.Echo('Failed!' + sp.LastError);
    }  
}