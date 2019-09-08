function  print(str) {
    WScript.Echo(str);
}

function assign(result) {
    return result ? 'Ok' : 'Failed';
}


var c = new ActiveXObject("OpenConf.Collection");

c.Add("test");
c.Add("test2");

print(assign(c.Item(0)==="test"));
print(assign(c.Item(1)==="test2"));
print(assign(c.Size() === 2));

// ToDo: More Tests !!!!
