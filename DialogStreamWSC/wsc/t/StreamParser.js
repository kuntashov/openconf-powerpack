function  print(str) {
    WScript.Echo(str);
}
               
var sp = new ActiveXObject("OpenConf.StreamParser");

sp.Source = '{"Main stream", """Test2""", {"Child stream", "0", "1", "2"}, """"""}';

if (!sp.Parse()) {
    print(sp.LastError.description);
} 
else {
}

print(sp.Stream.Data.Item(0).Data);
print(sp.Stream.Data.Item(1).Data);

var c = sp.Stream.Data.Item(2).Data;


var s = ' ';
for (var i=0; i<c.Size(); i++) {
    s += c.Item(i).Data + ' ';
}

print('Child stream data: ' + s);

print(sp.Stream.Data.Item(3).Data)
print("------------------------------------------");
print("Full source:");
print(sp.Stream.Stringify());
