function print(str)
{
    WScript.Echo(str);
}

var a = new Array(
//=========== Идентификатор  Номер с-ва     Комментарий
    new Array("Caption",         1,          "Заголовок элемента управления"),
    new Array("Type",            2,          "Тип элемента управления"),
    new Array("Flags1",          3,          "Набор дополнительных флагов элемента управления"),
    new Array("X",               4,          "Координата X"),
    new Array("Y",               5,          "Координата Y"),
    new Array("Width",           6,          "Ширина элемента управления"),
    new Array("Height",          7,          "Высота элемента управления"),
    new Array("Associated",      8,          "Признак привязки к реквизиту объекта конфишурации (1-привязан, 2 - нет)"),
    new Array("TabOrder",       10,          "Порядок обхода"),
    new Array("Action",         12,          "Формула"),
    new Array("Name",           13,          "Идентификатор элемента управления"),
    new Array("ObjAttribId",    14,          "Идентификатор реквизита объекта, к которому привязан реквизит"),
    new Array("ValueType",      15,          "Идентификатор типа значения"),
    new Array("ValueLength",    16,          "Длина значения"),
    new Array("ValuePrecision", 17,          "Точность значения"),
    new Array("ValueKind",      18,          "Идентификатор вида значения"),
    new Array("ValueFlags",     19,          'Набор двух флагов - "Не отрицательный" и "Разделять на триады"'),
    new Array("Flags2",         20,          "Основной набор флагов"),    
    new Array("Description",    22,          "Пользовательское описание элемента управления"),
    new Array("Hint",           23,          "Подсказка для пользователя"),
    new Array("Font",           "//ToDo!!!", "Шрифт элемента управления"),
    new Array("PictureId",      41,          "Идентификатор картинки"),
    new Array("Layer",          42,          "Слой, на котором располагается элемент управления"),    
    new Array("HotKey",         43,          "Горячая клавиша, ассоциированная с элементом управления")
);

var xml = "";
var imp = "";
var tst = "";

for (var i in a) {
    xml += '    <property name="' + a[i][0] + '">' + "\n"
        +  '        <get />' + "\n"
        +  '        <put />' + "\n"
        +  '        <comment><![CDATA[' + "\n"
        +  a[i][2] + "\n"
        +  '         ]]></comment>' + "\n"
        +  '    </property>' + "\n\n";

    imp += 'function get_' + a[i][0] + '()' + "\n"
        +  '{' + "\n"
        +  '    return get_Property(' + a[i][1] + ");\n"
        +  '}' + "\n\n"

        +  'function put_' + a[i][0] + '('+ a[i][0] + ')' + "\n"
        +  '{' + "\n"
        +  '    put_Property(' + a[i][1] + ', ' + a[i][0] + ");\n"
        +  '}' + "\n\n";

    tst += '/**' + "\n"
        +  ' * property ' + a[i][0] + ' as String, Read/Write' + "\n"
        +  ' *  Member of OpenConfDialogControl' + "\n"
        +  ' *      ' + a[i][2] + "\n"
        +  ' */' + "\n\n"       
        +  '//print("Проверка получения свойства ' + a[i][0] + ': " + assign(ctl.' + a[i][0] + '==""));' + "\n\n"
        +  '//ctl.' + a[i][0] + '= "";' + "\n"
        +  '//print("Проверка установки свойства ' + a[i][0] + ': " + assign(ctl.' + a[i][0] + '==""));' + "\n\n";
}

print(xml);
print(imp);
print(tst);




