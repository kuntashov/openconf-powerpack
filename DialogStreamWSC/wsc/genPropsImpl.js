function print(str)
{
    WScript.Echo(str);
}

var a = new Array(
//=========== �������������  ����� �-��     �����������
    new Array("Caption",         1,          "��������� �������� ����������"),
    new Array("Type",            2,          "��� �������� ����������"),
    new Array("Flags1",          3,          "����� �������������� ������ �������� ����������"),
    new Array("X",               4,          "���������� X"),
    new Array("Y",               5,          "���������� Y"),
    new Array("Width",           6,          "������ �������� ����������"),
    new Array("Height",          7,          "������ �������� ����������"),
    new Array("Associated",      8,          "������� �������� � ��������� ������� ������������ (1-��������, 2 - ���)"),
    new Array("TabOrder",       10,          "������� ������"),
    new Array("Action",         12,          "�������"),
    new Array("Name",           13,          "������������� �������� ����������"),
    new Array("ObjAttribId",    14,          "������������� ��������� �������, � �������� �������� ��������"),
    new Array("ValueType",      15,          "������������� ���� ��������"),
    new Array("ValueLength",    16,          "����� ��������"),
    new Array("ValuePrecision", 17,          "�������� ��������"),
    new Array("ValueKind",      18,          "������������� ���� ��������"),
    new Array("ValueFlags",     19,          '����� ���� ������ - "�� �������������" � "��������� �� ������"'),
    new Array("Flags2",         20,          "�������� ����� ������"),    
    new Array("Description",    22,          "���������������� �������� �������� ����������"),
    new Array("Hint",           23,          "��������� ��� ������������"),
    new Array("Font",           "//ToDo!!!", "����� �������� ����������"),
    new Array("PictureId",      41,          "������������� ��������"),
    new Array("Layer",          42,          "����, �� ������� ������������� ������� ����������"),    
    new Array("HotKey",         43,          "������� �������, ��������������� � ��������� ����������")
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
        +  '//print("�������� ��������� �������� ' + a[i][0] + ': " + assign(ctl.' + a[i][0] + '==""));' + "\n\n"
        +  '//ctl.' + a[i][0] + '= "";' + "\n"
        +  '//print("�������� ��������� �������� ' + a[i][0] + ': " + assign(ctl.' + a[i][0] + '==""));' + "\n\n";
}

print(xml);
print(imp);
print(tst);




