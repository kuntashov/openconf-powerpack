function addButton()
{
    var frm = new ActiveXObject("OpenConf.DialogObject");
    var stream = Windows.ActiveWnd.Document.Page(0).Stream;

    if (!stream)
        return;     

    frm.Stream = stream;

    var btn = frm.CreateControl("BUTTON");
    btn.Caption = "Hello!";

    frm.Controls.Add(btn);

    Windows.ActiveWnd.Document.Page(0).Stream = frm.Stream;
}