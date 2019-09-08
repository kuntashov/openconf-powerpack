/* ===================================================================================
	������: TelepatSettings.js
	������: $Revision: 1.2 $
	 
	������������ ���������� �������� �������� �������� ��� ������ ��������������
 ���� (��).

	��������� �������� � �������� ���� � ��������� ����� telepat.prm.

	���������� ��������� (������� ������������ �� ��������� � ��� �������, �����
��� ������� ���� ����� �������� �� ����������) �������� � ����� telepat.prm � 
�������� bin\config �������� ��������� 1�.

	��� �������� ������������� ��� �����-���� �� ������� ������������
������� ��������� ��������� ��������, �������� ��� ������� ����. ���� ���������
��� ������� �� �� ������ (�� ���������� ����� telepat.prm � �������� ��), ��
����������� ���������� ���������.

	��� ���������� � �������� �������� �������� ������������ ��������� �������:

	SaveSettings()
		��������� ��������� �������� ��� ������� ��

	LoadSettings()
		��������� ��������� �������� ��� ������� ��
	                    
	SaveGlobalSettings()
		��������� ���������� ��������� ��������

	LoadGlobalSettings()
		��������� ���������� ��������� ��������

	����� ����, �� ��������� ������ ������������� ��������� ��������� ��������
��� �������� �������������. ���� ��� �� ���������� ����� ���������, �� 
���������� �������� ���������� AutoSaveSettingsOnQuit (��. ����) � �������� false.

	������� �� ������������� ������� :-)
		
	�����: ��������� �������� <kuntashov@yandex.ru>

=================================================================================== */

// ��������� ��������� �������� ��� �������� �������������
// true - ���������, false - �� ���������

var AutoSaveSettingsOnQuit = true

var CoreObj = {

	telepat : null, 
	
	settings : ['Components', 'Language', 'UseStdMethodDlg', 'NoOrderMethodDlg', 'FilterMethodDlg', 
				'AutoParamInfo', 'ParamInfoAddMethDescr', 'ParamInfoAddParamDescr', 
				'AutoActiveCountSymbols', 'DisableTemplateInRemString', 'AddTemplate'],
			  
	Init : function () {
		this.telepat = Plugins('�������');
		if (!this.telepat) {
			// ������� �� ����������, ��������� ����
			Scripts.Unload(SelfScript.Name);
		}
		// ���� �� ���������� "����������" ����� ��������, �� 
		// ���������� ����������
		if (!this.LoadSettings(IBDir)) {
			this.LoadSettings(BinDir + "config");
		}
	},

	Unload : function ()
	{
		this.telepat = null;
	},
	
	SaveSettings : function (dir)
	{		
		var file = this.OpenSettingsFile(dir, true);
		for (var i=0; i<this.settings.length; i++) {
			file.WriteLine('telepat.' + this.settings[i] + " = " + this.telepat[this.settings[i]]);
		}		
		file.Close();
	},

	LoadSettings : function (dir)
	{
		var file = this.OpenSettingsFile(dir, false);
		if (!file) return false;
		with (this) eval(file.ReadAll());
		file.Close();
		return true;
	},

	OpenSettingsFile : function (dir, forWriting)
	{
		with (new ActiveXObject('Scripting.FileSystemObject')) {
			var path = BuildPath(dir, 'telepat.prm');
			if (!(FileExists(path) || forWriting)) {
				return null;
			}
			return OpenTextFile(path, forWriting?2:1, true);
		}		
	}
}

function Configurator::OnQuit()
{
	if (AutoSaveSettingsOnQuit) {
		SaveSettings();
	}
	CoreObj.Unload();
}

function SaveSettings() { CoreObj.SaveSettings(IBDir) }
function LoadSettings() { CoreObj.LoadSettings(IBDir) }

function SaveGlobalSettings() { CoreObj.SaveSettings(BinDir + "config") }
function LoadGlobalSettings() { CoreObj.LoadSettings(BinDir + "config") }

CoreObj.Init()
