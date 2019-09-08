'$NAME OLE-ActiveX *.ints Generator

'	intsOLEGenerator.vbs - ��������� *.ints ������ ��� Intellisence
'	� trdm 2005
'
' ������: $Revision: 1.4 $
'
'����� ������ �������, ��� trdm 2005 ���
'	trdm@mail333.com
'	ICQ 308-779-620
'
' ��� ������ ������� ������� ��������� ���������� TLBINF32.DLL
' �� ��������� ���������� ��� ������ � ������ 6-� ������ ������ �� microsoft, ������
' ������� � �����: �� �����, � ��� ���� �� ���: 2 �����
' http://download.microsoft.com/download/vstudio60pro/doc/1/win98/en-us/tlbinf32.exe
' http://support.microsoft.com/default.aspx?scid=kb;en-us;224331

Dim FileNameProgIDDumped
Dim DataDir
Dim WSH
Dim FSO

DataDir = BinDir & "\Config\Intell"
FileNameProgIDDumped		= DataDir + "\ProgIDDumped.txt"			' ��� �����, ����������� ����-��� �� ������� ������ ������������


DataDirAls = BinDir ' ��� Bin....

Dim AlsToFile


'::::::::::::::::::::::::::::::::::::::::::::::::::::::
' �������� � ����� ������� nArray ������� nElement
' � �� ���-�� ��� ������ ��������� ������ ��������........
' ���������:
' nArray - ������ ���� ���������
' nElement - ������� ��� �� ��������� � ������
Private Sub AppendArray ( nArray, nElement )
	On Error Resume Next
	Size = 1 + UBound(nArray)
	NewArr = Array()
	ReDim NewArr( Size )
	if Size>0 Then
		For i = 0 To Size-1
			If IsObject( nElement ) Then
				Set NewArr(i) = nArray(i)
			Else
				NewArr(i) = nArray(i)
			End If
		Next
	End If
	ReDim nArray( Size )
	if Size>0 Then
		For i = 0 To Size-1
			If IsObject( nElement ) Then
				Set nArray(i) = NewArr(i)
			Else
				nArray(i) = NewArr(i)
			End If
		Next
	End If
	Set nArray ( Size ) = nElement

	if Err.number<>0 Then
		Message " ������ ��� ���������� � ������...", mExclamation3
		Message err.Description, mExclamation3
		On Error goto 0
	End If
	On Error goto 0

End Sub

'::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ���� ���� � ���������� �� ProgID-�
Private Function GetPathFromProgID ( ProgID )
	GetPathFromProgID = Null
	RKey = "HKCR\"+ProgID+"\CLSID\"
	dim PathFile
	Dim val
	PathFile = Null
	On Error Resume Next
	val = WSH.RegRead(RKey)
	if Not IsEmpty(val) Then
		RKey = "HKCR\CLSID\"+val+"\InprocServer32\"
		KeyVal = WSH.RegRead(RKey)
		if IsEmpty(KeyVal) Then
			RKey = "HKCR\CLSID\"+val+"\LocalServer\"
			KeyVal = WSH.RegRead(RKey)
			if IsEmpty(KeyVal) Then
				RKey = "HKCR\CLSID\"+val+"\LocalServer32\"
				KeyVal = WSH.RegRead(RKey)
			End IF
		End IF
		if Not IsEmpty(KeyVal) Then
			PathFile = KeyVal
		End IF
	Else
		Message "�� ������ ���� ������� � ������ ��������" + RKey + " ��� �������: " + ProgID
	End IF
	if IsNull(PathFile) Then
		On Error goto 0
		Exit Function
	End IF
	GetPathFromProgID = PathFile
	On Error goto 0
End Function

'::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ������� ��������� ����-�� �� ������� � �� ���������� ProgID
' ProgID = ������ ���� ����������� ����������� ����������, ������� � ������.
Function IsProgID( ProgID )
	nProgID = Trim(ProgID)
	'�������� ��� ��� �	ProgID �����������....
	IsProgID = True
	For i=1 To Len (nProgID)
		Char = Mid (nProgID,i,1)
		ChKd = Asc(Char)
		if Not ((ChKd = Asc(".")) or ( (Asc("A")<=ChKd) And (Asc("Z")>=ChKd)) or ( (Asc("a")<=ChKd) And (Asc("z")>=ChKd)) or ( (Asc("0")<=ChKd) And (Asc("9")>=ChKd)) ) Then
			IsProgID = False
			'message "������������ ������ � ProgID: " + Char
			Exit Function
		End IF
	Next
	PathFile = GetPathFromProgID ( ProgID )
	If PathFile = Null Then
		IsProgID = False
		Exit Function
	End If
	IsProgID = True
End Function

'::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ���� nProgID � ����� (BinDir+"\Config\Intell\ProgIDDumped.txt"), ���� �������,
' �� ���������, ��� ������ ��� ��������������� � ����� �� ���� �������������.
Function ProgIDIsDumped( nProgID )
	ProgIDIsDumped = False
	FileName = FileNameProgIDDumped
	if fso.FileExists( FileName ) Then
		Set File0 = fso.GetFile( FileName )
		Set File = File0.OpenAsTextStream(1)
		if File0.Size > 0 Then
			TextStream = File.ReadAll()
			If InStr(1, lCase(TextStream), lcase(nProgID))>0 Then
				ProgIDIsDumped = True
			End IF
		End IF
		File.Close
	End IF
End Function

'::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ����� ��� �������
'
' ���������:
' ---------------------------
' Init (nName*, nAlias*)	- ����������� ��� � ���������� ���������...
' Dump ()				- �������� ������ � ���� "AlsToFile" � ������������ � ��������.
' ---------------------------
'  mWriting		- ���������*
'  mWritingEng	- ��������� ����������
'
'  mSyntax		- ���������*
'  mAssignment	- ����������
'  mParams		- ���������
'  mNotes		- ����������
'_______________________________________
'  ����������: mMamber*	- ������������ �������
Class ALSItem ' ����� ��� ���������
	Public mName	' ��� ��������
	Public mAlias	' �������...

	Public mWriting		' ���������*
	Public mWritingEng	' ��������� ����������

	Public mSyntax		' ���������....
	Public mAssignment	' ����������*
	Public mParams		' ���������
	Public mNotes		' ����������

	Public mLevel	' ������� �����������....

	Sub Init (nName, nAlias)
		mName	= nName
		mAlias	= nlias
	End Sub

	Private Sub Class_Initialize
		mName		=	"!!!Name!!!"
		mAlias		= "!!!Alias!!!"
		mWriting	= "!!!Writing!!!"
		mWritingEng = ""
		mSyntax		= ""
		mAssignment	= ""
		mParams		= ""
		mNotes		= ""
	End Sub



	' ���� �������� � ������ - ������ �������, �������� �����������...
	Private Function Normalize()
		Normalize = True

		' ������� �������, �������� �� ��������
		mName		= Replace( mName, """", "''")
		mAlias		= Replace( mAlias, """", "''")
		mWriting	= Replace( mWriting , """", "''")
		mWritingEng	= Replace( mWritingEng , """", "''")
		mSyntax		= Replace( mSyntax , """", "''")
		mAssignment	= Replace( mAssignment , """", "''")
		mParams		= Replace( mParams , """", "''")
		mNotes		= Replace( mNotes , """", "''")

	End Function

	' ��������� � ����.....................
	Function Dump()

		Dump	= Normalize
		Indent	= String(mLevel," ")

		' ������ ������
		StrToDump = Indent + "{""Item"",""AST"",""" + mName + """,""" + mAlias + ""","""+ mWriting + ""","""+ mWritingEng + ""","
		AllDump = AllDump + StrToDump + vbCrLf
		AlsToFile.WriteLine StrToDump
		AlsToFile.WriteLine """@���������:"
		if Len(mSyntax) = 0 Then
			AlsToFile.WriteLine "-------------------"
		Else
			AlsToFile.WriteLine mSyntax
		End IF

		if Len(mAssignment)>0 Then
			AlsToFile.WriteLine "@����������:"
			AlsToFile.WriteLine mAssignment
		Else
			'AlsToFile.WriteLine "----------------"
		End IF
		if Len(mParams) Then
			AlsToFile.WriteLine "@���������:"
			AlsToFile.WriteLine mParams
		End IF
		if Len(mNotes) Then
			AlsToFile.WriteLine "@���������:"
			AlsToFile.WriteLine mNotes
		End IF

		AlsToFile.WriteLine """"
		StrToDump = Indent + "}," ' ���� � Item ����������� �������...
		AlsToFile.WriteLine StrToDump
	End Function

End Class 'ALSItem

'::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ����� ��� ��������
'
' ���������:
' ---------------------------
' Init (nName*, nAlias*)	- ����������� ��� � ���������� ���������...
' Dump ()				- �������� ������ � ���� "AlsToFile" � ������������ � ��������.
' AddSubFolder ( nFolderName, nFolderAlias )	- ��������� "������"
' AddSubItem( nItemName, nItemAlias )			- ��������� "������"
'_______________________________________
'  ����������: mMamber*	- ������������ �������
Class ALSFolder ' ����� ��� �����
	Public mName		' ��� �����
	Public mAlias	' �������...
	Public mLevel	' ������� �����������....

	Public mArrSubFolder	'������ ��������
	Public mArrSubItem	'������ ��� ���������



	Private Sub Class_Initialize
		mName = ""
		mArrSubFolder = Array()
		mArrSubItem = Array()
		Level = 0
	End Sub

	Sub Init (nFolderName, nFolderAlias)
		mName	= nFolderName
		mAlias	= nFolderAlias
		mName = Replace(mName, """", "''")
		mAlias = Replace(mAlias, """", "''")
	End Sub


	Function AddSubFolder( nFolderName, nFolderAlias )

		On Error Resume Next
		Set NewFolder = New ALSFolder
		NewFolder.mLevel = mLevel + 1
		NewFolder.Init nFolderName, nFolderAlias

		AppendArray mArrSubFolder, NewFolder
		if Err.number<>0 Then
			Message "�� ������� ������� �����: " + nFolderName, mExclamation3
			Message err.Description, mExclamation3
			On Error goto 0
			Exit Function
		End If
		On Error goto 0


		Set AddSubFolder = NewFolder
	End Function

	Function AddSubItem( nItemName, nItemAlias )

	On Error Resume Next
		Set NewALSItem = New ALSItem
		NewALSItem.mLevel = mLevel + 1
		NewALSItem.Init nItemName, nItemAlias
		AppendArray mArrSubItem, NewALSItem
		Set AddSubItem = NewALSItem
		if Err.number<>0 Then
			Message "�� ������� ������� �����: " + nItemName, mExclamation3
			Message err.Description, mExclamation3
			On Error goto 0
		End If
		On Error goto 0

	End Function

	Function Dump(  )
		Dump = False
		If (UBound(mArrSubFolder) + UBound (mArrSubItem)) >-2 Then
			' ��������� �������...
			Indent = String(mLevel," ")
			' ������ ������
			StrToDump = Indent + "{""Folder"",""AST"",""" + mName + """,""" + mAlias + ""","
			AlsToFile.WriteLine StrToDump

			For i=0 To UBound (mArrSubItem)
				Set nALSItem = mArrSubItem(i)
				Dump = nALSItem.Dump
			Next

			For i=0 To UBound (mArrSubFolder)
				Set nFolder = mArrSubFolder(i)
				Dump = nFolder.Dump
			Next
			' ����� ������
			StrToDump = Indent + "},"
			AlsToFile.WriteLine StrToDump
		End If

	End Function

End Class 'ALSFolder

'::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ����� ��� ��������� als-������ ��������� � *.ints �������
'
' ���������:
' ---------------------------
' Init(nFileName*)		- ������������� ������!! ��� ����� ��� ���������
' Dump ()				- �������� ������ � ���� "AlsToFile" � ������������ � ��������.
' CreateRootFolder( nFolderName*, nFolderAlias*)	- ��������� �������� "������",
'		� � ��������� ������� ��� ����� ���������� ��� � ALSFolder-�
'_______________________________________
'  ����������: mMamber*	- ������������ �������

'ALSGenerator
Class ALSGenerator ' ��� ����� ��������� als-������....
	Public Redy				' ����� ����� � ������...
	Public RootFolder		' �������� �����...
	Public mFileName		' ��� �����

	' ���� ���������� ������
	Private Sub Class_Terminate	' Setup Terminate event.
		'��������� ����...
		On Error Resume Next
		if IsObject(AlsToFile) Then AlsToFile.Close

		if Err.number<>0 Then
			Message "�� ������� ������� ����" + mFileName, mExclamation3
			Message err.Description, mExclamation3
			On Error goto 0
		End If
		On Error goto 0
	End Sub

	Private Sub Class_Initialize
		Redy		= False
		mFileName	= ""
	End Sub

	' ���������������� �������������...
	' ��������: nFileName - ��� �����, ������� ��������� �������
	Function Init(nFileName)
		Init = True
		mFileName = nFileName

		Redy = True
	End Function

	Function CreateRootFolder( nFolderName, nFolderAlias) '(ParentFolder, nFolderName,  nFolderName)
		RootFolder = Null
		On Error Resume Next
		Set RootFolder = New ALSFolder
		RootFolder.mLevel = 1
		Set CreateRootFolder = RootFolder
		RootFolder.Init nFolderName, nFolderAlias

		if Err.number<>0 Then
			Message "�� ������� ������� �����: " + nFolderName, mExclamation3
			Message err.Description, mExclamation3
			On Error goto 0
		End If
		On Error goto 0

	End Function

	Function Dump()
		Dump = False

		StrQuestion = "������ ������ ��� ��������� als-�����!" + vbCr +_
		mFileName + vbCr +_
		"������������?"
		Dim MyAnswer
		MyAnswer = MsgBox (StrQuestion, vbOKCancel+vbQuestion, "OLE-ActiveX *.ints Generator")
		if MyAnswer = vbCancel Then
			Dump = True
			Exit Function
		End If
		On Error Resume Next
		if fso.FileExists(mFileName) Then fso.DeleteFile mFileName
		Set AlsToFile = fso.CreateTextFile( mFileName )
		'if Exept("�� ������� �������� ����" + mFileName) Then Exit Function

		if Err.number<>0 Then
			Message "�� ������� �������� ����" + mFileName, mExclamation3
			Message err.Description, mExclamation3
			On Error goto 0
			Exit Function
		End If
		On Error goto 0


		DeleteFile = false
		if not IsEmpty(RootFolder) Then
			AlsToFile.WriteLine "{""Shell"","
			if Not IsNull(RootFolder) Then Dump = RootFolder.Dump
			AlsToFile.WriteLine "}"
		else
			DeleteFile = true
		End if
		AlsToFile.Close
		AlsToFile = Null
		If len(mFileName)>0 And DeleteFile Then fso.DeleteFile mFileName
		if Err.number<>0 Then
			Message "��������� als-����� " + mFileName + "����������� �������...", mExclamation3
			Message err.Description, mExclamation3
			On Error goto 0
			Dump = False
		End If
		On Error goto 0


		if Dump Then Message "������������ als-file: " + mFileName

	End Function

End Class


' ����� ��� ������� ����������.
Class ITypeLib
	Public LibName '��� ���������
	Public LibPath '��� ���������
	Public Library '��������� (TypeLibInfo)
	Public DefaultInterfaseName ' ��������� �� ���������
	Public ThisProgID 'ProgID ����������
	Public ThisProgIDForIntel 'ProgID ����������, ��������������� ��� OtherTypesDefine:  Excel.Application>>excel_application

	Public LibraryIsLoad ' ���������� ������� � ������ � ������.

	Public DebugFlag '���� �������
	Public flOutNameAndTypeParameters '�������� �� ���������� � ���������� ������� (����� � ����)
	Public AllDumpInterfases '��� ���������� ����������.

	Dim DataDir							' �������, ��� �������� *.ints �����
	Dim FileNameTypesStructureExt		' ��� ����� �� �������� ����� ������������ ������� Intellisence "TypesStructureExt"
	Dim FileNameTypesStructureExtDef	' ��� ����� �� �������� ����� ������������ ������� Intellisence "TypesStructureExt"
	' �������� ������������ ���������� ��� ������ �� TypesStructureExt, ��� �������� ������ � �������-���������
	Dim FileNameOtherTypesDefine		' ��� ����� �� �������� ����� ������������ ������� Intellisence "OtherTypesDefine"
	Dim FileNameProgIDDumped			' ��� �����, ����������� ����-��� �� ������� ������ ������������
	' ��-�� ����, ��� ������, ��������� �� ����-��� ����� ���� ������
	' ��� ��� ���� �������������� *.ints �����

	Public TypeStr

	Public ErrorNumber				' ����� ��������� ������ ��������.

	'Public ALSCreate				' ����� �� ������������ als-����?
	Public mALSGenerator			' ������� ������ ��� ��������� als-�����....
	Public mALSParentF				' ������������ �����...

	Public TLITypeLibInfo
	Public TLIApplication

	Dim INVOKE_UNKNOWN
	Dim INVOKE_FUNC
	Dim INVOKE_PROPERTYGET
	Dim INVOKE_PROPERTYPUT
	Dim INVOKE_PROPERTYPUTREF
	Dim INVOKE_EVENTFUNC
	Dim INVOKE_CONST

	Dim VT_EMPTY
	Dim VT_NULL
	Dim VT_I2
	Dim VT_I4
	Dim VT_R4
	Dim VT_R8
	Dim VT_CY
	Dim VT_DATE
	Dim VT_BSTR
	Dim VT_DISPATCH
	Dim VT_ERROR
	Dim VT_BOOL
	Dim VT_VARIANT
	Dim VT_UNKNOWN
	Dim VT_DECIMAL
	Dim VT_I1
	Dim VT_UI1
	Dim VT_UI2
	Dim VT_UI4
	Dim VT_I8
	Dim VT_UI8
	Dim VT_INT
	Dim VT_UINT
	Dim VT_VOID
	Dim VT_HRESULT
	Dim VT_PTR
	Dim VT_SAFEARRAY
	Dim VT_CARRAY
	Dim VT_USERDEFINED
	Dim VT_LPSTR
	Dim VT_LPWSTR
	Dim VT_RECORD
	Dim VT_FILETIME
	Dim VT_BLOB
	Dim VT_STREAM
	Dim VT_STORAGE
	Dim VT_STREAMED_OBJECT
	Dim VT_STORED_OBJECT
	Dim VT_BLOB_OBJECT
	Dim VT_CF
	Dim VT_CLSID
	Dim VT_VECTOR
	Dim VT_ARRAY
	Dim VT_BYREF
	Dim VT_RESERVED

	Dim TKIND_ENUM
	Dim TKIND_RECORD
	Dim TKIND_MODULE
	Dim TKIND_INTERFACE
	Dim TKIND_DISPATCH
	Dim TKIND_COCLASS
	Dim TKIND_ALIAS
	Dim TKIND_UNION
	Dim TKIND_MAX



	' ������� ����������
	' CreateObj - ��������� ������, ���� �� ����-��� �� �����
	' ProgID - ����-�� �������
	Function LoadLibrary( ProgID, CreateObj )
		LoadLibrary = False
		LibraryIsLoad = False
		if ProgIDIsDumped( ProgID ) Then
			Message "����� ��� " + ProgID + " ��� ������������. ���� ����� ������ ��������������� ��, �������� ������: "+ProgID+" �� �����: " + FileNameProgIDDumped
			Exit Function
		End If

		ThisProgID = ProgID

		' ���� ����-�� � �������.
		LibPath = GetPathFromProgID ( ProgID )
		if IsNull(LibPath) Then
			ErrorNumber = 1
			SaveDumpedProgID lCase(ThisProgID)
			Message "����-�� " + ThisProgID + " �� ������ � ������� � ������� ��� ������������ ��� ���������� ��������� ���������.", mExclamation
			Exit Function
		End IF
		LibPath = Replace(LibPath,"/automation","")
		LibPath = Trim(LibPath)

		if not FSO.FileExists(LibPath) Then
			ErrorNumber = 1
			SaveDumpedProgID lCase(ThisProgID)
			Message "����-�� " + ThisProgID + "��� �����:" + LibPath	, mExclamation
			LibPath = Null
		End If
		If not IsNull(LibPath) then
			On Error Resume Next
			Set Library = TLIApplication.TypeLibInfoFromFile(LibPath)
			if Err.number = 0 Then
				If not IsNull(Library) And Not IsEmpty( Library ) Then
					LoadLibrary = true
				End if
			End if
			On Error Goto 0
		End if
		LoadLibrary = False


		StrQuestion = "��������! ������ ����� ����������� �������� ������� """+ProgID+""""+vbCr+_
		"��� ��������� ������ Intellisence. ��������� ������� �� ������������ "+vbCr+_
		"������������ �������� � ����� �������� � ���������� ���������."+vbCr+_
		"������������� ��������� ������." + vbCr +_
		"����������?"
		Dim MyAnswer
		MyAnswer = MsgBox (StrQuestion, vbOkCancel+vbQuestion, "OLE-ActiveX *.ints Generator")
		if MyAnswer = vbCancel Then
			LoadLibrary = False
			LibraryIsLoad = LoadLibrary
			StrQuestion = "�������� ������ """+ProgID+""" ��� ������������?"+ vbCr+_
			"�� ���������� �������� ��������� �� ������������."

			MyAnswer = MsgBox (StrQuestion, vbYesNo+vbQuestion, "OLE-ActiveX *.ints Generator")
			if MyAnswer = vbYes Then
				SaveDumpedProgID lCase(ThisProgID)
				ErrorNumber = 1
			End If
			Exit Function
		End If

		' ����� ������� ������ �� �������� �������� ��������� ��������� ����������� ���
		' ���� ��� ��� ��������� ��� �������� ��������� ��������� ��� �������� ������� ���� �����������.
		On Error Resume Next
		Set Obj = CreateObject(ProgID)
		if Err.number<>0 Then
			LoadLibrary = false
			LibraryIsLoad = false
			ErrorNumber = 1
			SaveDumpedProgID lCase(ThisProgID)
			Message "�� ����-�� " + ThisProgID + " ������� ������ �� �������", mExclamation
			On Error goto 0
			Exit Function
		End If
		On Error goto 0

		if not IsEmpty(Obj) Then
			On Error Resume Next
			Set DefaultInterfase = TLIApplication.InterfaceInfoFromObject(Obj)
			if Err.number<>0 Then
				LoadLibrary = false
				LibraryIsLoad = false
				ErrorNumber = 1
				SaveDumpedProgID lCase(ThisProgID)
				Message "�� ����-�� " + ThisProgID + " �� ������� �������� �������� ����������", mExclamation
				On Error goto 0
				Exit Function
			End If
			On Error goto 0

			if Not IsEmpty( DefaultInterfase ) Then
				On Error Resume Next
				LibPath = DefaultInterfase.Parent.ContainingFile
				Set Library = TLIApplication.TypeLibInfoFromFile(LibPath)
				DefaultInterfaseName = DefaultInterfase.Name
				LoadLibrary = true
				LibraryIsLoad = true
				if Err.number<>0 Then
					LoadLibrary = false
					LibraryIsLoad = false

					ErrorNumber = 1
					SaveDumpedProgID lCase(ThisProgID)
					Message "�� ����-�� " + ThisProgID + " ���������� �� ���������� �����..", mExclamation
					On Error goto 0
					Exit Function
				On Error goto 0
				End If
			On Error goto 0
			End IF
			Obj = Null
		End IF
		LibraryIsLoad = LoadLibrary
		ThisProgIDForIntel = lCase(ThisProgID)
		ThisProgIDForIntel = Replace(ThisProgIDForIntel,".", "_")
	'Stop
		OtherTypesDefineString = lCase(ThisProgID)+","+ThisProgIDForIntel
		DumpStringToFile OtherTypesDefineString, FileNameOtherTypesDefine
		SaveDumpedProgID ThisProgID
	End Function


	' ���������� ���������� ������ � ����.
	Private Sub DumpStringToFile (DumpString, FileName)
		if Not fso.FileExists( FileName ) Then
			Set File = fso.CreateTextFile( FileName, False)
		Else
			Set File0 = fso.GetFile( FileName )
			Set File = File0.OpenAsTextStream(1)
			if File0.Size > 0 Then
				TextStream = File.ReadAll()
				If InStr(1, lCase(TextStream), lcase(DumpString))>0 Then
					File.Close
					Exit Sub
				End IF
			End IF
			Set File = File0.OpenAsTextStream(8)

		End IF
		File.WriteLine DumpString
		File.Close
	End Sub


	' ���������� ������������ ProgID, �� �������� ������������� ����� *.ints
	'private ' ���� �� ������������, ����������
	Sub SaveDumpedProgID( DumpedProgID )
		DumpStringToFile DumpedProgID, FileNameProgIDDumped
	End Sub


	Private  Sub DMessage( mess )
		if DebugFlag Then
			Message mess
		End If
	End Sub

	function CheskClass() '�������� ����, ��� �� �����������������...
		CheskClass = true
		If IsEmpty(TLITypeLibInfo) Then
			CheskClass = false
		ElseIf IsEmpty(TLIApplication) Then
			CheskClass = false
		ElseIf IsEmpty(WSH) Then
			CheskClass = false
		End IF
	End Function

	Private Sub Class_Initialize
		LibName = null
		On Error Resume Next
		Set TLITypeLibInfo = CreateObject("TLI.TypeLibInfo")
		Set TLIApplication = CreateObject("TLI.TLIApplication")
		if Err.number<>0 then
			Message "�����: ITypeLib �� �����������������, ����������: TLBINF32.DLL �� ���������������� �� ������ ������� �����!"
		End If
		On Error Goto 0
		Set TypeStr = CreateObject("Scripting.Dictionary")
		Set AllDumpInterfases = CreateObject("Scripting.Dictionary")
		DefaultInterfase = Null
		DebugFlag = false
		flOutNameAndTypeParameters = False

		DataDir = BinDir & "\Config\Intell"
		FileNameTypesStructureExt		= DataDir + "\TypesStructureExt.txt"	' ��� ����� �� �������� ����� ������������ ������� Intellisence "TypesStructureExt"
		FileNameTypesStructureExtDef	= DataDir + "\TypesStructureExtDef.txt"	' ��� ����� �� �������� ����� ������������ ������� Intellisence "TypesStructureExtDef"
		FileNameOtherTypesDefine		= DataDir + "\OtherTypesDefine.txt"		' ��� ����� �� �������� ����� ������������ ������� Intellisence "OtherTypesDefine"
		FileNameProgIDDumped			= DataDir + "\ProgIDDumped.txt"			' ��� �����, ����������� ����-��� �� ������� ������ ������������
		ErrorNumber = 0



		INVOKE_UNKNOWN = 0
		INVOKE_FUNC = 1
		INVOKE_PROPERTYGET = 2
		INVOKE_PROPERTYPUT = 4
		INVOKE_PROPERTYPUTREF = 8
		INVOKE_EVENTFUNC = 16
		INVOKE_CONST = 32

		VT_EMPTY = 0
		VT_NULL = 1
		VT_I2 = 2
		VT_I4 = 3
		VT_R4 = 4
		VT_R8 = 5
		VT_CY = 6
		VT_DATE = 7
		VT_BSTR = 8
		VT_DISPATCH = 9
		VT_ERROR = 10
		VT_BOOL = 11
		VT_VARIANT = 12
		VT_UNKNOWN = 13
		VT_DECIMAL = 14
		VT_I1 = 16
		VT_UI1 = 17
		VT_UI2 = 18
		VT_UI4 = 19
		VT_I8 = 20
		VT_UI8 = 21
		VT_INT = 22
		VT_UINT = 23
		VT_VOID = 24
		VT_HRESULT = 25
		VT_PTR = 26
		VT_SAFEARRAY = 27
		VT_CARRAY = 28
		VT_USERDEFINED = 29
		VT_LPSTR = 30
		VT_LPWSTR = 31
		VT_RECORD = 36
		VT_FILETIME = 64
		VT_BLOB = 65
		VT_STREAM = 66
		VT_STORAGE = 67
		VT_STREAMED_OBJECT = 68
		VT_STORED_OBJECT = 69
		VT_BLOB_OBJECT = 70
		VT_CF = 71
		VT_CLSID = 72
		VT_VECTOR = 4096
		VT_ARRAY = 8192
		VT_BYREF = 16384
		VT_RESERVED = 32768

		TypeStr.Add VT_EMPTY ,"VT_EMPTY"
		TypeStr.Add VT_NULL ,"VT_NULL"
		TypeStr.Add VT_I2 ,"����� (VT_I2)"
		TypeStr.Add VT_I4 ,"����� (VT_I4)"
		TypeStr.Add VT_R4 ,"����� (VT_R4)"
		TypeStr.Add VT_R8  ,"����� (VT_R8)"
		TypeStr.Add VT_CY ,"VT_CY"
		TypeStr.Add VT_DATE  ,"���� (VT_DATE)"
		TypeStr.Add VT_BSTR ,"������ (VT_BSTR)"
		TypeStr.Add VT_DISPATCH ,"VT_DISPATCH"
		TypeStr.Add VT_ERROR ,"VT_ERROR"
		TypeStr.Add VT_BOOL ,"������ (VT_BOOL)"
		TypeStr.Add VT_VARIANT ,"VT_VARIANT"
		TypeStr.Add VT_UNKNOWN ,"VT_UNKNOWN"
		TypeStr.Add VT_DECIMAL ,"����� (VT_DECIMAL)"
		TypeStr.Add VT_I1 ,"����� (VT_I1)"
		TypeStr.Add VT_UI1 ,"����� (VT_UI1)"
		TypeStr.Add VT_UI2 ,"����� (VT_UI2)"
		TypeStr.Add VT_UI4 ,"����� (VT_UI4)"
		TypeStr.Add VT_I8 ,"����� (VT_I8)"
		TypeStr.Add VT_UI8 ,"����� (VT_UI8)"
		TypeStr.Add VT_INT ,"����� (VT_INT)"
		TypeStr.Add VT_UINT ,"����� (VT_UINT)"
		TypeStr.Add VT_VOID,"VT_VOID"
		TypeStr.Add VT_HRESULT,"VT_HRESULT"
		TypeStr.Add VT_PTR,"VT_PTR"
		TypeStr.Add VT_SAFEARRAY ,"VT_SAFEARRAY"
		TypeStr.Add VT_CARRAY ,"VT_CARRAY"
		TypeStr.Add VT_USERDEFINED,"VT_USERDEFINED"
		TypeStr.Add VT_LPSTR,"VT_LPSTR"
		TypeStr.Add VT_LPWSTR,"VT_LPWSTR"
		TypeStr.Add VT_RECORD,"VT_RECORD"
		TypeStr.Add VT_FILETIME,"VT_FILETIME"
		TypeStr.Add VT_BLOB,"VT_BLOB"
		TypeStr.Add VT_STREAM,"VT_STREAM"
		TypeStr.Add VT_STORAGE,"VT_STORAGE"
		TypeStr.Add VT_STREAMED_OBJECT,"VT_STREAMED_OBJECT"
		TypeStr.Add VT_STORED_OBJECT,"VT_STORED_OBJECT"
		TypeStr.Add VT_BLOB_OBJECT,"VT_BLOB_OBJECT"
		TypeStr.Add VT_CF,"VT_CF"
		TypeStr.Add VT_CLSID,"VT_CLSID"
		TypeStr.Add VT_VECTOR,"VT_VECTOR"
		TypeStr.Add VT_ARRAY,"VT_ARRAY"
		TypeStr.Add VT_BYREF,"VT_BYREF"
		TypeStr.Add VT_RESERVED,"VT_RESERVED"

		TKIND_ENUM = 0
		TKIND_RECORD = 1
		TKIND_MODULE = 2
		TKIND_INTERFACE = 3
		TKIND_DISPATCH = 4
		TKIND_COCLASS = 5
		TKIND_ALIAS = 6
		TKIND_UNION = 7
		TKIND_MAX = 8


	End Sub

	Private Function OstDel(Pa1, Pa2)
		OstDel = Pa1 - ((Pa1\Pa2)*Pa2)
	End Function

	Private Sub Class_Terminate	' Setup Terminate event.
		TLITypeLibInfo = Null
		TLIApplication = Null
		DefaultInterfaseName = Null
	End Sub

	Function AttrIsEnum ( InvokeKinds ) ' ������� ���� ������������
		IsEnumAttr = False
		if OstDel(Fix(InvokeKinds/INVOKE_CONST),2) = 1 Then
			IsEnumAttr = True
		End IF
	End Function

	Function AttrIsEvent ( InvokeKinds ) ' ������� ���� �������
		AttrIsEvent = False
		if OstDel(Fix(InvokeKinds/INVOKE_EVENTFUNC),2) = 1 Then
			AttrIsEvent = True
		End IF
	End Function

	Function AttrIsFunc ( InvokeKinds ) ' ������� ���� �����
		AttrIsFunc = False
		if OstDel(Fix(InvokeKinds/INVOKE_FUNC),2) = 1 Then
			AttrIsFunc = True
		End IF
	End Function

	Function AttrIsProps ( InvokeKinds ) ' ������� ���� ��������
		AttrIsProps = False
		if OstDel(Fix(InvokeKinds/INVOKE_PROPERTYGET),2) = 1 Then
			AttrIsProps = True
		Elseif OstDel(Fix(InvokeKinds/INVOKE_PROPERTYPUT),2) = 1 Then
			AttrIsProps = True
		Elseif OstDel(Fix(InvokeKinds/INVOKE_PROPERTYPUTREF),2) = 1 Then
			AttrIsProps = True
		End IF
	End Function


	Function GetObjectByName( InterfaseName )
		GetObjectByName = Null
		if IsObject(Library) Then
			set GetObjectByName = Library.Interfaces.NamedItem(""+ InterfaseName )
		End If
	End Function
	Function GetObjectDefault(  )
		GetObjectDefault = Null
		if IsObject(Library) Then
			set GetObjectDefault = Library.Interfaces.NamedItem(""+ DefaultInterfaseName )
		End If
	End Function

	'������� ���������������������(���� ��������������������,���� ����������,���������������=0,������� = ����)
	private Function NoMyMakeSearchData (TypeInfoNumber) ',127,0,false)
		NoMyMakeSearchData = 127 * 16777216 + (4096 +  TypeInfoNumber - 4096 * Fix(TypeInfoNumber/4096))
	end Function

	' ��������� ��������, ������� � ������ ���������� ��� ������� �����,
	' � ��� �������� ����� ����������� ���� �� ��� ��� ���� ������������ �
	' ��� NoMyMakeSearchData � ���� �� ���������� (((, �� ������ ))) ���� ����������......
	Function GetMemberOfObjectInfo ( Member, Object )
		GetMemberOfObjectInfo = null
		if IsObject( Member ) And IsObject( Object ) Then
			SearchData = NoMyMakeSearchData( Object.TypeInfoNumber )
			set GetMemberOfObjectInfo = Library.GetMemberInfo ( SearchData, Member.InvokeKinds,  Member.MemberId, Member.Name)
		End if
	end Function


	' ��������������� �������� ���� ������� � ��������� �������������
	Function VarTypeToString ( VarType )
		VarTypeToString = ""
		If TypeStr.Exists( VarType ) Then
			VarTypeToString = TypeStr.Item( VarType )
		End If
	end Function

	' ����� �� �������� �������� �� ���������?
	Function ParamHaveDefValue ( Param )
		ParamHaveDefValue = False
		If Param.Default And Param.Optional Then
			ParamHaveDefValue = True
		End If
	end Function

	' ��� ������������ ��������
	Function ParamIsBinding ( Param )
		ParamIsBinding = False
		If Param.Default Or Param.Optional Then
			ParamIsBinding = True
		End If
	end Function

	' ��������� �� ��������� � ���. DumpInterfaseObject
	Function InterfaseIsDump ( Name )
		InterfaseIsDump = AllDumpInterfases.Exists(Trim(LCase(Name)))
	end Function

	' �������� ��� ��������� �������. DumpInterfaseObject
	Function InterfaseIsDumpMark ( Name )
		If Not InterfaseIsDump(Trim(LCase(Name))) Then
			AllDumpInterfases.Add Trim(LCase(Name)), Trim(LCase(Name))
		end if
	end Function

	' ���������� ��������� ������������� ���� ���������
	Function GetParamTypeStr ( Param )
		ParamTipeStr = ""
		Set ParaVarTypeInfo = Param.VarTypeInfo
		ParaVarType  = ParaVarTypeInfo.VarType
		Fixed = Fix ( ParaVarType / 4096)
		FlagVT_ARRAY = OstDel(Fixed, 2)
		FlagVT_VECTOR = OstDel( Fix( Fixed / 2),2)
		TypeVars = Fix(Fixed/4)*4*4096+(ParaVarType - Fixed)
		If TypeVars = 0 Then ' ������������.....
			TKind = Null
			Do ' ����� ��������� ���� ��� ������
				On Error Resume Next
				set InfoOfType = ParaVarTypeInfo.TypeInfo
				set TIType =  ParaVarTypeInfo.TypeInfo

				TIResolved = ParaVarTypeInfo.TypeInfo
				'set TIResolved = ParaVarTypeInfo.TypeInfo
				TKind = TIResolved.TypeKind
				if Err.Number <> 0 Then
					TKind = ParaVarTypeInfo.TypeInfo.TypeKind
					on error goto 0
				End If
			Loop Until True
			While TKind = TKIND_ALIAS
				TKind = TKIND_MAX
				On Error Resume Next
				set TIResolved = TIResolved.ResolvedType
				TKind = TIResolved.TypeKind
				on error goto 0
			Wend

			If TKind = Null Then
				ParamTipeStr = "?"
			else
				IF ParaVarTypeInfo.IsExternalType Then
					ParamTipeStr = ParaVarTypeInfo.TypeLibInfoExternal.Name+"."+TIType.Name
				Else
					ParamTipeStr = TIType.Name
				End If
			End If

		Else
			If ParaVarTypeInfo.VarType <> VT_VARIANT Then
				ParamTipeStr = VarTypeToString(Param.VarTypeInfo.VarType)
			Else
				ParamTipeStr = "VT_VARIANT"
			End IF

		End If
		If IsEmpty( ParamTipeStr ) Or (ParamTipeStr = "")  Then
			if Param.VarTypeInfo.TypeInfoNumber <> -1 Then
				ParamTipeStr = Param.VarTypeInfo.TypeInfo.TypeKindString
			else
				ParamTipeStr = VarTypeToString(Param.VarTypeInfo.VarType)
			End IF
		End IF
		GetParamTypeStr = ParamTipeStr
	end Function

	' ��������������� ��� ���������� ��� ������
	' ��������� ��������� � ��� ���� ��� ������ ����-��, ��������� � ���������� �� ����-���.
	Private Function MakeIName( iName )
		nMakeIName = lcase(iName)
		If LCase(nMakeIName) = LCase(DefaultInterfaseName) Then nMakeIName = ""
		MakeIName = ThisProgIDForIntel + nMakeIName
	end Function

	' ��������� ������ ��� ������� �� Intellisence "TypesStructureExt"
	Sub SaveForTypesStructureExt (Interf1, FuncPropName, Interf2)
		SaveStr = MakeIName(""+Interf1)+ "." + Lcase(FuncPropName)+"," + "VALUE|"+MakeIName(""+Interf2)
		DumpStringToFile SaveStr, FileNameTypesStructureExt
		DumpStringToFile MakeIName(""+Interf2)+","+Interf2, FileNameTypesStructureExtDef
	end Sub

	' ���������� ���������� �� ���������� � ���� ���������...
	' � �� ���� �� �������� �������� ����� ��� ������������...
	' Name - ��� ����������
	' AllStrIterf - ���������� ���� � ����� ����������
	Sub DumpInterfaseObject ( Name , AllStrIterf )
		'Message "%% " + AllStrIterf + "." + Name
		'Message "���������� �� ����������: " + Name
		FindInterfases = "" ' ��������� "�� ����" ����������...
		On Error Resume Next
		Set Object = GetObjectByName( "" + Name )
		set TrueMembers = Object.Members.GetFilteredMembers
		IF err.number<>0 Then
			On Error goto 0
			Exit Sub
		End If
		IntsFileName = DataDir+"\" + MakeIName(Name)+".ints"
		if fso.FileExists(IntsFileName) Then fso.DeleteFile IntsFileName,true
		Set IntsFile = fso.CreateTextFile(IntsFileName)
		Status "����������: " + IntsFileName

		Set ALSIntefaseFolder = mALSParentF.AddSubFolder( NAme,	"" )
		Set ALSFolderProp = ALSIntefaseFolder.AddSubFolder( "��������",	"" )
		Set ALSFolderFunc = ALSIntefaseFolder.AddSubFolder( "������",		"" )

		for i=1 To TrueMembers.Count
			set Member = TrueMembers.item(i)
			Set MemberInfo = GetMemberOfObjectInfo ( Member, Object )

			MemberIK = Member.InvokeKinds
			If AttrIsProps( MemberIK) Then 'INVOKE_EVENTFUNC  �������
				Set ALSProp = ALSFolderProp.AddSubItem( Member.Name,"" )
				ALSProp.mWriting = Member.Name ': ALSProp.mWritingEng = Member.Name
				ALSProp.mSyntax = Member.Name
				ALSProp.mAssignment = MemberInfo.HelpString

				if MemberInfo.ReturnType.TypeInfoNumber <> -1 Then
					If MemberInfo.ReturnType.TypeInfo.TypeKindString = "coclass" Then
						' ���� ��� ������, ����� ���� ���������
						'Message " prop - " + Member.Name + " dispinterface " + MemberInfo.ReturnType.TypeInfo.DefaultInterface.Name
						FindInterfases = FindInterfases + vbCrLf + MemberInfo.ReturnType.TypeInfo.DefaultInterface.Name
						SaveForTypesStructureExt Name, Member.Name, MemberInfo.ReturnType.TypeInfo.DefaultInterface.Name
						ALSProp.mParams = "����������: """ + MemberInfo.ReturnType.TypeInfo.DefaultInterface.Name+""""
					ElseIf MemberInfo.ReturnType.TypeInfo.TypeKindString = "dispinterface" Then
						'Message " prop - " + Member.Name + " " + MemberInfo.ReturnType.TypeInfo.TypeKindString + " " + MemberInfo.ReturnType.TypeInfo.Name
						FindInterfases = FindInterfases + vbCrLf + MemberInfo.ReturnType.TypeInfo.Name
						SaveForTypesStructureExt Name, Member.Name, MemberInfo.ReturnType.TypeInfo.Name
						ALSProp.mParams = "����������: """ + MemberInfo.ReturnType.TypeInfo.Name + """"
					Else
						'Message " prop - " + Member.Name
					End If
				Else

					ALSProp.mParams = "����������: " + VarTypeToString ( MemberInfo.ReturnType.VarType )
					'Message " prop - " + Member.Name
				End If

				IntsFile.WriteLine "0000 "+ Member.Name
			'ElseIf AttrIsFunc ( MemberIK) And False Then 'INVOKE_FUNC  �������/�����
			ElseIf AttrIsFunc ( MemberIK) Then 'INVOKE_FUNC  �������/�����
				' ��������� ������ ������....
				Set ALSProp = ALSFolderFunc.AddSubItem( Member.Name,"" )
				ALSProp.mAssignment = MemberInfo.HelpString

				StrForMessage = ""
				StrForAls = Member.Name+"("
				StrForAlsParams = ""
				if MemberInfo.ReturnType.TypeInfoNumber = -1 Then
					StrForMessage = " meth - " + Member.Name
				Else
					StrForMessage = " func - " + Member.Name
				end if
				ForFileString = Member.Name + "("
				For ParCnt = 1 To MemberInfo.Parameters.Count
					Set Param = MemberInfo.Parameters.Item(ParCnt)
					' Param.Default = true -> ����� �������� �� ��������� = ��������������
					'if flOutNameAndTypeParameters Then
						if Param.Default Then
							StrForMessage = StrForMessage + "["
							StrForAls = StrForAls + "["
						End if
						StrForMessage = StrForMessage + Param.Name + " as " + GetParamTypeStr ( Param )
						StrForAls = StrForAls + Param.Name
						StrForAlsParams = StrForAlsParams + "<"+ Param.Name +"> - " + GetParamTypeStr ( Param )
						if Param.Default Then
							StrForMessage = StrForMessage + "]"
							StrForAls = StrForAls + "]"
						end if
					'end if
					if ParCnt <> MemberInfo.Parameters.Count Then ForFileString = ForFileString + ", " : StrForAlsParams = StrForAlsParams + vbCrLf
					if ParCnt <> MemberInfo.Parameters.Count Then StrForMessage = StrForMessage + ", "
					if ParCnt <> MemberInfo.Parameters.Count Then StrForAls = StrForAls + ", "
					StrForMessage = Replace(StrForMessage," ","") ' ������� �������
					ForFileString = Replace(ForFileString," ","") ' ������� �������
				Next
				ForFileString = ForFileString + ")"
				StrForMessage = StrForMessage + ")"
				StrForAls = StrForAls + ")"
				ALSProp.mWriting = ForFileString
				ALSProp.mSyntax = StrForAls
				ALSProp.mParams = StrForAlsParams
				' ���� �������� ��� ������������� ��������
				if MemberInfo.ReturnType.TypeInfoNumber <> -1 Then
					if MemberInfo.ReturnType.TypeInfo.TypeKindString = "dispinterface" Then
						FindInterfases = FindInterfases + vbCrLf + MemberInfo.ReturnType.TypeInfo.Name
						StrForMessage = StrForMessage + " as Interfase " + MemberInfo.ReturnType.TypeInfo.Name
						SaveForTypesStructureExt Name, Member.Name, MemberInfo.ReturnType.TypeInfo.Name
						ALSProp.mParams = ALSProp.mParams + vbCrLf + "����������: """ + MemberInfo.ReturnType.TypeInfo.Name+""""
					Elseif MemberInfo.ReturnType.TypeInfo.TypeKindString = "coclass" Then
						FindInterfases = FindInterfases + vbCrLf + MemberInfo.ReturnType.TypeInfo.DefaultInterface.Name
						SaveForTypesStructureExt Name, Member.Name, MemberInfo.ReturnType.TypeInfo.DefaultInterface.Name
						FindInterfases = FindInterfases + vbCrLf + MemberInfo.ReturnType.TypeInfo.DefaultEventInterface.Name
						SaveForTypesStructureExt Name, Member.Name, MemberInfo.ReturnType.TypeInfo.DefaultEventInterface.Name
						ALSProp.mParams = ALSProp.mParams + vbCrLf + "����������: """ + MemberInfo.ReturnType.TypeInfo.DefaultInterface.Name+""""
					end if
				end if
				'if MemberInfo.HelpString <> "" Then StrForMessage = StrForMessage + " Help -> " + MemberInfo.HelpString
				'Message StrForMessage
				IntsFile.WriteLine "0000 " + ForFileString
			End IF
		Next
		IntsFile.Close
		InterfaseIsDumpMark(Name)

		ArrNameOfIntrf = Split(FindInterfases, vbcrlf)
		for i=0 To UBound(ArrNameOfIntrf)
			iName = trim(ArrNameOfIntrf(i))
			if iName <> "" Then
				if Not InterfaseIsDump(iName) Then
					DumpInterfaseObject iName, AllStrIterf + "." + Name
				End IF
			End IF
		next
	end Sub
End class 'ITypeLib



' �������� ��������� ���������
Function IntsGenerator( ProgID )
'Stop
	IntsGenerator = False
	if ProgIDIsDumped( ProgID ) Then Exit Function
	if not IsProgID( ProgID )  Then Exit Function

	Set li = new ITypeLib
	li.DebugFlag = true
	li.flOutNameAndTypeParameters = false
	if li.LoadLibrary( ProgId , True) Then
		Set li.mALSGenerator = New ALSGenerator
		li.mALSGenerator.Init DataDirAls+ ProgID +".als"
		Set li.mALSParentF = li.mALSGenerator.CreateRootFolder( ProgID, "")

		IntsGenerator = True
		messageStr =				"######################################################################" + vbCrLf
		messageStr = messageStr +	"����������: " + li.LibPath + " ������� �� ����-���: " + li.ThisProgID + vbCrLf
		messageStr = messageStr +	"��������� �� ���������:" + li.DefaultInterfaseName + vbCrLf
		AllStrIterf = ""
		li.DumpInterfaseObject  li.DefaultInterfaseName , AllStrIterf
		messageStr = messageStr +	"������������� ������ ��: " + ProgID
		message messageStr, mExclamation3
		Set ALSProp			= li.mALSParentF.AddSubItem( ProgID,"" )

		ALSProp.mWriting	= ProgID ': ALSProp.mWritingEng = Member.Name
		ALSProp.mSyntax		= "������ = �������������(""" + ProgID+ """);" + vbCrLf + "������.<?>;"
		ALSProp.mWriting	= "������ = �������������(""" + ProgID+ """);" + vbCrLf + "������.<?>;"
		ALSProp.mNotes		= "@������_����������:"	+ vbCrLf +_
								"Help: "		+ li.Library.HelpString + vbCrLf +_
								"Name: "		+ li.Library.Name + vbCrLf +_
								"GUID: "		+ li.Library.GUID + vbCrLf +_
								"LibPath: """	+ li.LibPath +""""+ vbCrLf +_
								"===================================================="	 + vbCrLf +_
								"als-file-path:"	+ li.mALSGenerator.mFileName + vbCrLf +_
								"���� ������������ c �������: intsOLEGenerator.vbs" + vbCrLf +_
								"openconf �����!" + vbCrLf +_
								"�����: ������ �. ��� trdm, 2005 �. trdm�fromru.com"	 + vbCrLf +_
								"===================================================="

		li.mALSGenerator.Dump
		message "######################################################################",mExclamation3
		message "�������-��������: ������� ������������ ����, ������� ����� ""��������""", mExclamation3
	ElseIF li.ErrorNumber<>0 Then
		message "������ ��������� ������ �� ����-���: " + li.ThisProgID, mExclamation
	End IF
End Function

' ���������� *.ints-����� �� ������������� ����-���
Sub Generator()
	ProgID = ""
	ProgID = InputBox ("������� ProgID", "OLE-ActiveX *.ints Generator",ProgID)
	ProgID = Trim(ProgID)
	If Len(ProgID)>0 Then
		if IntsGenerator( ProgID ) Then
			on Error resume next
			Scripts("Intellisence").ReloadDictionary
			if err.number<>0 Then
				message "� Intellisence �� ������� ��������� ReloadDictionary!!!", mExclamation3
				message "��������� �������������� �������� ����������!", mExclamation3
			End If
			on Error goto 0
		End If
	End IF
End Sub

' � ������ ������������� �������� ����� ���������� ���������� ������,
' ���������������� ����-���, �� ���� ���������������� �� ����������,
' � �� ������� ������� ��� ���� ����������� ��������� ������� ���������
' *.ints-������, ������ ����������� � ����� ProgIDDumped.txt � ��� �������
' ��������� �� ���� ����������� � ��������.
'
' ������ ��������� ��������� ���������� ����������� ������ �� ������� �������.
' ��� ����������� ��� ���� ������� �� ������������� ������.
Sub ReGenerator()
	If FSO.FileExists(FileNameProgIDDumped) Then
		Set File0 = fso.GetFile( FileNameProgIDDumped )
		Set File = File0.OpenAsTextStream(1)
		if File0.Size > 0 Then
			TextStream = File.ReadAll()
			Set ch = CreateObject("svcsvc.service")
			needProgID = ch.SelectValue(TextStream,"�������� ������ �����������")
			File.Close
			if Len(needProgID)>0 Then
				fso.DeleteFile  FileNameProgIDDumped, true
				Set File = fso.CreateTextFile(FileNameProgIDDumped)

				AllProgID = Split(TextStream, vbCrLf)
				for i=0 To UBound (AllProgID)
					if needProgID<>AllProgID(i) And Len(Trim(AllProgID(i)))>0 Then
						File.WriteLine AllProgID(i)
					End If
				Next
				File.Close
				if IntsGenerator( needProgID ) Then
					on Error resume next
					Scripts("Intellisence").ReloadDictionary
					if err.number<>0 Then
						message "� Intellisence �� ������� ��������� ReloadDictionary!!!", mExclamation3
						message "��������� �������������� �������� ����������!", mExclamation3
					End If
					on Error goto 0
				End If

			End If
		Else
			File.Close
			Message "����: " + FileNameProgIDDumped + " ����!", mExclamation
			Exit Sub
		End IF
	Else
		Message "����: " + FileNameProgIDDumped + " �� ����������!", mExclamation
	End IF
End Sub


' ��� ���� ��������� ���:
'1. ints-����� � ������ ����������
'2. ��� TypesStructureExt �� Intellisence ���� ��� ������� ��������� ������
'	TypesStructureExt.Add "excel_application.range", "VALUE|excel_applicationrange"
'	TypesStructureExt.Add "excel_applicationrange.cells", "VALUE|excel_applicationcells"
'3 ��� OtherTypesDefine �� �� Intellisence ���� ��� ������� ��������� ������
'	OtherTypesDefine.Add "word.application", "word_application"
'4 �������� ����-���� �� ������� ��� ������������� ������ ��� ��� �� ������� ������ �����������
' ����� ����������, ��-�� ����, ��� ������, ��������� �� ����-��� ����� ���� ������.

Sub CommonGenerator( )
	IntsGenerator( "Scripting.FileSystemObject" )
	IntsGenerator( "ADODB.Connection" )
	IntsGenerator( "WScript.Shell" )
	IntsGenerator( "MSXML2.DOMDocument" )

Exit sub ' ��������� ���� �������.......
	IntsGenerator( "Excel.Application" )
	IntsGenerator( "Word.application" )
	IntsGenerator( "SYSINFO.SysInfo" )
	IntsGenerator( "Svcsvc.Service" )
	IntsGenerator( "v77.Application" )
	IntsGenerator( "VBScript.RegExp" ) ' ���� ��� ����������� ���������� (((( �������� �������....
	IntsGenerator( "SHPCE.Profiler" )
	IntsGenerator( "ODBCSQL.RarusSQL" )
	IntsGenerator( "TLI.TypeLibInfo" )
	IntsGenerator( "TLI.TLIApplication" )
	IntsGenerator( "Rarus.ApiExtender" )
End Sub

Sub TestLibrary()
	On Error Resume Next
	Set TLITypeLibInfo = CreateObject("TLI.TypeLibInfo")
	Set TLIApplication = CreateObject("TLI.TLIApplication")
	if Err.number<>0 then
		Message "����������: TLBINF32.DLL �� ���������������� �� ������ ������� �����!"
	Else
		Message "����������: TLBINF32.DLL ����������������! ��� � �������!"
	End If
	On Error Goto 0
End Sub
' �������������....
Sub InitScript( para )
	set WSH = CreateObject("WScript.Shell")
	Set FSO = CreateObject("Scripting.FileSystemObject")
End Sub

' �������� ���������, ������������ ����������� ���������.....
Private Sub TestAlsCreation()
	'Stop
	Set AG = New ALSGenerator
	AG.Init DataDirAls+"aaa.als"
	Set PaFo = AG.CreateRootFolder("����� 1-�� ������", "�������")
	If IsObject(PaFo) Then
		for i=0 To 10
			Set F2L= PaFo.AddSubFolder("����� 2-�� ������","--")
			for ii = 0 To 10
				Set F3L = F2L.AddSubFolder ( "����� 3-�� ������","---")
				Set Si3L = F2L.AddSubItem ( "������� 3-�� ������","---" )
				Si3L.mWriting = "������������������� �� �����"
				Si3L.mWritingEng = "KHDSs sdvsdm wefke"
				For iii = 0 To 10
					F3L.AddSubItem  "������� 4-�� ������","----"
				Next
			Next
		Next
	End If
	AG.Dump
End Sub


InitScript 0