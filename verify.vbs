Option Explicit

' �������� ���������� ��������� ������
If WScript.Arguments.Count = 0 Then
    WScript.Echo "�� �������� ��� �����."
    WScript.Quit 1
End If

Dim fileToVerify
fileToVerify = WScript.Arguments(0) ' ��� ����� ���������� ��� ������ ��������

Const CAPICOM_CURRENT_USER_STORE = 2
Const CAPICOM_CERTIFICATE_INCLUDE_WHOLE_CHAIN = 1
Const CADESCOM_CADES_BES = 1
Const CADESCOM_BASE64_TO_BINARY = 1
Const CADESCOM_STRING_TO_UCS2LE = 0

Dim oSignedData
Dim oSettings
Set oSignedData = CreateObject("CAdESCOM.CadesSignedData")

WScript.Echo "�������� ������..."

Dim fileContent : fileContent = LoadFileToBase64(fileToVerify)
Dim fileSig : fileSig = LoadFile(fileToVerify & ".sig")
WScript.Echo "�������� ���������."

WScript.Echo "�������� �������..."

oSignedData.ContentEncoding = CADESCOM_BASE64_TO_BINARY
oSignedData.Content = fileContent

' �������� ���������� �������
oSignedData.VerifyCades fileSig, CADESCOM_CADES_BES, True

oSignedData.Display

' �������� ���������� � �������
Dim signers
Set signers = oSignedData.Signers
Dim signer 
Set signer = signers.Item(1) ' ����� ������ �������
Dim cert
Set cert = signer.Certificate

' ������� ���������� � ��������� �����������
WScript.Echo "�������� �����������: " & cert.SubjectName
WScript.Echo "�������� �����: " & cert.SerialNumber


WScript.Echo "�������� ������� ���������."

WScript.Echo fileToVerify & " - ������� �������������."

Function GetSignerCertificate(SerialNumber)
Set GetSignerCertificate = Nothing
Dim oCert
Dim oStore
Set oStore = CreateObject("CAdESCOM.Store")
oStore.Open CAPICOM_CURRENT_USER_STORE
For Each oCert In oStore.Certificates
	If oCert.SerialNumber = SerialNumber Then
	Set GetSignerCertificate = oCert
	Exit For
	End If
Next
End Function

Function LoadFile (FileName)
	Const ForReading = 1
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If Not fso.FileExists(FileName) Then
		Err.Raise vbObjectError+1, "LoadFile", FileName & " - File not found"
	End If 
	Dim ts
	Set ts = fso.OpenTextFile(FileName, ForReading)  
	LoadFile = ts.ReadAll
End Function

' anthropic/claude-3.5-sonnet
Function LoadFileToBase64(filePath)
    ' Create ADODB.Stream for binary reading
    Set binStream = CreateObject("ADODB.Stream")
	Dim binStream
    binStream.Type = 1 'Binary
    binStream.Open
    binStream.LoadFromFile filePath
    
    ' Get binary content
	Dim bytes
    bytes = binStream.Read
    binStream.Close
    Set binStream = Nothing
    
    ' Create XML DOM for base64 encoding
	Dim xml
    Set xml = CreateObject("MSXML2.DOMDocument")
	Dim element
    Set element = xml.createElement("base64")
    element.dataType = "bin.base64"
    element.nodeTypedValue = bytes
    
    ' Get base64 string
    LoadFileToBase64 = element.text
    
    Set element = Nothing
    Set xml = Nothing
End Function