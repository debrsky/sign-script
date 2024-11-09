Option Explicit
Const CAPICOM_CURRENT_USER_STORE = 2
Const CAPICOM_CERTIFICATE_INCLUDE_WHOLE_CHAIN = 1
Const CADESCOM_CADES_BES = 1
Const CADESCOM_BASE64_TO_BINARY = 1

Dim oSigner
Set oSigner = CreateObject("CAdESCOM.CPSigner")
' Укажите правильный серийный номер сертификата.
Dim sSerialNumber : sSerialNumber = "7C00176E4CF1A304C0CF542651000A00176E4C"
' Укажите правильный адрес службы штампов времени.
Dim sTSAAddress : sTSAAddress = "http://domain/tsp/tsp.srf"

oSigner.Certificate = GetSignerCertificate(sSerialNumber)
oSigner.TSAAddress = sTSAAddress

Dim oSignedData
Dim oSettings
Set oSignedData = CreateObject("CAdESCOM.CadesSignedData")

oSignedData.Content = LoadFileToBase64("file.bin")

Dim sSignedData

' sSignedData = oSignedData.SignCades(oSigner, CADESCOM_CADES_BES, false) ' присоединная
sSignedData = oSignedData.SignCades(oSigner, CADESCOM_CADES_BES, true) ' отсоединенная

' Сохранение signed data в файл с расширением .sig
Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile("file.bin.sig", True) ' "True" для перезаписи, если файл существует
file.Write sSignedData
file.Close

WScript.Echo 'Готово.';

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

' written by anthropic/claude-3.5-sonnet
Function LoadFileToBase64(filePath)
    Dim fso, file, stream, contents, base64Chars
    Dim result, buffer, binStr, sixBits
    Dim i, j, ascVal, decVal, padLen
    
    ' Read binary file using FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.GetFile(filePath)
    Set stream = file.OpenAsTextStream(1, -2) ' ForReading, TristateTrue for binary
    
    ' Read the file contents
    contents = stream.ReadAll()
    stream.Close
    
    ' Initialize base64 chars array
    base64Chars = Array("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z", _
                        "a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z", _
                        "0","1","2","3","4","5","6","7","8","9","+","/")
    
    ' Convert to base64
    result = ""
    buffer = ""
    
    For i = 1 To Len(contents)
        ' Convert char to binary
        ascVal = Asc(Mid(contents, i, 1))
        binStr = ""
        
        For j = 7 To 0 Step -1
            If (ascVal And 2^j) Then
                binStr = binStr & "1"
            Else
                binStr = binStr & "0"
            End If
        Next
        
        buffer = buffer & binStr
        
        ' Process every 6 bits
        Do While Len(buffer) >= 6
            sixBits = Left(buffer, 6)
            buffer = Mid(buffer, 7)
            
            ' Convert 6 bits to decimal
            decVal = 0
            For j = 0 To 5
                If Mid(sixBits, j + 1, 1) = "1" Then
                    decVal = decVal + 2^(5-j)
                End If
            Next
            
            ' Add base64 char
            result = result & base64Chars(decVal)
        Loop
    Next
    
    ' Handle remaining bits
    If Len(buffer) > 0 Then
        ' Pad with zeros
        While Len(buffer) < 6
            buffer = buffer & "0"
        Wend
        
        decVal = 0
        For j = 0 To 5
            If Mid(buffer, j + 1, 1) = "1" Then
                decVal = decVal + 2^(5-j)
            End If
        Next
        
        result = result & base64Chars(decVal)
        
        ' Add padding if needed
        padLen = (3 - ((Len(contents) Mod 3))) Mod 3
        For i = 1 To padLen
            result = result & "="
        Next
    End If
    
    LoadFileToBase64 = result
End Function
