Option Explicit

' Проверка наличия аргументов командной строки
If WScript.Arguments.Count = 0 Then
    WScript.Echo "Необходимо указать файл для проверки."
    WScript.Quit 1
End If

Dim fileToVerify
fileToVerify = WScript.Arguments(0) ' Путь к файлу, который будет проверяться

Const CAPICOM_CURRENT_USER_STORE = 2
Const CAPICOM_CERTIFICATE_INCLUDE_WHOLE_CHAIN = 1
Const CADESCOM_CADES_BES = 1
Const CADESCOM_BASE64_TO_BINARY = 1
Const CADESCOM_STRING_TO_UCS2LE = 0

Dim oSignedData
Dim oSettings
Set oSignedData = CreateObject("CAdESCOM.CadesSignedData")

WScript.Echo "Начинаем проверку подписи..."

Dim fileContent : fileContent = LoadFileToBase64(fileToVerify)
Dim fileSig : fileSig = LoadFile(fileToVerify & ".sig")
WScript.Echo "Файлы загружены."

WScript.Echo "Проверяем подпись..."

oSignedData.ContentEncoding = CADESCOM_BASE64_TO_BINARY
oSignedData.Content = fileContent

' Проверка подписи
oSignedData.VerifyCades fileSig, CADESCOM_CADES_BES, True

oSignedData.Display

' Получение информации о подписанте
Dim signers
Set signers = oSignedData.Signers
Dim signer 
Set signer = signers.Item(1) ' Получение первого подписанта
Dim cert
Set cert = signer.Certificate

' Вывод информации о сертификате
WScript.Echo "Подписывающий сертификат: " & cert.SubjectName
WScript.Echo "Серийный номер сертификата: " & cert.SerialNumber

WScript.Echo "Проверка завершена."

WScript.Echo fileToVerify & " - подпись проверена успешно."

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
        Err.Raise vbObjectError + 1, "LoadFile", FileName & " - Файл не найден"
    End If 
    Dim ts
    Set ts = fso.OpenTextFile(FileName, ForReading)  
    LoadFile = ts.ReadAll
End Function

' Функция для загрузки файла и кодирования его в Base64
Function LoadFileToBase64(filePath)
    ' Создаем ADODB.Stream для бинарного чтения
    Dim binStream
    Set binStream = CreateObject("ADODB.Stream")
    binStream.Type = 1 ' Binary
    binStream.Open
    binStream.LoadFromFile filePath
    
    ' Получаем бинарное содержимое
    Dim bytes
    bytes = binStream.Read
    binStream.Close
    Set binStream = Nothing
    
    ' Создаем XML DOM для кодирования Base64
    Dim xml
    Set xml = CreateObject("MSXML2.DOMDocument")
    Dim element
    Set element = xml.createElement("base64")
    element.dataType = "bin.base64"
    element.nodeTypedValue = bytes
    
    ' Получаем строку Base64
    LoadFileToBase64 = element.text
    
    Set element = Nothing
    Set xml = Nothing
End Function
