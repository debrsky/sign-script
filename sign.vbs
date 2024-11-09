Option Explicit

const SERIAL_NUMBER = "7C00176E4CF1A304C0CF542651000A00176E4C"

Const CAPICOM_CURRENT_USER_STORE = 2
Const CAPICOM_CERTIFICATE_INCLUDE_WHOLE_CHAIN = 1
Const CADESCOM_CADES_BES = 1
Const CADESCOM_BASE64_TO_BINARY = 1
Const CADESCOM_STRING_TO_UCS2LE = 0

WScript.Echo "Инициализация..."

Dim oSigner : Set oSigner = CreateObject("CAdESCOM.CPSigner")

On Error Resume Next
oSigner.Certificate = GetSignerCertificate(SERIAL_NUMBER)
If Err.Number <> 0 Then
    WScript.Echo "ОШИБКА при получении сертификата: " & Err.Description & " (0x" & Hex(Err.Number) & ")"
    WScript.Quit 1
End If
On Error Goto 0

Dim args
Set args = WScript.Arguments

If args.Count = 0 Then
    WScript.Echo "Перетащите файлы на пакетный файл для подписи."
    WScript.Quit
End If

WScript.Echo "Найдено файлов для подписи: " & args.Count
WScript.Echo "-----------------------------------"

Dim i
For i = 0 To args.Count - 1
    Dim filePath : filePath = args(i)
    WScript.Echo "Обработка файла (" & (i + 1) & "/" & args.Count & "): " & filePath
    
    On Error Resume Next
    
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        WScript.Echo "  ОШИБКА: Файл не существует"
    Else
        Dim fileObj : Set fileObj = fso.GetFile(filePath)
        If fileObj.Size = 0 Then
            WScript.Echo "  ОШИБКА: Файл пустой. Пропускаем..."
        Else
            WScript.Echo "  Размер: " & FormatFileSize(fileObj.Size)
            Dim signedData
            signedData = SignFile(filePath, oSigner)
            
            If Err.Number <> 0 Then
                WScript.Echo "  ОШИБКА: " & Err.Description & " (0x" & Hex(Err.Number) & ")"
                Err.Clear
            Else
                Dim file
                Set file = fso.CreateTextFile(filePath & ".sig", True)
                file.WriteLine "-----BEGIN CMS-----"
                file.WriteLine signedData
                file.WriteLine "-----END CMS-----"
                file.Close
                WScript.Echo "  Создана подпись: " & filePath & ".sig"
            End If
        End If
        Set fileObj = Nothing
    End If
    
    Set fso = Nothing
    WScript.Echo "-----------------------------------"
Next

Function FormatFileSize(bytes)
    If bytes < 1024 Then 
        FormatFileSize = bytes & " Б"
    ElseIf bytes < 1024*1024 Then
        FormatFileSize = Round(bytes/1024, 1) & " КБ"
    Else
        FormatFileSize = Round(bytes/(1024*1024), 1) & " МБ"
    End If
End Function

WScript.Echo "Обработка завершена."

Function SignFile(filePath, oSigner)
    On Error Resume Next
    
    Dim oSignedData
    Set oSignedData = CreateObject("CAdESCOM.CadesSignedData")
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "SignFile", "Ошибка создания CadesSignedData: " & Err.Description
        Exit Function
    End If
    
    Dim content
    content = LoadFileToBase64(filePath)
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "SignFile", "Ошибка чтения файла: " & Err.Description
        Exit Function
    End If
    
    If Len(content) = 0 Then
        Err.Raise vbObjectError + 1, "SignFile", "Пустой файл не может быть подписан"
        Exit Function
    End If
    
    oSignedData.ContentEncoding = CADESCOM_BASE64_TO_BINARY
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "SignFile", "Ошибка установки ContentEncoding: " & Err.Description
        Exit Function
    End If
    
    oSignedData.Content = content
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "SignFile", "Ошибка установки контента: " & Err.Description & ". Размер контента: " & Len(content)
        Exit Function
    End If
    
    SignFile = oSignedData.SignCades(oSigner, CADESCOM_CADES_BES, true)
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "SignFile", "Ошибка подписи: " & Err.Description
        Exit Function
    End If
    
    On Error Goto 0
End Function

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

Function LoadFileToBase64(filePath)
    On Error Resume Next
    
    ' Create ADODB.Stream for binary reading
    Dim binStream
    Set binStream = CreateObject("ADODB.Stream")
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "LoadFileToBase64", "Ошибка создания ADODB.Stream: " & Err.Description
        Exit Function
    End If
    
    binStream.Type = 1 'Binary
    binStream.Open
    
    binStream.LoadFromFile filePath
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "LoadFileToBase64", "Ошибка загрузки файла в поток: " & Err.Description
        Exit Function
    End If
    
    Dim bytes
    bytes = binStream.Read
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "LoadFileToBase64", "Ошибка чтения из потока: " & Err.Description
        Exit Function
    End If
    
    binStream.Close
    Set binStream = Nothing
    
    Dim xml
    Set xml = CreateObject("MSXML2.DOMDocument")
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "LoadFileToBase64", "Ошибка создания MSXML2.DOMDocument: " & Err.Description
        Exit Function
    End If
    
    Dim element
    Set element = xml.createElement("base64")
    element.dataType = "bin.base64"
    
    element.nodeTypedValue = bytes
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "LoadFileToBase64", "Ошибка конвертации в base64: " & Err.Description
        Exit Function
    End If
    
    LoadFileToBase64 = element.text
    
    Set element = Nothing
    Set xml = Nothing
    
    On Error Goto 0
End Function
