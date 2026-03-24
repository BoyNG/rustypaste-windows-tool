Option Explicit

Dim shell, fso
Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")

' === SETTINGS ===
Const AUTH_TOKEN = "auth_token_from_rustypaste_config"
Const DELETE_TOKEN = "del_token_from_rustypaste_config"
Const SERVER_URL = "https://rustypaste.home.serv/"
Const ADD_DATE_SUFFIX = True
Const DEFAULT_EXPIRE = "60d"
Const LANG = "en"  ' "ru" or "en"
' =================

' Text strings on any language (en and ru for axample)
Dim TXT_MODE, TXT_FILE, TXT_PATH, TXT_IMAGE, TXT_TEXT, TXT_URL, TXT_EMPTY, TXT_ERROR, TXT_INFO, TXT_NAME, TXT_SIZE, TXT_TYPE, TXT_MODIFIED, TXT_EXPIRE_LABEL, TXT_EXPIRE_HINT, TXT_EXPIRE_DEFAULT, TXT_INPUT_TITLE, TXT_URL_CHOICE_TITLE, TXT_URL_CHOICE_MSG, TXT_URL_ACTION_SHORTEN, TXT_URL_ACTION_DOWNLOAD, TXT_COPIED, TXT_FILE_INFO, TXT_CLIPBOARD_FILE, TXT_CLIPBOARD_IMAGE, TXT_CLIPBOARD_TEXT, TXT_CLIPBOARD_PATH, TXT_DRAG_FILE, TXT_CLIPBOARD_EMPTY

If LANG = "en" Then
    TXT_MODE = "MODE:"
    TXT_FILE = "FILE"
    TXT_PATH = "PATH"
    TXT_IMAGE = "IMAGE"
    TXT_TEXT = "TEXT"
    TXT_URL = "URL"
    TXT_EMPTY = "Clipboard is empty."
    TXT_ERROR = "Failed to get link."
    TXT_INFO = "=== FILE INFO ==="
    TXT_NAME = "Name"
    TXT_SIZE = "Size"
    TXT_TYPE = "Type"
    TXT_MODIFIED = "Modified"
    TXT_EXPIRE_LABEL = "File lifetime (example: 60d, 24h, 120min, 30):"
    TXT_EXPIRE_HINT = "Default"
    TXT_INPUT_TITLE = "File Upload"
    TXT_URL_CHOICE_TITLE = "URL Mode"
    TXT_URL_CHOICE_MSG = "Select action:"
    TXT_URL_ACTION_SHORTEN = "YES - Shorten URL"
    TXT_URL_ACTION_DOWNLOAD = "NO - Download file from URL"
    TXT_COPIED = "Link copied to clipboard."
    TXT_FILE_INFO = "File from clipboard (Ctrl+C in Explorer)"
    TXT_CLIPBOARD_FILE = "File from clipboard"
    TXT_CLIPBOARD_IMAGE = "Image from clipboard"
    TXT_CLIPBOARD_TEXT = "Text from clipboard"
    TXT_CLIPBOARD_PATH = "File path from clipboard"
    TXT_DRAG_FILE = "Dragged file"
    TXT_CLIPBOARD_EMPTY = "Clipboard is empty."
Else
    TXT_MODE = "РЕЖИМ:"
    TXT_FILE = "ФАЙЛ"
    TXT_PATH = "ПУТЬ"
    TXT_IMAGE = "ИЗОБРАЖЕНИЕ"
    TXT_TEXT = "ТЕКСТ"
    TXT_URL = "URL"
    TXT_EMPTY = "В буфере обмена пусто."
    TXT_ERROR = "Не удалось получить ссылку."
    TXT_INFO = "=== ИНФОРМАЦИЯ О ФАЙЛЕ ==="
    TXT_NAME = "Имя"
    TXT_SIZE = "Размер"
    TXT_TYPE = "Тип"
    TXT_MODIFIED = "Изменён"
    TXT_EXPIRE_LABEL = "Время жизни файла (пример: 60d, 24h, 120min, 30):"
    TXT_EXPIRE_HINT = "По умолчанию"
    TXT_INPUT_TITLE = "Загрузка файла"
    TXT_URL_CHOICE_TITLE = "Режим работы с URL"
    TXT_URL_CHOICE_MSG = "Выберите действие:"
    TXT_URL_ACTION_SHORTEN = "ДА - Укоротить ссылку"
    TXT_URL_ACTION_DOWNLOAD = "НЕТ - Загрузить файл из ссылки"
    TXT_COPIED = "Ссылка скопирована в буфер обмена."
    TXT_FILE_INFO = "Файл из буфера обмена (Ctrl+C в проводнике)"
    TXT_CLIPBOARD_FILE = "Файл из буфера"
    TXT_CLIPBOARD_IMAGE = "Изображение из буфера"
    TXT_CLIPBOARD_TEXT = "Текст из буфера"
    TXT_CLIPBOARD_PATH = "Путь файла из буфера"
    TXT_DRAG_FILE = "Перетянутый файл"
    TXT_CLIPBOARD_EMPTY = "В буфере обмена пусто."
End If

Dim output
output = ""
Dim expireTime
expireTime = DEFAULT_EXPIRE

Function NormalizeExpire(exp)
    exp = Trim(exp)
    If exp = "" Then
        NormalizeExpire = DEFAULT_EXPIRE
        Exit Function
    End If
    If IsNumeric(exp) Then
        NormalizeExpire = exp & "min"
    Else
        NormalizeExpire = exp
    End If
End Function

Function AskExpireTime(info)
    Dim msg, userInput
    msg = info & vbCrLf & _
          TXT_EXPIRE_LABEL & vbCrLf & _
          TXT_EXPIRE_HINT & ": " & DEFAULT_EXPIRE
    userInput = InputBox(msg, TXT_INPUT_TITLE, DEFAULT_EXPIRE)
    AskExpireTime = NormalizeExpire(userInput)
End Function

' === 1. Check comand line arguments ===
If WScript.Arguments.Count > 0 Then
    Dim filePath
    filePath = WScript.Arguments(0)
    If fso.FileExists(filePath) Then
        output = output & TXT_MODE & " " & TXT_DRAG_FILE & vbCrLf
        output = output & TXT_FILE & ": " & filePath & vbCrLf & vbCrLf
        output = output & GetFileInfo(filePath)

        expireTime = AskExpireTime(output)

        UploadFile filePath
        WScript.Quit
    End If
End If

' === 2. Chech file from exchange bufer (CF_HDROP) ===
Dim fileFromClipboard
fileFromClipboard = GetFileFromClipboard()

If fileFromClipboard <> "" And fso.FileExists(fileFromClipboard) Then
    output = output & TXT_MODE & " " & TXT_FILE_INFO & vbCrLf
    output = output & TXT_FILE & ": " & fileFromClipboard & vbCrLf & vbCrLf
    output = output & GetFileInfo(fileFromClipboard)

    expireTime = AskExpireTime(output)

    UploadFile fileFromClipboard
    WScript.Quit
End If

' === 3. Chech picture from exchange bufer ===
Dim imageFromClipboard
imageFromClipboard = GetImageFromClipboard()

If imageFromClipboard <> "" Then
    output = output & TXT_MODE & " " & TXT_CLIPBOARD_IMAGE & vbCrLf
    output = output & TXT_FILE & ": " & imageFromClipboard & vbCrLf & vbCrLf
    output = output & GetFileInfo(imageFromClipboard)

    expireTime = AskExpireTime(output)

    UploadImage imageFromClipboard
    WScript.Quit
End If

' === 4. Chech text from exchange bufer ===
Dim clipCmd, exec, clipText
clipCmd = "powershell -NoLogo -NoProfile -Command ""Get-Clipboard -Raw"""
Set exec = shell.Exec(clipCmd)
clipText = exec.StdOut.ReadAll
clipText = CleanText(clipText)

If clipText = "" Then
    MsgBox TXT_CLIPBOARD_EMPTY, vbExclamation
    WScript.Quit
End If

If fso.FileExists(clipText) Then
    output = output & TXT_MODE & " " & TXT_CLIPBOARD_PATH & vbCrLf
    output = output & TXT_PATH & ": " & clipText & vbCrLf & vbCrLf
    output = output & GetFileInfo(clipText)

    expireTime = AskExpireTime(output)

    UploadFile clipText
    WScript.Quit
End If

' Check for URL
If IsURL(clipText) Then
    Dim choice
    choice = MsgBox(TXT_URL_CHOICE_MSG & vbCrLf & vbCrLf & _
                    TXT_URL & ": " & clipText & vbCrLf & vbCrLf & _
                    TXT_URL_ACTION_SHORTEN & vbCrLf & _
                    TXT_URL_ACTION_DOWNLOAD, _
                    vbYesNo + vbQuestion + vbDefaultButton1, TXT_URL_CHOICE_TITLE)

    output = output & TXT_URL & ": " & clipText & vbCrLf & vbCrLf
    expireTime = AskExpireTime(output)

    If choice = vbYes Then
        ShortenURL clipText
    Else
        UploadFromURL clipText
    End If
    WScript.Quit
End If

output = output & TXT_MODE & " " & TXT_CLIPBOARD_TEXT & vbCrLf & vbCrLf

expireTime = AskExpireTime(output)

Dim clipTempFolder, clipTempFile
clipTempFolder = shell.ExpandEnvironmentStrings("%TEMP%")
clipTempFile   = clipTempFolder & "\rustypaste_clip.txt"

Dim saveCmd
saveCmd = "powershell -NoLogo -NoProfile -Command ""Set-Content -Path '" & Replace(clipTempFile, "'", "''") & "' -Value (Get-Clipboard -Raw) -Encoding UTF8"""
shell.Run saveCmd, 0, True

Dim textFileName
textFileName = GetTextFileName()

Dim curlExec, curlCmd, url
curlCmd = "curl -s -H ""Authorization: " & AUTH_TOKEN & """ -H ""expire:" & expireTime & """ -F ""file=@" & clipTempFile & ";filename=" & textFileName & """ " & SERVER_URL
Set curlExec = shell.Exec(curlCmd)
url = curlExec.StdOut.ReadAll
url = Trim(url)

If url = "" Then
    MsgBox TXT_ERROR, vbCritical
    WScript.Quit
End If

Dim psSetClip
psSetClip = "powershell -NoLogo -NoProfile -Command ""Set-Clipboard -Value '" & Replace(url, "'", "''") & "'"""
shell.Run psSetClip, 0, True

shell.Run "cmd /c echo " & url & " & echo. & echo " & TXT_COPIED & " & echo. & pause", 1, False

' ============================================================================

Function GetFileFromClipboard()
    On Error Resume Next
    
    Dim tempPs1, tempResult, uploadTempFolder, ps1Content, stream
    uploadTempFolder = shell.ExpandEnvironmentStrings("%TEMP%")
    tempPs1 = uploadTempFolder & "\rustypaste_getfile.ps1"
    tempResult = uploadTempFolder & "\rustypaste_result.txt"
    
    ps1Content = "Add-Type -AssemblyName System.Windows.Forms" & vbCrLf & _
                 "$f = [System.Windows.Forms.Clipboard]::GetFileDropList()" & vbCrLf & _
                 "if ($f.Count -gt 0) { $f[0] | Out-File -FilePath '" & tempResult & "' -Encoding UTF8 }"
    
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText ps1Content
    stream.SaveToFile tempPs1, 2
    stream.Close
    Set stream = Nothing
    
    shell.Run "powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File """ & tempPs1 & """", 0, True
    
    If fso.FileExists(tempResult) Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2
        stream.Charset = "utf-8"
        stream.Open
        stream.LoadFromFile tempResult
        GetFileFromClipboard = CleanText(stream.ReadText)
        stream.Close
        Set stream = Nothing
        
        On Error Resume Next
        fso.DeleteFile tempResult
        On Error GoTo 0
    Else
        GetFileFromClipboard = ""
    End If
    
    On Error GoTo 0
End Function

Function GetImageFromClipboard()
    On Error Resume Next

    Dim tempPs1, uploadTempFolder, ps1Content, stream
    uploadTempFolder = shell.ExpandEnvironmentStrings("%TEMP%")
    tempPs1 = uploadTempFolder & "\rustypaste_image.ps1"

    ' PowerShell for save image in file
    ps1Content = "Add-Type -AssemblyName System.Windows.Forms" & vbCrLf & _
                 "Add-Type -AssemblyName System.Drawing" & vbCrLf & _
                 "$img = [System.Windows.Forms.Clipboard]::GetImage()" & vbCrLf & _
                 "if ($img -ne $null) {" & vbCrLf & _
                 "    $path = '" & uploadTempFolder & "\clipboard_image.png'" & vbCrLf & _
                 "    $img.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)" & vbCrLf & _
                 "    Write-Output $path" & vbCrLf & _
                 "    $img.Dispose()" & vbCrLf & _
                 "}"
    
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText ps1Content
    stream.SaveToFile tempPs1, 2
    stream.Close
    Set stream = Nothing
    
    Dim psExec, result
    Set psExec = shell.Exec("powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File """ & tempPs1 & """")
    result = psExec.StdOut.ReadAll
    result = CleanText(result)
    
    If result <> "" And fso.FileExists(result) Then
        GetImageFromClipboard = result
    Else
        GetImageFromClipboard = ""
    End If
    
    On Error GoTo 0
End Function

Function CleanText(text)
    text = Trim(text)
    text = Replace(text, vbCrLf, "")
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")
    CleanText = text
End Function

Function GetFileInfo(filePath)
    Dim info
    info = TXT_INFO & vbCrLf
    info = info & TXT_NAME & ": " & fso.GetFileName(filePath) & vbCrLf
    info = info & TXT_SIZE & ": " & FormatFileSize(fso.GetFile(filePath).Size) & vbCrLf
    info = info & TXT_TYPE & ": " & fso.GetExtensionName(filePath) & vbCrLf
    info = info & TXT_MODIFIED & ": " & fso.GetFile(filePath).DateLastModified & vbCrLf
    GetFileInfo = info
End Function

Function GetPreview(text)
    If Len(text) > 50 Then
        GetPreview = Left(text, 50) & "..."
    Else
        GetPreview = text
    End If
End Function

Function FormatFileSize(size)
    Const KB = 1024, MB = 1048576, GB = 1073741824
    Dim unit
    If LANG = "en" Then
        If size >= GB Then
            unit = " GB"
        ElseIf size >= MB Then
            unit = " MB"
        ElseIf size >= KB Then
            unit = " KB"
        Else
            unit = " bytes"
        End If
    Else
        If size >= GB Then
            unit = " ГБ"
        ElseIf size >= MB Then
            unit = " МБ"
        ElseIf size >= KB Then
            unit = " КБ"
        Else
            unit = " байт"
        End If
    End If
    
    If size >= GB Then
        FormatFileSize = Round(size / GB, 2) & unit
    ElseIf size >= MB Then
        FormatFileSize = Round(size / MB, 2) & unit
    ElseIf size >= KB Then
        FormatFileSize = Round(size / KB, 2) & unit
    Else
        FormatFileSize = size & unit
    End If
End Function

Function GetDateSuffix()
    Dim dt
    dt = Now
    GetDateSuffix = "." & Year(dt) & "-" & Right("0" & Month(dt), 2) & "-" & Right("0" & Day(dt), 2) & "_" & Right("0" & Hour(dt), 2) & "-" & Right("0" & Minute(dt), 2) & "-" & Right("0" & Second(dt), 2)
End Function

Function AddSuffixToFilename(filePath)
    If Not ADD_DATE_SUFFIX Then
        AddSuffixToFilename = fso.GetFileName(filePath)
        Exit Function
    End If
    Dim fileName, baseName, ext
    fileName = fso.GetFileName(filePath)
    ext = fso.GetExtensionName(fileName)
    If ext <> "" Then
        baseName = Left(fileName, Len(fileName) - Len(ext) - 1)
        AddSuffixToFilename = baseName & GetDateSuffix() & "." & ext
    Else
        AddSuffixToFilename = fileName & GetDateSuffix()
    End If
End Function

Function GetTextFileName()
    If Not ADD_DATE_SUFFIX Then
        GetTextFileName = "clipboard.txt"
    Else
        GetTextFileName = "clipboard" & GetDateSuffix() & ".txt"
    End If
End Function

Function GetURLFileName()
    If Not ADD_DATE_SUFFIX Then
        GetURLFileName = "url.link"
    Else
        GetURLFileName = "url" & GetDateSuffix() & ".link"
    End If
End Function

Function GetRemoteFileName(url)
    Dim fileName, baseName, ext, pos
    pos = InStr(url, "?")
    If pos > 0 Then
        url = Left(url, pos - 1)
    End If
    pos = InStrRev(url, "/")
    If pos > 0 Then
        fileName = Mid(url, pos + 1)
    Else
        fileName = "downloaded_file"
    End If
    If ADD_DATE_SUFFIX Then
        ext = ""
        pos = InStrRev(fileName, ".")
        If pos > 0 Then
            ext = Mid(fileName, pos)
            baseName = Left(fileName, pos - 1)
        Else
            baseName = fileName
        End If
        fileName = baseName & GetDateSuffix() & ext
    End If
    GetRemoteFileName = fileName
End Function

Function IsURL(text)
    IsURL = (Left(LCase(text), 7) = "http://" Or Left(LCase(text), 8) = "https://")
End Function

Sub UploadFile(filePath)
    Dim tempPs1, uploadTempFolder, fileName
    fileName = AddSuffixToFilename(filePath)
    uploadTempFolder = shell.ExpandEnvironmentStrings("%TEMP%")
    tempPs1 = uploadTempFolder & "\rustypaste_upload.ps1"

    Dim ps1Content
    ps1Content = "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf & _
                 "$url = curl.exe -s -H 'Authorization: " & AUTH_TOKEN & "' -H 'expire: " & expireTime & "' -F 'file=@" & filePath & ";filename=" & fileName & "' " & SERVER_URL & vbCrLf & _
                 "Set-Clipboard -Value $url" & vbCrLf & _
                 "Write-Host $url" & vbCrLf & _
                 "Start-Sleep -Seconds 3"

    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText ps1Content
    stream.SaveToFile tempPs1, 2
    stream.Close
    Set stream = Nothing

    shell.Run "powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File """ & tempPs1 & """", 1, False
End Sub

Sub UploadImage(filePath)
    Dim tempPs1, uploadTempFolder, fileName
    fileName = AddSuffixToFilename(filePath)
    uploadTempFolder = shell.ExpandEnvironmentStrings("%TEMP%")
    tempPs1 = uploadTempFolder & "\rustypaste_upload.ps1"

    Dim ps1Content
    ps1Content = "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf & _
                 "$url = curl.exe -s -H 'Authorization: " & AUTH_TOKEN & "' -H 'expire: " & expireTime & "' -F 'file=@" & filePath & ";filename=" & fileName & "' " & SERVER_URL & vbCrLf & _
                 "Set-Clipboard -Value $url" & vbCrLf & _
                 "Write-Host $url" & vbCrLf & _
                 "Start-Sleep -Seconds 3"

    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText ps1Content
    stream.SaveToFile tempPs1, 2
    stream.Close
    Set stream = Nothing

    shell.Run "powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File """ & tempPs1 & """", 1, False
End Sub

Sub UploadFromURL(urlText)
    Dim tempPs1, uploadTempFolder, fileName
    fileName = GetRemoteFileName(urlText)
    uploadTempFolder = shell.ExpandEnvironmentStrings("%TEMP%")
    tempPs1 = uploadTempFolder & "\rustypaste_upload.ps1"

    Dim ps1Content
    ps1Content = "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf & _
                 "$url = curl.exe -s -H 'Authorization: " & AUTH_TOKEN & "' -H 'filename: " & fileName & "' -H 'expire: " & expireTime & "' -F 'remote=" & urlText & "' " & SERVER_URL & vbCrLf & _
                 "Set-Clipboard -Value $url" & vbCrLf & _
                 "Write-Host $url" & vbCrLf & _
                 "Start-Sleep -Seconds 3"

    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText ps1Content
    stream.SaveToFile tempPs1, 2
    stream.Close
    Set stream = Nothing

    shell.Run "powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File """ & tempPs1 & """", 1, False
End Sub

Sub ShortenURL(urlText)
    Dim tempPs1, uploadTempFolder, fileName
    fileName = GetURLFileName()
    uploadTempFolder = shell.ExpandEnvironmentStrings("%TEMP%")
    tempPs1 = uploadTempFolder & "\rustypaste_upload.ps1"

    Dim ps1Content
    ps1Content = "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf & _
                 "$url = curl.exe -s -H 'Authorization: " & AUTH_TOKEN & "' -H 'filename: " & fileName & "' -H 'expire: " & expireTime & "' -F 'url=" & urlText & "' " & SERVER_URL & vbCrLf & _
                 "Set-Clipboard -Value $url" & vbCrLf & _
                 "Write-Host $url" & vbCrLf & _
                 "Start-Sleep -Seconds 3"

    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText ps1Content
    stream.SaveToFile tempPs1, 2
    stream.Close
    Set stream = Nothing

    shell.Run "powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File """ & tempPs1 & """", 1, False
End Sub
