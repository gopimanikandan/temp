$assemblyPath = 'C:\Users\Gopimanikandan\Desktop\outlook\ReadOutlookMsg.dll'

Add-Type -Path $assemblyPath

$message = New-Object -TypeName  'ReadOutlookMsg.Message'

$email = $message.ReadMessage("RE_ Issue with PIN reset.msg")

Write-Output $email
