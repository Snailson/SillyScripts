$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)
$Mail.To = "$emailaddress"
$Mail.Subject = "Donuts"
$Mail.Body ="Free Donuts tomorrow from me!"
$Mail.Send()