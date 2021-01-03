$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)
$Mail.To = "Donuts@wavetronix.com"
$Mail.Subject = "Donuts!"
$Mail.Body ="I'm pleased to announce that I'm bringing fresh donuts for everyone next time I'm in the office."
$Mail.Send()
