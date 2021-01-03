$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)
$Mail.To = "chad@wavetronix.com"
$Mail.Subject = "Donuts3!"
$Mail.Body ="I'm pleased to announce that I'm bringing fresh donuts for everyone next time I'm in the office."
$Mail.Send()
exit
