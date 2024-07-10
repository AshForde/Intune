Connect-ExchangeOnline -ShowBanner:$false

Import-Csv .\ExternalContacts.csv|%{New-MailContact `
    -Name $_.Name `
    -DisplayName $_.Name `
    -ExternalEmailAddress $_.ExternalEmailAddress `
    -FirstName $_.FirstName `
    -LastName $_.LastName
    }