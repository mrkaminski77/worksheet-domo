$smtp_server = 'smtpgw.ap.serco.com'
$temp_file = 'C:\temp\data.csv'

try {
    Remove-Item -Path $temp_file -Force
} catch {}

$Excel = New-Object -ComObject Excel.Application

$workbook = $Excel.Workbooks.Open($WorkbookPath)
$worksheet = $workbook.Worksheets($WorksheetName)

# do this for tab separated values
#$data = "$($fieldnames -join "`t")`n"
#for ($i=$RowStart; $i -le $RowEnd; $i++) {
#    $line = "$($worksheet.Range("$ColumnStart$($i):$ColumnEnd$($i)").Value() -join "`t")`n"
#    $data = "$($data)$($line)"
#}


try {
    "$ColumnStart$($RowStart):$ColumnEnd$($RowEnd)"
    if ($fieldnames.Count -ne $worksheet.Range("$ColumnStart$($RowStart):$ColumnEnd$($RowEnd)").Columns().Count ) {
        "Field count doesn't match"
        "Field list has $($fieldnames.Count)"
        "Column has $($worksheet.Range("$ColumnStart$($RowStart):$ColumnEnd$($RowEnd)").Columns().Count)"
        throw
    }
    $rows = $worksheet.Range("$ColumnStart$($RowStart):$ColumnEnd$($RowEnd)").Rows()
    $records = foreach($row in $rows){
        $i = 0
        $record = @{}
        foreach($cell in $row.Cells()){
            #Write-Host "$($fieldnames[$i]),$($cell.Value())"
            $record.Add($fieldnames[$i++],($cell.Value() -replace '^-$',''))
        }
        (New-Object -TypeName PSObject -Property $record)
    }

    $records | Export-Csv $temp_file -NoTypeInformation 

    #Set-Content -Path 'c:\temp\data.csv' -Value $data -Force

    $mod1 = New-Object Net.Mail.SmtpClient($smtp_server)

    $email = New-Object System.Net.Mail.Mailmessage
    $email.Subject = $subjectLine
    $email.Body = ''
    $email.IsBodyHTML = $false
    $email.from = ($from)
    $email.To.Add($to)
    $email.CC.Add($cc)
    $email.Attachments.Add($temp_file)

    $mod1.UseDefaultCredentials = 1

    $mod1.Send($email)
}
catch{
    
    throw
}

$workbook.Close(0)

$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
$workbook = $null
$worksheet = $null
$email = $null
$Excel = $null
$mod1 = $null
[System.GC]::Collect()


# SIG # Begin signature block
# MIII3QYJKoZIhvcNAQcCoIIIzjCCCMoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUN38wa8aGgz0W1IzGUGT9TRVf
# 6fCgggY1MIIGMTCCBRmgAwIBAgIKSU7nJAABAABQADANBgkqhkiG9w0BAQUFADBj
# MRIwEAYKCZImiZPyLGQBGRYCYXUxEzARBgoJkiaJk/IsZAEZFgNjb20xGDAWBgoJ
# kiaJk/IsZAEZFghzZXJjb2JwbzEeMBwGA1UEAxMVc2VyY29icG8tRVhCRU5EQzAy
# LUNBMB4XDTE3MDUyOTA3MTE0NVoXDTE4MDUyOTA3MTE0NVowgZ8xEjAQBgoJkiaJ
# k/IsZAEZFgJhdTETMBEGCgmSJomT8ixkARkWA2NvbTEYMBYGCgmSJomT8ixkARkW
# CHNlcmNvYnBvMREwDwYDVQQLEwhFeGNlbGlvcjEOMAwGA1UECxMFVXNlcnMxDDAK
# BgNVBAsTA1ZJQzESMBAGA1UECxMJTWVsYm91cm5lMRUwEwYDVQQDEwxEYXZpZCBM
# ZXlkZW4wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCm4FqaO84rZlLj
# PF3SfFgwpREKLNylT5DyOPz0Q5DTvxj2aRN/9STUYeviZTqYT1wrbyVFC48ByaUl
# zs0oeHZP4obJz7jfOBfEgNFCymHo+fkN0+nOGEHwnbWyZGg0puHZN79P8NOASta7
# Q5tPkRRvgFV/j1AoAxs3kCiwbYcXQmMtuZ3WfUnD76o8r9POY47tZNoi2yFGWtPO
# OL3sSiXWtm6XKHEGyQw1d+WRN4+j5gDqka9UF9EN02aR3EUCqTPoe4aY2pho0oNm
# pRvI3IlzgkQ2oEB+q2YMmlad54o7WmNC9h8IOV3dvUQaKj2P/T9E/gpSdpWVBCOL
# 7MMOqPmlAgMBAAGjggKoMIICpDA8BgkrBgEEAYI3FQcELzAtBiUrBgEEAYI3FQjQ
# kwiEhvxRgtmFB4S6nkiCypUhgTyD9Jky4rdQAgFlAgEAMBMGA1UdJQQMMAoGCCsG
# AQUFBwMDMA4GA1UdDwEB/wQEAwIHgDAbBgkrBgEEAYI3FQoEDjAMMAoGCCsGAQUF
# BwMDMB0GA1UdDgQWBBTXV1y3ze996JzB19bbCrrnImKjdTAfBgNVHSMEGDAWgBQg
# cQm4j0b8NvqaIE/ouImrb41ckjCB3AYDVR0fBIHUMIHRMIHOoIHLoIHIhoHFbGRh
# cDovLy9DTj1zZXJjb2Jwby1FWEJFTkRDMDItQ0EsQ049RVhCRU5EQzAyLENOPUNE
# UCxDTj1QdWJsaWMlMjBLZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25m
# aWd1cmF0aW9uLERDPXNlcmNvYnBvLERDPWNvbSxEQz1hdT9jZXJ0aWZpY2F0ZVJl
# dm9jYXRpb25MaXN0P2Jhc2U/b2JqZWN0Q2xhc3M9Y1JMRGlzdHJpYnV0aW9uUG9p
# bnQwgc4GCCsGAQUFBwEBBIHBMIG+MIG7BggrBgEFBQcwAoaBrmxkYXA6Ly8vQ049
# c2VyY29icG8tRVhCRU5EQzAyLUNBLENOPUFJQSxDTj1QdWJsaWMlMjBLZXklMjBT
# ZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPXNlcmNvYnBv
# LERDPWNvbSxEQz1hdT9jQUNlcnRpZmljYXRlP2Jhc2U/b2JqZWN0Q2xhc3M9Y2Vy
# dGlmaWNhdGlvbkF1dGhvcml0eTAyBgNVHREEKzApoCcGCisGAQQBgjcUAgOgGQwX
# ZGxleWRlbkBzZXJjb2Jwby5jb20uYXUwDQYJKoZIhvcNAQEFBQADggEBAFhciq8Z
# E270aT5LVnupHGduNpak4M0Lk5+hCx2aZSc5mwYZjPkofH97MDrSeddl4k9urB+Q
# ROlQjQNv+fTS+/mI1iazvDXH0Z6AwXELxefwZIR1HXoiyRm/WLn9auHHQC5a7qGh
# T1TvVz9YiNmjFt6HHFTWPw90PUsQT4t4p17P+n8owdfz0TtCb+af5GjebOBoyKg3
# lUN1M+/4XMvPJPSgDMnm6oc6C4V+JoKqw8PQy2GPj8nH94tMtJXKo4wbPtvAVudZ
# 3BC52D32LgedTEDBPZfjzAwP/fd+OUQ03AoCyhgXG5YlVSISQWP4D0VLITzzdf04
# H+Qx6k2rpDso7xMxggISMIICDgIBATBxMGMxEjAQBgoJkiaJk/IsZAEZFgJhdTET
# MBEGCgmSJomT8ixkARkWA2NvbTEYMBYGCgmSJomT8ixkARkWCHNlcmNvYnBvMR4w
# HAYDVQQDExVzZXJjb2Jwby1FWEJFTkRDMDItQ0ECCklO5yQAAQAAUAAwCQYFKw4D
# AhoFAKB4MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZI
# hvcNAQkEMRYEFNxPPazXJyggrUs/bh7SuQAU3Q1mMA0GCSqGSIb3DQEBAQUABIIB
# ACh7g3JcblPwdnwq2tMsuHdeSVIWWEWAiu1A6SzIc2S/OyoQd6K94kzubspNMJpa
# mXpiMNU74vqJbWnYqTgTjl1Tg+jaHZesI9tKeTC10fw3/BLlbV6Qie3/OTz4Hco/
# mYVhaAZdbgk36o/IdBzwI0BlWbpwPBecAgiLjKE4cA/CeppVo5ndjf2i4n8ps4gL
# NAn3VBxQ5CAqolRzjbeJn0ONB7RmvKqGzpl6dYwyWJQMfSP8Uit9piIh7AkU3do5
# NKpoBFbnh/vqhsRxRFPUr2oSOlxRQ5iqDxQO0EVqdzImlo4Mtpj8slglMNdFNUFE
# w/06sEUih9xNJDecY1ekO8E=
# SIG # End signature block
