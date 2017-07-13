$WorkbookPath = 'W:\Reporting\03-Client Reports\ATO\Summary Scorecard\2017\06 June\Daily Scorecard  version 201706.xlsb'
$WorksheetName = 'MTD Summary'

$ColumnStart = 'B'
$ColumnEnd = 'AF'
$RowStart = 10
$RowEnd = 40



$subjectLine = 'CSS Scorecard Data'
$to = '04b9e7f931384a14b73e7264724a16d4@serco-ap-au.mail.domo.com'
$cc = 'david.leyden@serco-ap.com.au'
$from = 'david.leyden@serco-ap.com.au'

$fieldnames = @(
    'Date',
    "TotalAnswered",
    "Modified Forecast",
    "Reweighted Target",
    "ADJ",
    "Variance 1",
    "Actual",
    "Reweighted AHT",
    "Variance 2",
    "Talk",
    "Hold",
    "ACW",
    "Actual(sec)",
    "Actual Missing WL",
    "Actual Total Workload",
    "Target(Latest LTF)",
    "Type of Increased Purchase",
    "Amount of Increased Purchase",
    "Target (Incl Increased Purchase)",
    "Variance 3",
    "NPE % (<4%)",
    "Consult % (<5%)",
    "Transfer % (<15%)",
    "Resolve % (>85%)",
    "Escalations % (Rolling)",
    "CSS QA KPI % (Rolling)",
    "QA PF KPI % (Rolling)",
    "Local CSS QA % (Rolling)",
    "CSS QA KPI Met",
    "CSS QA KPI Not Met",
    "CSS QA KPI Total Assessed"
)



# SIG # Begin signature block
# MIII3QYJKoZIhvcNAQcCoIIIzjCCCMoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUmXyTVovfK0y06FqadR7pRFq4
# XGCgggY1MIIGMTCCBRmgAwIBAgIKSU7nJAABAABQADANBgkqhkiG9w0BAQUFADBj
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
# hvcNAQkEMRYEFB0h6j1hCCh8wp8k1ysO5GUjwmtlMA0GCSqGSIb3DQEBAQUABIIB
# AGWz/4k3ax18VB5+AFZ+y6CJFwk3O0n0oQkcNsSFQiZ3E8njPsihSon4PlwkJRFE
# 9r1ttvVvCtcF3APgSoazwPbkjKYcV+cO2QFNgpORXzntIDcEuDEzXOrWdRNjslAn
# xlqkIYMt3o5xcefuHrPttXTwV8ii1rTIT2qbhOL93AfwJknbgXxVWT9jrhYWpR9L
# rnayU705ZIcd7lEAeUAc/fZpqSrkBHfVtOtHm4EUdnWtwvknRuh+uMdJ6lZ7udo8
# XIYm2onyVli8ij5HAIQGmpYVrHUYWYO7Zlz/Lb3EX5JRyAnNlB4GhCoJ1HkSzgDo
# rwtWlVxlNTbysgO6pVPtPkA=
# SIG # End signature block
