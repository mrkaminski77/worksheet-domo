$WorkbookPath = 'W:\Reporting\03-Client Reports\ATO\Debt Summary Scorecard\2017\07 July\Debt Daily Scorecard 201707.xlsb'
$WorksheetName = 'MTD Summary'

$ColumnStart = 'B'
$ColumnEnd = 'AJ'
$RowStart = 11
$RowEnd = 41

$subjectLine = 'DEI Scorecard Data'
$to = 'david.leyden@serco-ap.com.au'
$from = 'david.leyden@serco-ap.com.au'
$cc = 'david.leyden@serco-ap.com.au'


$fieldnames = @(
    "Date",
    "TotalAnswered",
    "Reweighted Target",
    "ADJ",
    "Variance ",
    "Actual",
    "Reweighted AHT",
    "Variance  ",
    "Talk",
    "Hold",
    "ACW",
    "Actual(sec)",
    "Missing Agent Data (including preview, and CME Config issues)",
    "Total Workload",
    "Target(Latest LTF)",
    "Type of Increased Purchase",
    "Amount of Increased Purchase",
    "Target (Incl Increased Purchase)",
    "Variance",
    "NPE % (<4%)",
    "Consult % (<5%)",
    "Transfer % (<15%)",
    "Resolve % (>85%)",
    "Nat QA",
    "Local QA",
    "Nat QA Met",
    "Nat QA Not Met",
    "Nat QA Total Assessed",
    "No. Cases",
    "Entered - MTD",
    "Value - MTD",
    "Conversion",
    "Kept Rate (13 - 16 wks after creation)",
    "Kept Value (13 - 16 wks after creation)",
    "Direct Debt % (benchmark 20%)"
)




# SIG # Begin signature block
# MIII3QYJKoZIhvcNAQcCoIIIzjCCCMoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU6zHY+ICcJcyc40J0dvPCoKUO
# JUOgggY1MIIGMTCCBRmgAwIBAgIKSU7nJAABAABQADANBgkqhkiG9w0BAQUFADBj
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
# hvcNAQkEMRYEFDuYd9wcyGWb4ob5VoajpTeNP9KeMA0GCSqGSIb3DQEBAQUABIIB
# AJpcHztzMaVrX9AQZQODrMCxSC+NF1+t3nf6HfUHlRC9D5z1htzLriO8YQoDCHyA
# JUKLHCD1PwgXXSuDduJVXr/kdNC+Nd+dnc5aGhZVCMuBja8fSk9rsUFgs+U2Wfw5
# 8wpXt0kCD7oAf/VKwirKZKUYFn0g203LK5DAn5Jv8SX/o/np6kAavjpvdmuuKHEs
# xNiBoQxH1iNL3Y3+E2oYbQF7beYwvLEUwCNbbxKwglkuRmLO9T+5HHL8CLlOVCZN
# iUxMUmov78W+nBwRRXMhlxDCRv4oaN3WP7XfXVIWP3mGtfaFcqnya7P3fL/bcT8O
# Z9im45MCs2qntbZp/wBmYiQ=
# SIG # End signature block
