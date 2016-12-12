function GlobalOption {
<#
    .SYNOPSIS
        Initializes a new PScribo global options/settings.
    .NOTES
        Options are reset upon each invocation.
#>
    [CmdletBinding(DefaultParameterSetName = 'Margin')]
    param (
        ## Forces document header to be displayed in upper case.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $ForceUppercaseHeader,

        ## Forces all section headers to be displayed in upper case.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $ForceUppercaseSection,

        ## Enable section/heading numbering
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $EnableSectionNumbering,

        ## Default space replacement separator
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('Separator')]
        [AllowNull()]
        [ValidateLength(0,1)]
        [System.String] $SpaceSeparator,

        ## Default page top, bottom, left and right margin (pt)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Margin')]
        [System.UInt16] $Margin = 72,

        ## Default page top and bottom margins (pt)
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'CustomMargin')]
        [System.UInt16] $MarginTopAndBottom,

        ## Default page left and right margins (pt)
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'CustomMargin')]
        [System.UInt16] $MarginLeftAndRight,

        ## Default page size
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('A4','Legal','Letter')]
        [System.String] $PageSize = 'A4',

        ## Page orientation
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('Portrait','Landscape')]
        [System.String] $Orientation = 'Portrait',

        ## Default document font(s)
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String[]] $DefaultFont = @('Calibri','Candara','Segoe','Segoe UI','Optima','Arial','Sans-Serif')
    )
    process {

        $localized.DocumentOptions | WriteLog;
        if ($SpaceSeparator) {
            WriteLog -Message ($localized.DocumentOptionSpaceSeparator -f $SpaceSeparator);
            $pscriboDocument.Options['SpaceSeparator'] = $SpaceSeparator;
        }

        if ($ForceUppercaseHeader) {
            $localized.DocumentOptionUppercaseHeadings | WriteLog;
            $pscriboDocument.Options['ForceUppercaseHeader'] = $true;
            $pscriboDocument.Name = $pscriboDocument.Name.ToUpper();
        } #end if ForceUppercaseHeader

        if ($ForceUppercaseSection) {
            $localized.DocumentOptionUppercaseSections | WriteLog;
            $pscriboDocument.Options['ForceUppercaseSection'] = $true;
        } #end if ForceUppercaseSection

        if ($EnableSectionNumbering) {
            $localized.DocumentOptionSectionNumbering | WriteLog;
            $pscriboDocument.Options['EnableSectionNumbering'] = $true;
        }

        if ($DefaultFont) {
            WriteLog -Message ($localized.DocumentOptionDefaultFont -f ([System.String]::Join(', ', $DefaultFont)));
            $pscriboDocument.Options['DefaultFont'] = $DefaultFont;
        }

        if ($PSCmdlet.ParameterSetName -eq 'CustomMargin') {
            if ($MarginTopAndBottom -eq 0) { $MarginTopAndBottom = 72; }
            if ($MarginLeftAndRight -eq 0) { $MarginTopAndBottom = 72; }
            $pscriboDocument.Options['MarginTop'] = ConvertPtToMm -Point $MarginTopAndBottom;
            $pscriboDocument.Options['MarginBottom'] = $pscriboDocument.Options['MarginTop'];
            $pscriboDocument.Options['MarginLeft'] = ConvertPtToMm -Point $MarginLeftAndRight;
            $pscriboDocument.Options['MarginRight'] = $pscriboDocument.Options['MarginLeft'];
        }
        else {
            $pscriboDocument.Options['MarginTop'] = ConvertPtToMm -Point $Margin;
            $pscriboDocument.Options['MarginBottom'] = $pscriboDocument.Options['MarginTop'];
            $pscriboDocument.Options['MarginLeft'] = $pscriboDocument.Options['MarginTop'];
            $pscriboDocument.Options['MarginRight'] = $pscriboDocument.Options['MarginTop'];
        }
        WriteLog -Message ($localized.DocumentOptionPageTopMargin -f $pscriboDocument.Options['MarginTop']);
        WriteLog -Message ($localized.DocumentOptionPageRightMargin -f $pscriboDocument.Options['MarginRight']);
        WriteLog -Message ($localized.DocumentOptionPageBottomMargin -f $pscriboDocument.Options['MarginBottom']);
        WriteLog -Message ($localized.DocumentOptionPageLeftMargin -f $pscriboDocument.Options['MarginLeft']);

        ## Convert page size
        ($localized.DocumentOptionPageSize -f $PageSize) | WriteLog;
        switch ($PageSize) {
            'A4' {
                $pscriboDocument.Options['PageWidth'] = 210.0;
                $pscriboDocument.Options['PageHeight'] = 297.0;
            }
            'Legal' {
                $pscriboDocument.Options['PageWidth'] = 215.9;
                $pscriboDocument.Options['PageHeight'] = 355.6;
            }
            'Letter' {
                $pscriboDocument.Options['PageWidth'] = 215.9;
                $pscriboDocument.Options['PageHeight'] = 279.4;
            }
        } #end switch

        ## Convert page size
        ($localized.DocumentOptionPageOrientation -f $Orientation) | WriteLog;
        if ($Orientation -eq 'Landscape') {
            ## Swap the height/width measurements
            $pageHeight = $pscriboDocument.Options['PageHeight'];
            $pscriboDocument.Options['PageHeight'] = $pscriboDocument.Options['PageWidth'];
            $pscriboDocument.Options['PageWidth'] = $pageHeight;
        }
        ($localized.DocumentOptionPageHeight -f $pscriboDocument.Options['PageHeight']) | WriteLog;
        ($localized.DocumentOptionPageWidth -f $pscriboDocument.Options['PageWidth']) | WriteLog;

    } #end process
} #end function GlobalOption

# SIG # Begin signature block
# MIIXtwYJKoZIhvcNAQcCoIIXqDCCF6QCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUH4wSoL1jqDGwPrLMtOgsVe0t
# rVSgghLqMIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
# AQUFADCBizELMAkGA1UEBhMCWkExFTATBgNVBAgTDFdlc3Rlcm4gQ2FwZTEUMBIG
# A1UEBxMLRHVyYmFudmlsbGUxDzANBgNVBAoTBlRoYXd0ZTEdMBsGA1UECxMUVGhh
# d3RlIENlcnRpZmljYXRpb24xHzAdBgNVBAMTFlRoYXd0ZSBUaW1lc3RhbXBpbmcg
# Q0EwHhcNMTIxMjIxMDAwMDAwWhcNMjAxMjMwMjM1OTU5WjBeMQswCQYDVQQGEwJV
# UzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFu
# dGVjIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMjCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBALGss0lUS5ccEgrYJXmRIlcqb9y4JsRDc2vCvy5Q
# WvsUwnaOQwElQ7Sh4kX06Ld7w3TMIte0lAAC903tv7S3RCRrzV9FO9FEzkMScxeC
# i2m0K8uZHqxyGyZNcR+xMd37UWECU6aq9UksBXhFpS+JzueZ5/6M4lc/PcaS3Er4
# ezPkeQr78HWIQZz/xQNRmarXbJ+TaYdlKYOFwmAUxMjJOxTawIHwHw103pIiq8r3
# +3R8J+b3Sht/p8OeLa6K6qbmqicWfWH3mHERvOJQoUvlXfrlDqcsn6plINPYlujI
# fKVOSET/GeJEB5IL12iEgF1qeGRFzWBGflTBE3zFefHJwXECAwEAAaOB+jCB9zAd
# BgNVHQ4EFgQUX5r1blzMzHSa1N197z/b7EyALt0wMgYIKwYBBQUHAQEEJjAkMCIG
# CCsGAQUFBzABhhZodHRwOi8vb2NzcC50aGF3dGUuY29tMBIGA1UdEwEB/wQIMAYB
# Af8CAQAwPwYDVR0fBDgwNjA0oDKgMIYuaHR0cDovL2NybC50aGF3dGUuY29tL1Ro
# YXd0ZVRpbWVzdGFtcGluZ0NBLmNybDATBgNVHSUEDDAKBggrBgEFBQcDCDAOBgNV
# HQ8BAf8EBAMCAQYwKAYDVR0RBCEwH6QdMBsxGTAXBgNVBAMTEFRpbWVTdGFtcC0y
# MDQ4LTEwDQYJKoZIhvcNAQEFBQADgYEAAwmbj3nvf1kwqu9otfrjCR27T4IGXTdf
# plKfFo3qHJIJRG71betYfDDo+WmNI3MLEm9Hqa45EfgqsZuwGsOO61mWAK3ODE2y
# 0DGmCFwqevzieh1XTKhlGOl5QGIllm7HxzdqgyEIjkHq3dlXPx13SYcqFgZepjhq
# IhKjURmDfrYwggSjMIIDi6ADAgECAhAOz/Q4yP6/NW4E2GqYGxpQMA0GCSqGSIb3
# DQEBBQUAMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3Jh
# dGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBD
# QSAtIEcyMB4XDTEyMTAxODAwMDAwMFoXDTIwMTIyOTIzNTk1OVowYjELMAkGA1UE
# BhMCVVMxHTAbBgNVBAoTFFN5bWFudGVjIENvcnBvcmF0aW9uMTQwMgYDVQQDEytT
# eW1hbnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIFNpZ25lciAtIEc0MIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAomMLOUS4uyOnREm7Dv+h8GEKU5Ow
# mNutLA9KxW7/hjxTVQ8VzgQ/K/2plpbZvmF5C1vJTIZ25eBDSyKV7sIrQ8Gf2Gi0
# jkBP7oU4uRHFI/JkWPAVMm9OV6GuiKQC1yoezUvh3WPVF4kyW7BemVqonShQDhfu
# ltthO0VRHc8SVguSR/yrrvZmPUescHLnkudfzRC5xINklBm9JYDh6NIipdC6Anqh
# d5NbZcPuF3S8QYYq3AhMjJKMkS2ed0QfaNaodHfbDlsyi1aLM73ZY8hJnTrFxeoz
# C9Lxoxv0i77Zs1eLO94Ep3oisiSuLsdwxb5OgyYI+wu9qU+ZCOEQKHKqzQIDAQAB
# o4IBVzCCAVMwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAO
# BgNVHQ8BAf8EBAMCB4AwcwYIKwYBBQUHAQEEZzBlMCoGCCsGAQUFBzABhh5odHRw
# Oi8vdHMtb2NzcC53cy5zeW1hbnRlYy5jb20wNwYIKwYBBQUHMAKGK2h0dHA6Ly90
# cy1haWEud3Muc3ltYW50ZWMuY29tL3Rzcy1jYS1nMi5jZXIwPAYDVR0fBDUwMzAx
# oC+gLYYraHR0cDovL3RzLWNybC53cy5zeW1hbnRlYy5jb20vdHNzLWNhLWcyLmNy
# bDAoBgNVHREEITAfpB0wGzEZMBcGA1UEAxMQVGltZVN0YW1wLTIwNDgtMjAdBgNV
# HQ4EFgQURsZpow5KFB7VTNpSYxc/Xja8DeYwHwYDVR0jBBgwFoAUX5r1blzMzHSa
# 1N197z/b7EyALt0wDQYJKoZIhvcNAQEFBQADggEBAHg7tJEqAEzwj2IwN3ijhCcH
# bxiy3iXcoNSUA6qGTiWfmkADHN3O43nLIWgG2rYytG2/9CwmYzPkSWRtDebDZw73
# BaQ1bHyJFsbpst+y6d0gxnEPzZV03LZc3r03H0N45ni1zSgEIKOq8UvEiCmRDoDR
# EfzdXHZuT14ORUZBbg2w6jiasTraCXEQ/Bx5tIB7rGn0/Zy2DBYr8X9bCT2bW+IW
# yhOBbQAuOA2oKY8s4bL0WqkBrxWcLC9JG9siu8P+eJRRw4axgohd8D20UaF5Mysu
# e7ncIAkTcetqGVvP6KUwVyyJST+5z3/Jvz4iaGNTmr1pdKzFHTx/kuDDvBzYBHUw
# ggUZMIIEAaADAgECAhADViTO4HBjoJNSwH9//cwJMA0GCSqGSIb3DQEBCwUAMHIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJ
# RCBDb2RlIFNpZ25pbmcgQ0EwHhcNMTUwNTE5MDAwMDAwWhcNMTcwODIzMTIwMDAw
# WjBgMQswCQYDVQQGEwJHQjEPMA0GA1UEBxMGT3hmb3JkMR8wHQYDVQQKExZWaXJ0
# dWFsIEVuZ2luZSBMaW1pdGVkMR8wHQYDVQQDExZWaXJ0dWFsIEVuZ2luZSBMaW1p
# dGVkMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAqLQmabdimcQtYPTQ
# 9RSjv3ThEmFTRJt/MzseYYtZpBTcR6BnSfj8RfkC4aGZvspFgH0cGP/SNJh1w67b
# iX9oT5NFL9sUJHUsVdyPBA1LhpWcF09PP28mGGKO3oQHI4hTLD8etiIlF9qFantd
# 1Pmo0jdqT4uErSmx0m4kYGUUTa5ZPAK0UZSuAiNX6iNIL+rj/BPbI3nuPJzzx438
# oHYkZGRtsx11+pLA6hIKyUzRuIDoI7JQ0nZ0MkCziVyc6xGfS54JVLaVCEteTKPz
# Gc4yyvCqp6Tfe9gs8UuxJiEMdH5fvllTU4aoXbm+W8tonkE7i/19rv8S1A2VPiVV
# xNLbpwIDAQABo4IBuzCCAbcwHwYDVR0jBBgwFoAUWsS5eyoKo6XqcQPAYPkt9mV1
# DlgwHQYDVR0OBBYEFP2RNOWYipdNCSRVb5jIcyRp9tUDMA4GA1UdDwEB/wQEAwIH
# gDATBgNVHSUEDDAKBggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Axhi9odHRwOi8v
# Y3JsMy5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDA1oDOgMYYv
# aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmww
# QgYDVR0gBDswOTA3BglghkgBhv1sAwEwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93
# d3cuZGlnaWNlcnQuY29tL0NQUzCBhAYIKwYBBQUHAQEEeDB2MCQGCCsGAQUFBzAB
# hhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTgYIKwYBBQUHMAKGQmh0dHA6Ly9j
# YWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFNIQTJBc3N1cmVkSURDb2RlU2ln
# bmluZ0NBLmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3DQEBCwUAA4IBAQCclXHR
# DhDyJr81eiD0x+AL04ryDwdKT+PooKYgOxc7EhRn59ogxNO7jApQPSVo0I11Zfm6
# zQ6K6RPWhxDenflf2vMx7a0tIZlpHhq2F8praAMykK7THA9F3AUxIb/lWHGZCock
# yD/GQvJek3LSC5NjkwQbnubWYF/XZTDzX/mJGU2DcG1OGameffR1V3xODHcUE/K3
# PWy1bzixwbQCQA96GKNCWow4/mEW31cupHHSo+XVxmjTAoC93yllE9f4Kdv6F29H
# bRk0Go8Yn8WjWeLE/htxW/8ruIj0KnWkG+YwmZD+nTegYU6RvAV9HbJJYUEIfhVy
# 3DeK5OlY9ima2sdtMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkq
# hkiG9w0BAQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5j
# MRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBB
# c3N1cmVkIElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAwWhcNMjgxMDIyMTIwMDAw
# WjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3Vy
# ZWQgSUQgQ29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIB
# CgKCAQEA+NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/5aid2zLXcep2nQUut4/6
# kkPApfmJ1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH03sjlOSRI5aQd4L5oYQj
# ZhJUM1B0sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxKhwjfDPXiTWAYvqrEsq5w
# MWYzcT6scKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr/mzLfnQ5Ng2Q7+S1TqSp
# 6moKq4TzrGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi6CxR93O8vYWxYoNzQYIH
# 5DiLanMg0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCCAckwEgYDVR0TAQH/BAgw
# BgEB/wIBADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMweQYI
# KwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6
# Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmww
# OqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJ
# RFJvb3RDQS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1sAAIEMCowKAYIKwYBBQUH
# AgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCgYIYIZIAYb9bAMwHQYD
# VR0OBBYEFFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1UdIwQYMBaAFEXroq/0ksuC
# MS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+7A1aJLPzItEVyCx8JSl2
# qB1dHC06GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbRknUPUbRupY5a4l4kgU4Q
# pO4/cY5jDhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7uq+1UcKNJK4kxscnKqEp
# KBo6cSgCPC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7qPjFEmifz0DLQESlE/Dm
# ZAwlCEIysjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPas7CM1ekN3fYBIM6ZMWM9
# CBoYs4GbT8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR6mhsRDKyZqHnGKSaZFHv
# MYIENzCCBDMCAQEwgYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0
# IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNl
# cnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQQIQA1YkzuBwY6CTUsB/
# f/3MCTAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUHUfs1E4YrdAAPmrxrtBkyUuJg9YwDQYJKoZI
# hvcNAQEBBQAEggEAOcrMgPQh88hKrb/cBF2Q0IGwUnjbmoMvVL6nstHIGs7A0aO7
# Zt13uAAyFJwG0cZs2XBcVAaQl+Ew7/Bx2HNNkhn1N+JSfxshMmQmyYtfmB8IOprS
# 35uGMicId3b5so71S05o3GZTneLcoW7O24BV+2en1Vnx9/rgoDYsZhIYSDcifjL2
# h/XC3jpYaC9xmS9FlFjgO6GEHtyiSuAcCaDd2lauzLjhvVs7EHOO5OpRxoe9VAs7
# c4pkYr1xA0VZnZEYTy/rVKT7gkL1Q5IosXz+zuVAQ3xfPyl39WQKjjzYYRtsJIvg
# Kc7Gr3hGok1r8nuxnx4XlNgRbVLpO3Mmi8juVaGCAgswggIHBgkqhkiG9w0BCQYx
# ggH4MIIB9AIBATByMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBD
# b3Jwb3JhdGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2
# aWNlcyBDQSAtIEcyAhAOz/Q4yP6/NW4E2GqYGxpQMAkGBSsOAwIaBQCgXTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA4MDEyMTQ2
# MjBaMCMGCSqGSIb3DQEJBDEWBBTtbKjP6Xl/lVa+dnqu9VmtTaYRdDANBgkqhkiG
# 9w0BAQEFAASCAQB87gutx2OMGx9oK2puOl0ytkFmprhw7WornGiIKyYLYCvmWmW+
# vemBGjOTETKcsb275QnQ1Bz5C7D86QXzSuQNLW81iSr88W8KTDMjX4kRWu0QMgt3
# ipJAN3psPyi1rGBqyKGDP3MyeEOxlIIFzgismK7Fj0qTjQoJwS9Nv4Q5J0ac/KpT
# 2XzgCP5GAH4fGGp2LQR/pGsUPXJ4fkRu9yMeO9+45n7WByN3NYpVd7yp0hxbh097
# 1uj4JYph7R4Kxf8BqsLyRKUa978b0MoQGzybByb85HS4TbEaBpryDiWf0tyDzPYS
# eB45+KgospxIs3suP5Zwi6FBAhEl4kHbBaT6
# SIG # End signature block
