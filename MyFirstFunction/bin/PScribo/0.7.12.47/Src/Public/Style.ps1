function Resolve-PScriboStyleColor {
<#
    .SYNOPSIS
        Resolves a HTML color format or Word color constant to a RGB value
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [ValidateNotNull()]
        [System.String] $Color
    )
    begin {

        # http://www.jadecat.com/tuts/colorsplus.html
        $wordColorConstants = @{
            AliceBlue = 'F0F8FF'; AntiqueWhite = 'FAEBD7'; Aqua = '00FFFF'; Aquamarine = '7FFFD4'; Azure = 'F0FFFF'; Beige = 'F5F5DC';
            Bisque = 'FFE4C4'; Black = '000000'; BlanchedAlmond = 'FFEBCD'; Blue = '0000FF'; BlueViolet = '8A2BE2'; Brown = 'A52A2A';
            BurlyWood = 'DEB887'; CadetBlue = '5F9EA0'; Chartreuse = '7FFF00'; Chocolate = 'D2691E'; Coral = 'FF7F50';
            CornflowerBlue = '6495ED'; Cornsilk = 'FFF8DC'; Crimson = 'DC143C'; Cyan = '00FFFF'; DarkBlue = '00008B'; DarkCyan = '008B8B';
            DarkGoldenrod = 'B8860B'; DarkGray = 'A9A9A9'; DarkGreen = '006400'; DarkKhaki = 'BDB76B'; DarkMagenta = '8B008B';
            DarkOliveGreen = '556B2F'; DarkOrange = 'FF8C00'; DarkOrchid = '9932CC'; DarkRed = '8B0000'; DarkSalmon = 'E9967A';
            DarkSeaGreen = '8FBC8F'; DarkSlateBlue = '483D8B'; DarkSlateGray = '2F4F4F'; DarkTurquoise = '00CED1'; DarkViolet = '9400D3';
            DeepPink = 'FF1493'; DeepSkyBlue = '00BFFF'; DimGray = '696969'; DodgerBlue = '1E90FF'; Firebrick = 'B22222';
            FloralWhite = 'FFFAF0'; ForestGreen = '228B22'; Fuchsia = 'FF00FF'; Gainsboro = 'DCDCDC'; GhostWhite = 'F8F8FF';
            Gold = 'FFD700'; Goldenrod = 'DAA520'; Gray = '808080'; Green = '008000'; GreenYellow = 'ADFF2F'; Honeydew = 'F0FFF0';
            HotPink = 'FF69B4'; IndianRed = 'CD5C5C'; Indigo = '4B0082'; Ivory = 'FFFFF0'; Khaki = 'F0E68C'; Lavender = 'E6E6FA';
            LavenderBlush = 'FFF0F5'; LawnGreen = '7CFC00'; LemonChiffon = 'FFFACD'; LightBlue = 'ADD8E6'; LightCoral = 'F08080';
            LightCyan = 'E0FFFF'; LightGoldenrodYellow = 'FAFAD2'; LightGreen = '90EE90'; LightGrey = 'D3D3D3'; LightPink = 'FFB6C1';
            LightSalmon = 'FFA07A'; LightSeaGreen = '20B2AA'; LightSkyBlue = '87CEFA'; LightSlateGray = '778899'; LightSteelBlue = 'B0C4DE';
            LightYellow = 'FFFFE0'; Lime = '00FF00'; LimeGreen = '32CD32'; Linen = 'FAF0E6'; Magenta = 'FF00FF'; Maroon = '800000';
            McMintGreen = 'BED6C9'; MediumAuqamarine = '66CDAA'; MediumBlue = '0000CD'; MediumOrchid = 'BA55D3'; MediumPurple = '9370D8';
            MediumSeaGreen = '3CB371'; MediumSlateBlue = '7B68EE'; MediumSpringGreen = '00FA9A'; MediumTurquoise = '48D1CC';
            MediumVioletRed = 'C71585'; MidnightBlue = '191970'; MintCream = 'F5FFFA'; MistyRose = 'FFE4E1'; Moccasin = 'FFE4B5';
            NavajoWhite = 'FFDEAD'; Navy = '000080'; OldLace = 'FDF5E6'; Olive = '808000'; OliveDrab = '688E23'; Orange = 'FFA500';
            OrangeRed = 'FF4500'; Orchid = 'DA70D6'; PaleGoldenRod = 'EEE8AA'; PaleGreen = '98FB98'; PaleTurquoise = 'AFEEEE';
            PaleVioletRed = 'D87093'; PapayaWhip = 'FFEFD5'; PeachPuff = 'FFDAB9'; Peru = 'CD853F'; Pink = 'FFC0CB'; Plum = 'DDA0DD';
            PowderBlue = 'B0E0E6'; Purple = '800080'; Red = 'FF0000'; RosyBrown = 'BC8F8F'; RoyalBlue = '4169E1'; SaddleBrown = '8B4513';
            Salmon = 'FA8072'; SandyBrown = 'F4A460'; SeaGreen = '2E8B57'; Seashell = 'FFF5EE'; Sienna = 'A0522D'; Silver = 'C0C0C0';
            SkyBlue = '87CEEB'; SlateBlue = '6A5ACD'; SlateGray = '708090'; Snow = 'FFFAFA'; SpringGreen = '00FF7F'; SteelBlue = '4682B4';
            Tan = 'D2B48C'; Teal = '008080'; Thistle = 'D8BFD8'; Tomato = 'FF6347'; Turquoise = '40E0D0'; Violet = 'EE82EE'; Wheat = 'F5DEB3';
            White = 'FFFFFF'; WhiteSmoke = 'F5F5F5'; Yellow = 'FFFF00'; YellowGreen = '9ACD32';
        };

    } #end begin
    process {

        $pscriboColor = $Color;
        if ($wordColorConstants.ContainsKey($pscriboColor)) {
            return $wordColorConstants[$pscriboColor].ToLower();
        }
        elseif ($pscriboColor.Length -eq 6 -or $pscriboColor.Length -eq 3) {
            $pscriboColor = '#{0}' -f $pscriboColor;
        }
        elseif ($pscriboColor.Length -eq 7 -or $pscriboColor.Length -eq 4) {
            if (-not ($pscriboColor.StartsWith('#'))) { return $null; }
        }
        if ($pscriboColor -notmatch '^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$') { return $null; }
        return $pscriboColor.TrimStart('#').ToLower();

    } #end process
} #end function ResolvePScriboColor


function Test-PScriboStyleColor {
<#
    .SYNOPSIS
        Tests whether a color string is a valid HTML color.
#>
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Color
    )
    process {

        if (Resolve-PScriboStyleColor -Color $Color) { return $true; }
        else { return $false; }

    } #end process
} #end function test-pscribostylecolor


function Test-PScriboStyle {
<#
    .SYNOPSIS
        Tests whether a style has been defined.
#>
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name
    )
    process {

        return $PScriboDocument.Styles.ContainsKey($Name);

    }
} #end function Test-PScriboStyle


function Style {
<#
    .SYNOPSIS
        Defines a new PScribo formatting style.
    .DESCRIPTION
        Creates a standard format formatting style that can be applied
        to PScribo document keywords, e.g. a combination of font style, font
        weight and font size.
    .NOTES
        Not all plugins support all options.
#>
    [CmdletBinding()]
    param (
        ## Style name
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name,

        ## Font size (pt)
        [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
        [System.UInt16] $Size = 11,

        ## Font color/colour
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('Colour')]
        [ValidateNotNullOrEmpty()]
        [System.String] $Color = '000',

        ## Background color/colour
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('BackgroundColour')]
        [ValidateNotNullOrEmpty()]
        [System.String] $BackgroundColor,

        ## Bold typeface
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Bold,

        ## Italic typeface
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Italic,

        ## Underline typeface
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Underline,

        ## Text alignment
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('Left','Center','Right','Justify')]
        [System.String] $Align = 'Left',

        ## Set as default style
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Default,

        ## Style id
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Id = $Name -Replace(' ',''),

        ## Font name (array of names for HTML output)
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String[]] $Font
    )
    begin {

        <#! Style.Internal.ps1 !#>

    }
    process {

        WriteLog -Message ($localized.ProcessingStyle -f $Id);
        Add-PScriboStyle @PSBoundParameters;

    } #end process
} #end function Style


function Set-Style {
<#
    .SYNOPSIS
        Sets the style for an individual table row or cell.
#>
    [CmdletBinding()]
    [OutputType([System.Object])]
    param (
        ## PSCustomObject to apply the style to
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Object[]] [Ref] $InputObject,

        ## PScribo style Id to apply
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.String] $Style,

        ## Property name(s) to apply the selected style to. Leave blank to apply the style to the entire row.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String[]] $Property = '',

        ## Passes the modified object back to the pipeline
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $PassThru
    ) #end param
    begin {

        if (-not (Test-PScriboStyle -Name $Style)) {
            Write-Error ($localized.UndefinedStyleError -f $Style);
            return;
        }

    }
    process {

        foreach ($object in $InputObject) {
            foreach ($p in $Property) {
                ## If $Property not set, __Style will apply to the whole row.
                $propertyName = '{0}__Style' -f $p;
                $object | Add-Member -MemberType NoteProperty -Name $propertyName -Value $Style -Force;
            }
        }
        if ($PassThru) {
            return $object;
        }

    } #end process
} #end function Set-Style

# SIG # Begin signature block
# MIIXtwYJKoZIhvcNAQcCoIIXqDCCF6QCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUIpkoAeMpKysIm/B9fj+rHoCz
# zQKgghLqMIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQU/i1x1HwsKyKBGK10R6/pF6cu2gcwDQYJKoZI
# hvcNAQEBBQAEggEATsgR1+CysooytKMVnF4y5cqvZrtYYoLxfoZnhRw0QHwjAhWh
# +sZPorlgGUhniyXH7mSER2ZyKdmGWCpAaXfi+HR5mHat3+fyFORmUa4l78poR1uT
# pge2w5wPTFcZvp1AzNhN1C+3qie5uYg4F9OrMylpYTQIPvzgoK3rULIWNBd62BFi
# q8AOVlhmdXxZOVGTGp78iPaLPWMv3dXIVbBEOzSW3mQ6MHpAgR28YmMYeZB6W28F
# tGLP76WhaguPCVI5Ra0dxQkUdFWNYnroBJ9swhvDo22QbCHIECJXufcRr0OI9vTl
# uQWNyrQdbdXWWM7JekCzRbhEOZovE8neTa1sF6GCAgswggIHBgkqhkiG9w0BCQYx
# ggH4MIIB9AIBATByMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBD
# b3Jwb3JhdGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2
# aWNlcyBDQSAtIEcyAhAOz/Q4yP6/NW4E2GqYGxpQMAkGBSsOAwIaBQCgXTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA4MDEyMTQ2
# MjFaMCMGCSqGSIb3DQEJBDEWBBSZ38Slq1fLifnsAH5XHT0Iz13f1jANBgkqhkiG
# 9w0BAQEFAASCAQAvGO05bt2o8s/9WGkpm3zmvWDqX4xh2gsK9jaqis9sm7zFwOHq
# rVAw5ssKTWd8pTmJK1C+DBi1UHopA4sCU+h8g+XrBiHi0PVEARSaMmAY4tmLtYTs
# XfnDQAgsDnshFxCIqB18Y04tY3q6klSt4LgjoEspmynPv+y1ioZ9l50gi8OgyqwJ
# 2o6KurpUrtufH8NIztodCtL46Ap9yP4wmwIGAljlWoPlYwPtc+ocMZKh23N/llqu
# x4yuIpjPtaMivG1mYfF8Hoj1sSF0i9dYEsDD6Wq/S/xkD/Bg7lWA/saJVwTuC+PG
# 9qt8y0hfRI8OgXs6R13qJLrMzb3dMSYEPma5
# SIG # End signature block
