        #region OutText Private Functions

        function New-PScriboTextOptions {
        <#
            .SYNOPSIS
                Sets the text plugin specific formatting/output options.
            .NOTES
                All plugin options should be prefixed with the plugin name.
        #>
            [CmdletBinding()]
            param (
                ## Text/output width. 0 = none/no wrap.
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNull()]
                [System.Int32] $TextWidth = 120,

                ## Document header separator character.
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateLength(1,1)]
                [System.String] $HeaderSeparator = '=',

                ## Document section separator character.
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateLength(1,1)]
                [System.String] $SectionSeparator = '-',

                ## Document section separator character.
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateLength(1,1)]
                [System.String] $LineBreakSeparator = '_',

                ## Default header/section separator width.
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNull()]
                [System.Int32] $SeparatorWidth = $TextWidth,

                ## Text encoding
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateSet('ASCII','Unicode','UTF7','UTF8')]
                [System.String] $Encoding = 'ASCII'
            )
            process {

                ## Flag that the Text options have been set. There should be a flag that is checked
                ## by the plugin if the user has set plugin-specific options.
                $options = @{
                    TextWidth = $TextWidth;
                    HeaderSeparator = $HeaderSeparator;
                    SectionSeparator = $SectionSeparator;
                    LineBreakSeparator = $LineBreakSeparator;
                    SeparatorWidth = $SeparatorWidth;
                    Encoding = $Encoding;
                }
                return $options;

            } #end process
        } #end function New-PScriboTextOptions


        function OutTextTOC {
        <#
            .SYNOPSIS
                Output formatted Table of Contents
        #>
            [CmdletBinding()]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $TOC
            )
            begin {

                ## Fix Set-StrictMode
                if (Test-Path -Path Variable:\Options) {

                	$options = Get-Variable -Name Options -ValueOnly;
                    if (-not ($options.ContainsKey('SeparatorWidth'))) {
                        $options['SeparatorWidth'] = 120;
                    }
                    if (-not ($options.ContainsKey('LineBreakSeparator'))) {
                        $options['LineBreakSeparator'] = '_';
                    }
                    if (-not ($options.ContainsKey('TextWidth'))) {
                        $options['TextWidth'] = 120;
                    }
                    if(-not ($Options.ContainsKey('SectionSeparator'))) {
                        $options['SectionSeparator'] = "-";
                    }
                }
                else { $options = New-PScriboTextOptions; }

            }
            process {

                $tocBuilder = New-Object -TypeName System.Text.StringBuilder;
                [ref] $null = $tocBuilder.AppendLine($TOC.Name);
                [ref] $null = $tocBuilder.AppendLine(''.PadRight($options.SeparatorWidth, $options.SectionSeparator));

                if ($Options.ContainsKey('EnableSectionNumbering')) {

                    $maxSectionNumberLength = ([System.String] ($Document.TOC.Number | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum)).Length;
                    foreach ($tocEntry in $Document.TOC) {
                        $sectionNumberPaddingLength = $maxSectionNumberLength - $tocEntry.Number.Length;
                        $sectionNumberIndent = ''.PadRight($tocEntry.Level, ' ');
                        $sectionPadding = ''.PadRight($sectionNumberPaddingLength, ' ');
                        [ref] $null = $tocBuilder.AppendFormat('{0}{1}  {2}{3}', $tocEntry.Number, $sectionPadding, $sectionNumberIndent, $tocEntry.Name).AppendLine();
                    } #end foreach TOC entry
                }
                else {

                    $maxSectionNumberLength = $Document.TOC.Level | Sort-Object | Select-Object -Last 1;
                    foreach ($tocEntry in $Document.TOC) {
                        $sectionNumberIndent = ''.PadRight($tocEntry.Level, ' ');
                        [ref] $null = $tocBuilder.AppendFormat('{0}{1}', $sectionNumberIndent, $tocEntry.Name).AppendLine();
                    } #end foreach TOC entry
                }

                return $tocBuilder.ToString();

            } #end process
        } #end function OutTextTOC


        function OutTextBlankLine {
        <#
            .SYNOPSIS
                Output formatted text blankline.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $BlankLine
            )
            process {

                $blankLineBuilder = New-Object -TypeName System.Text.StringBuilder;
                for ($i = 0; $i -lt $BlankLine.LineCount; $i++) {
                    [ref] $null = $blankLineBuilder.AppendLine();
                }
                return $blankLineBuilder.ToString();

            } #end process
        } #end function OutHtmlBlankLine


        function OutTextSection {
        <#
            .SYNOPSIS
                Output formatted text section.
        #>
            [CmdletBinding()]
            param (
                ## Section to output
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Section
            )
            begin {

                ## Fix Set-StrictMode
                if (Test-Path -Path Variable:\Options) {
                    $options = Get-Variable -Name Options -ValueOnly;
                    if (-not ($options.ContainsKey('SeparatorWidth'))) {
                        $options['SeparatorWidth'] = 120;
                    }
                    if (-not ($options.ContainsKey('LineBreakSeparator'))) {
                        $options['LineBreakSeparator'] = '_';
                    }
                    if (-not ($options.ContainsKey('TextWidth'))) {
                        $options['TextWidth'] = 120;
                    }
                    if(-not ($Options.ContainsKey('SectionSeparator'))) {
                        $options['SectionSeparator'] = "-";
                    }
                }
                else { $options = New-PScriboTextOptions; }

            }
            process {

                $sectionBuilder = New-Object -TypeName System.Text.StringBuilder;
                if ($Document.Options['EnableSectionNumbering']) { [string] $sectionName = '{0} {1}' -f $Section.Number, $Section.Name; }
                else { [string] $sectionName = '{0}' -f $Section.Name; }
                [ref] $null = $sectionBuilder.AppendLine();
                [ref] $null = $sectionBuilder.AppendLine($sectionName.TrimStart());
                [ref] $null = $sectionBuilder.AppendLine(''.PadRight($options.SeparatorWidth, $options.SectionSeparator));
                foreach ($s in $Section.Sections.GetEnumerator()) {
                    if ($s.Id.Length -gt 40) { $sectionId = '{0}..' -f $s.Id.Substring(0,38); }
                    else { $sectionId = $s.Id; }
                    $currentIndentationLevel = 1;
                    if ($null -ne $s.PSObject.Properties['Level']) { $currentIndentationLevel = $s.Level +1; }
                    WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
                    switch ($s.Type) {
                        'PScribo.Section' { [ref] $null = $sectionBuilder.Append((OutTextSection -Section $s)); }
                        'PScribo.Paragraph' { [ref] $null = $sectionBuilder.Append(($s | OutTextParagraph)); }
                        'PScribo.PageBreak' { [ref] $null = $sectionBuilder.AppendLine((OutTextPageBreak)); }  ## Page breaks implemented as line break with extra padding
                        'PScribo.LineBreak' { [ref] $null = $sectionBuilder.AppendLine((OutTextLineBreak)); }
                        'PScribo.Table' { [ref] $null = $sectionBuilder.AppendLine(($s | OutTextTable)); }
                        'PScribo.BlankLine' { [ref] $null = $sectionBuilder.AppendLine(($s | OutTextBlankLine)); }
                        Default { WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning; }
                    } #end switch
                } #end foreach
                return $sectionBuilder.ToString();

            } #end process
        } #end function outtextsection


        function OutTextParagraph {
        <#
            .SYNOPSIS
                Output formatted paragraph text.
        #>
            [CmdletBinding()]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Paragraph
            )
            begin {

                ## Fix Set-StrictMode
                if (Test-Path -Path Variable:\Options) {
                    $options = Get-Variable -Name Options -ValueOnly;
                    if (-not ($options.ContainsKey('TextWidth'))) {
                        $options['TextWidth'] = 120;
                    }
                }
                else { $options = New-PScriboTextOptions; }

            }
            process {

                $padding = ''.PadRight(($Paragraph.Tabs * 4), ' ');
                if ([string]::IsNullOrEmpty($Paragraph.Text)) { $text = "$padding$($Paragraph.Id)"; }
                else { $text = "$padding$($Paragraph.Text)"; }

                $formattedText = OutStringWrap -InputObject $text -Width $Options.TextWidth;

                if ($Paragraph.NewLine) { return "$formattedText`r`n"; }
                else { return $formattedText; }

            } #end process
        } #end outtextparagraph


        function OutTextLineBreak {
        <#
            .SYNOPSIS
                Output formatted line break text.
        #>
            [CmdletBinding()]
            param ( )
            begin {

                ## Fix Set-StrictMode
                if (Test-Path -Path Variable:\Options) { $options = Get-Variable -Name Options -ValueOnly; }
                else { $options = New-PScriboTextOptions; }

                if (-not ($options.ContainsKey('SeparatorWidth'))) {
                    $options['SeparatorWidth'] = 120;
                }
                if (-not ($options.ContainsKey('LineBreakSeparator'))) {
                    $options['LineBreakSeparator'] = '_';
                }
                if (-not ($options.ContainsKey('TextWidth'))) {
                    $options['TextWidth'] = 120;
                }

            }
            process {

                ## Use the specified output width
                if ($options.TextWidth -eq 0) { $options.TextWidth = $Host.UI.RawUI.BufferSize.Width -1; }
                $lb = ''.PadRight($options.SeparatorWidth, $options.LineBreakSeparator);
                return "$(OutStringWrap -InputObject $lb -Width $options.TextWidth)`r`n";

            } #end process
        } #end function OutTextLineBreak


        function OutTextPageBreak {
        <#
            .SYNOPSIS
                Output formatted line break text.
        #>
            [CmdletBinding()]
            param ( )
            process {
                return "$(OutTextLineBreak)`r`n";
            } #end process
        } #end function OutTextLineBreak


        function OutTextTable {
        <#
            .SYNOPSIS
                Output formatted text table.
        #>
            [CmdletBinding()]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Table
            )
            begin {

                ## Fix Set-StrictMode
                if (Test-Path -Path Variable:\Options) { $options = Get-Variable -Name Options -ValueOnly; }
                else { $options = New-PScriboTextOptions; }

            }
            process {

                ## Use the specified output width
                if ($options.TextWidth -eq 0) { $options.TextWidth = $Host.UI.RawUI.BufferSize.Width -1; }
                if ($Table.List) {
                    $text = ($Table.Rows | Select-Object -Property * -ExcludeProperty '*__Style' | Format-List | Out-String -Width $options.TextWidth).Trim();
                } else {
                    ## Don't trim tabs for table headers
                    ## Tables set to AutoSize as otherwise, rendering is different between PoSh v4 and v5
                    $text = ($Table.Rows | Select-Object -Property * -ExcludeProperty '*__Style' | Format-Table -Wrap -AutoSize | Out-String -Width $options.TextWidth).Trim("`r`n");
                }
                # Ensure there's a space before and after the table.
                return "`r`n$text`r`n";

            } #end process
        } #end function outtexttable


        function OutStringWrap {
        <#
            .SYNOPSIS
                Outputs objects to strings, wrapping as required.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [Object[]] $InputObject,

                [Parameter()]
                [ValidateNotNull()]
                [System.Int32] $Width = $Host.UI.RawUI.BufferSize.Width
            )
            begin {

                ## 2 is the minimum, therefore default to wiiiiiiiiiide!
                if ($Width -lt 2) { $Width = 4096; }
                WriteLog -Message ('Wrapping text at "{0}" characters.' -f $Width) -IsDebug;

            }
            process {

                foreach ($object in $InputObject) {
                    $textBuilder = New-Object -TypeName System.Text.StringBuilder;
                    $text = (Out-String -InputObject $object).TrimEnd("`r`n");
                    for ($i = 0; $i -le $text.Length; $i += $Width) {
                        if (($i + $Width) -ge ($text.Length -1)) { [ref] $null = $textBuilder.Append($text.Substring($i)); }
                        else { [ref] $null = $textBuilder.AppendLine($text.Substring($i, $Width)); }
                    } #end for
                    return $textBuilder.ToString();
                    $textBuilder = $null;
                } #end foreach

            } #end process
        } #end function OutStringWrap

        #endregion OutText Private Functions

# SIG # Begin signature block
# MIIXtwYJKoZIhvcNAQcCoIIXqDCCF6QCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUXsGLX0cOpGXmg/BcuLoVNOl9
# w3CgghLqMIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUr7iedlwZu73OaojRMYUoxuSZv+0wDQYJKoZI
# hvcNAQEBBQAEggEACXphGKZ0r3j+d/1FpyHkYjB+JgRA8zzmvFCg/lZZVzaJfaDy
# 8zsE+KNTIC1RNUF07jLefj31hxuvmjD5vduJ4V/hRwhlwk54I397PBmdp7TInQ8i
# 9/FVq4ezIYxR6n65PjeGccNNH8A9daUXxQpkTh/SIEx+rz6/mDl5C+UzBreHGK3S
# pQgptncdDAt3hUKL91fAwKmxVFgqB918xO/6TiFbJqa/M9XsgcyKAumwMEI7bM8v
# kcczuEZ0lHhHIs0kaUoTIU4CSJXDKxc84CIEoKMod/NK+tvMWSyTAx7nUkW++WLx
# nPkN0Oif3rXCVkL/sjXwvC1leUm6dBRZryM+zKGCAgswggIHBgkqhkiG9w0BCQYx
# ggH4MIIB9AIBATByMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBD
# b3Jwb3JhdGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2
# aWNlcyBDQSAtIEcyAhAOz/Q4yP6/NW4E2GqYGxpQMAkGBSsOAwIaBQCgXTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA4MDEyMTQ2
# MTVaMCMGCSqGSIb3DQEJBDEWBBTkO1jFZS7DW0U3qwEe3gHXb0Z8oTANBgkqhkiG
# 9w0BAQEFAASCAQB8p660Yytyf09slFqtN269EdBiuCIx1tGUNl7+UHRn0eKPlrg0
# rc37aXGoLj8xx9n4imhGyQNAiQ9AYdAUqZjxSTCp98wYtB31MaMS1oO7W6Eg0tMZ
# UHysXQuF4oAENsOojI6wh3nkcZ7NXYm7AKrzj4pIDmFF06RY7e7lMSiCv+Y8d2Zj
# DE8T7/2PKk0xFXOm1ncaxe/Ybe5jY9v8O+znTSlp9Smxff1garpDJcViXQb1cn7G
# 6KN5BswyArsAXWC+zh0M0+UGEfVA12b1RmQi8qTAl8pFFbZclI7KU0CFpgGtLvFV
# OOlz4KS5LPBIvYpm2GGR4GPclcywX2g74t+S
# SIG # End signature block
