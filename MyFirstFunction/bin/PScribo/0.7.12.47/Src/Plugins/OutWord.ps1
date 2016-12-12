function OutWord {
<#
    .SYNOPSIS
        Microsoft Word output plugin for PScribo.
    .DESCRIPTION
        Outputs a Word document representation of a PScribo document object.
#>
    [CmdletBinding()]
    [OutputType([System.IO.FileInfo])]
    param (
        ## ThePScribo document object to convert to a text document
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Object] $Document,

        ## Output directory path for the .txt file
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.String] $Path,

        ### Hashtable of all plugin supported options
        [Parameter()]
        [AllowNull()]
        [System.Collections.Hashtable] $Options
    )
    begin {

        $pluginName = 'Word';

        <#! OutWord.Internal.ps1 !#>

    }
    process {

        $stopwatch = [Diagnostics.Stopwatch]::StartNew();
        WriteLog -Message ($localized.DocumentProcessingStarted -f $Document.Name);
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        $xmlDocument = New-Object -TypeName 'System.Xml.XmlDocument';
        [ref] $null = $xmlDocument.AppendChild($xmlDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'));
        $documentXml = $xmlDocument.AppendChild($xmlDocument.CreateElement('w', 'document', $xmlnsMain));
        [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:xml', 'http://www.w3.org/XML/1998/namespace');
        $body = $documentXml.AppendChild($xmlDocument.CreateElement('w', 'body', $xmlnsMain));
        ## Setup the document page size/margins
        $sectionPrParams = @{
            PageHeight = $Document.Options['PageHeight']; PageWidth = $Document.Options['PageWidth'];
            PageMarginTop = $Document.Options['MarginTop']; PageMarginBottom = $Document.Options['MarginBottom'];
            PageMarginLeft = $Document.Options['MarginLeft']; PageMarginRight = $Document.Options['MarginRight'];
        }
        [ref] $null = $body.AppendChild((GetWordSectionPr @sectionPrParams -XmlDocument $xmlDocument));
        foreach ($s in $Document.Sections.GetEnumerator()) {
            if ($s.Id.Length -gt 40) { $sectionId = '{0}[..]' -f $s.Id.Substring(0,36); }
            else { $sectionId = $s.Id; }
            $currentIndentationLevel = 1;
            if ($null -ne $s.PSObject.Properties['Level']) { $currentIndentationLevel = $s.Level +1; }
            WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
            switch ($s.Type) {
                'PScribo.Section' { $s | OutWordSection -RootElement $body -XmlDocument $xmlDocument; }
                'PScribo.Paragraph' { [ref] $null = $body.AppendChild((OutWordParagraph -Paragraph $s -XmlDocument $xmlDocument)); }
                'PScribo.PageBreak' { [ref] $null = $body.AppendChild((OutWordPageBreak -PageBreak $s -XmlDocument $xmlDocument)); }
                'PScribo.LineBreak' { [ref] $null = $body.AppendChild((OutWordLineBreak -LineBreak $s -XmlDocument $xmlDocument)); }
                'PScribo.Table' { OutWordTable -Table $s -XmlDocument $xmlDocument -Element $body; }
                'PScribo.TOC' { [ref] $null = $body.AppendChild((OutWordTOC -TOC $s -XmlDocument $xmlDocument)); }
                'PScribo.BlankLine' { OutWordBlankLine -BlankLine $s -XmlDocument $xmlDocument -Element $body; }
                Default { WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning; }
            } #end switch
        } #end foreach
        ## Generate the Word 'styles.xml' document part
        $stylesXml = OutWordStylesDocument -Styles $Document.Styles -TableStyles $Document.TableStyles;
        ## Generate the Word 'settings.xml' document part
        if (($Document.Properties['TOCs']) -and ($Document.Properties['TOCs'] -gt 0)) {
            ## We have a TOC so flag to update the document when opened
            $settingsXml = OutWordSettingsDocument -UpdateFields;
        }
        else {
            $settingsXml = OutWordSettingsDocument;
        }
        #Convert relative or PSDrive based path to the absolute filesystem path
        $AbsolutePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        $destinationPath = Join-Path -Path $AbsolutePath ('{0}.docx' -f $Document.Name);
        Add-Type -AssemblyName WindowsBase;
        try {
            $package = [System.IO.Packaging.Package]::Open($destinationPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite);
        }
        catch {
            WriteLog -Message ($localized.OpenPackageError -f $destinationPath) -IsWarning;
            throw $_;
        }
        ## Create document.xml part
        $documentUri = New-Object System.Uri('/word/document.xml', [System.UriKind]::Relative);
        WriteLog -Message ($localized.ProcessingDocumentPart -f $documentUri);
        $documentPart = $package.CreatePart($documentUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml');
        $streamWriter = New-Object System.IO.StreamWriter($documentPart.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite));
        $xmlWriter = [System.Xml.XmlWriter]::Create($streamWriter);
        WriteLog -Message ($localized.WritingDocumentPart -f $documentUri);
        $xmlDocument.Save($xmlWriter);
        $xmlWriter.Close();
        $streamWriter.Close();

        ## Create styles.xml part
        $stylesUri = New-Object System.Uri('/word/styles.xml', [System.UriKind]::Relative);
        WriteLog -Message ($localized.ProcessingDocumentPart -f $stylesUri);
        $stylesPart = $package.CreatePart($stylesUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml');
        $streamWriter = New-Object System.IO.StreamWriter($stylesPart.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite));
        $xmlWriter = [System.Xml.XmlWriter]::Create($streamWriter);
        WriteLog -Message ($localized.WritingDocumentPart -f $stylesUri);
        $stylesXml.Save($xmlWriter);
        $xmlWriter.Close();
        $streamWriter.Close();

        ## Create settings.xml part
        $settingsUri = New-Object System.Uri('/word/settings.xml', [System.UriKind]::Relative);
        WriteLog -Message ($localized.ProcessingDocumentPart -f $settingsUri);
        $settingsPart = $package.CreatePart($settingsUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
        $streamWriter = New-Object System.IO.StreamWriter($settingsPart.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite));
        $xmlWriter = [System.Xml.XmlWriter]::Create($streamWriter);
        WriteLog -Message ($localized.WritingDocumentPart -f $settingsUri);
        $settingsXml.Save($xmlWriter);
        $xmlWriter.Close();
        $streamWriter.Close();

        ## Create the Package relationships
        WriteLog -Message $localized.GeneratingPackageRelationships;
        [ref] $null = $package.CreateRelationship($documentUri, [System.IO.Packaging.TargetMode]::Internal, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'rId1');
        [ref] $null = $documentPart.CreateRelationship($stylesUri, [System.IO.Packaging.TargetMode]::Internal, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'rId1');
        [ref] $null = $documentPart.CreateRelationship($settingsUri, [System.IO.Packaging.TargetMode]::Internal, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'rId2');
        WriteLog -Message ($localized.SavingFile -f $destinationPath);
        $package.Flush();
        $package.Close();

        $stopwatch.Stop();
        WriteLog -Message ($localized.DocumentProcessingCompleted -f $Document.Name);
        WriteLog -Message ($localized.TotalProcessingTime -f $stopwatch.Elapsed.TotalSeconds);
        ## Return the file reference to the pipeline
        Write-Output (Get-Item -Path $destinationPath);

    } #end process
} #end function OutWord

# SIG # Begin signature block
# MIIXtwYJKoZIhvcNAQcCoIIXqDCCF6QCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU9pM9C4XXn4xPJ9hiHx+iPffk
# ZP+gghLqMIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUdqzd8CwJlcgQhWCTlitEpnOrqVkwDQYJKoZI
# hvcNAQEBBQAEggEAlHokbMUNdpxyM6OYNgimQIP+cYZu+1ZGQOaCbj8BxW/1WisA
# zOpH6FJi0kKD7xo70wkVl2jBQepJGv8ScYp5d1sBcMaX6tfum5KwHo9OeEDoV++s
# HTOy5I12LbfwToEQTYNVaSDyoA7/s/gRDiTchERxyoxRAyaKTrQti95zgBrUPQ4X
# kK31yy2m6Smbj39c52/xjRjEYtcuGFySEYP1hwwwW4B+0eYfvy3SmWogRRUmJWRg
# yPLdMer/gv8Ao8iIMmneqSbnJ3bJPCank9Xn847GcQlL4KaUspQ0sCFqNtkzgtJN
# 8OZsWBbI/biUm35dq6Rcgw0legSvjnW1aui0RqGCAgswggIHBgkqhkiG9w0BCQYx
# ggH4MIIB9AIBATByMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBD
# b3Jwb3JhdGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2
# aWNlcyBDQSAtIEcyAhAOz/Q4yP6/NW4E2GqYGxpQMAkGBSsOAwIaBQCgXTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA4MDEyMTQ2
# MTVaMCMGCSqGSIb3DQEJBDEWBBS6L0yy7sBXMIMfQx9MVL1Yn0MqejANBgkqhkiG
# 9w0BAQEFAASCAQAjWWCUAWUvWXIr1P30+DoaxFgha6kIHRqJ1PBf+TTjrLe9AAuP
# VlwPW4jvs9/QI25aVPsxJSjtUJO4LHK7FEk1AHAtx6yY41q3bGVMIezTprzIRvt1
# 2FT1H06iERomOjvvCA3VpLv1046us68yN1nsON+BOZZ5jhBMfOhDM6W7B2S+1G9r
# F+s9z4qkDUGt90a/MqGYYIL/hnwVEUOKeWKunHqIronyAcmqNeU8Uzk4ldOFLR7i
# uE/GNWC2yIuD6dYgWzsAoRShV2S/KdPLWJUg17DOz+Ll3Uh/x4QWgjJEt5F2uf7Y
# 1amVeAIJc+ZJ8cVIV8Wn0cnQBeMJUmuOvZu7
# SIG # End signature block
