        #region OutHtml Private Functions

        function GetHtmlStyle {
        <#
            .SYNOPSIS
                Generates html stylesheet style attributes from a PScribo document style.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                ## PScribo document style
                [Parameter(Mandatory, ValueFromPipeline)] [System.Object] $Style
            )
            process {

                $styleBuilder = New-Object -TypeName System.Text.StringBuilder;
                [ref] $null = $styleBuilder.AppendFormat(" font-family: '{0}';", $Style.Font -Join "','");
                ## Create culture invariant decimal https://github.com/iainbrighton/PScribo/issues/6
                $invariantFontSize =  ($Style.Size / 12).ToString('f2', [System.Globalization.CultureInfo]::InvariantCulture);
                [ref] $null = $styleBuilder.AppendFormat(' font-size: {0}em;', $invariantFontSize);
                [ref] $null = $styleBuilder.AppendFormat(' font-size: {0:0.00}em;', $Style.Size / 12);
                [ref] $null = $styleBuilder.AppendFormat(' text-align: {0};', $Style.Align.ToLower());
                if ($Style.Bold) { [ref] $null = $styleBuilder.Append(' font-weight: bold;'); }
                else { [ref] $null = $styleBuilder.Append(' font-weight: normal;'); }
                if ($Style.Italic) { [ref] $null = $styleBuilder.Append(' font-style: italic;'); }
                if ($Style.Underline) { [ref] $null = $styleBuilder.Append(' text-decoration: underline;'); }
                if ($Style.Color.StartsWith('#')) { [ref] $null = $styleBuilder.AppendFormat(' color: {0};', $Style.Color.ToLower()); }
                else { [ref] $null = $styleBuilder.AppendFormat(' color: #{0};', $Style.Color); }
                if ($Style.BackgroundColor) {
                    if ($Style.BackgroundColor.StartsWith('#')) { [ref] $null = $styleBuilder.AppendFormat(' background-color: {0};', $Style.BackgroundColor.ToLower()); }
                    else { [ref] $null = $styleBuilder.AppendFormat(' background-color: #{0};', $Style.BackgroundColor.ToLower()); }
                }
                return $styleBuilder.ToString();

            }
        } #end function GetHtmlStyle


        function GetHtmlTableStyle {
        <#
            .SYNOPSIS
                Generates html stylesheet style attributes from a PScribo document table style.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                ## PScribo document table style
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $TableStyle
            )
            process {

                $tableStyleBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                [ref] $null = $tableStyleBuilder.AppendFormat(' padding: {0}em {1}em {2}em {3}em;',
                                                                (ConvertMmToEm $TableStyle.PaddingTop),
                                                                    (ConvertMmToEm $TableStyle.PaddingRight),
                                                                        (ConvertMmToEm $TableStyle.PaddingBottom),
                                                                            (ConvertMmToEm $TableStyle.PaddingLeft));
                [ref] $null = $tableStyleBuilder.AppendFormat(' border-style: {0};', $TableStyle.BorderStyle.ToLower());
                if ($TableStyle.BorderWidth -gt 0) {
                    [ref] $null = $tableStyleBuilder.AppendFormat(' border-width: {0}em;', (ConvertMmToEm $TableStyle.BorderWidth));
                    if ($TableStyle.BorderColor.Contains('#')) {
                        [ref] $null = $tableStyleBuilder.AppendFormat(' border-color: {0};', $TableStyle.BorderColor);
                    }
                    else {
                        [ref] $null = $tableStyleBuilder.AppendFormat(' border-color: #{0};', $TableStyle.BorderColor);
                    }
                }
                [ref] $null = $tableStyleBuilder.Append(' border-collapse: collapse;');
                ## <table align="center"> is deprecated in Html5
                if ($TableStyle.Align -eq 'Center') {
                    [ref] $null = $tableStyleBuilder.Append(' margin-left: auto; margin-right: auto;');
                }
                elseif ($TableStyle.Align -eq 'Right') {
                    [ref] $null = $tableStyleBuilder.Append(' margin-left: auto; margin-right: 0;');
                }
                return $tableStyleBuilder.ToString();

            }
        } #end function outhtmltablestyle


        function GetHtmlTableDiv {
        <#
            .SYNOPSIS
                Generates Html <div style=..><table style=..> tags based upon table width, columns and indentation
            .NOTES
                A <div> is required to ensure that the table stays within the "page" boundaries/margins.
        #>
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table
            )
            process {

                $divBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                if ($Table.Tabs -gt 0) {
                    [ref] $null = $divBuilder.AppendFormat('<div style="margin-left: {0}em;">' -f (ConvertMmToEm -Millimeter (12.7 * $Table.Tabs)));
                }
                else {
                    [ref] $null = $divBuilder.Append('<div>' -f (ConvertMmToEm -Millimeter (12.7 * $Table.Tabs)));
                }
                if ($Table.List) {
                    [ref] $null = $divBuilder.AppendFormat('<table class="{0}-list"', $Table.Style.ToLower());
                }
                else {
                    [ref] $null = $divBuilder.AppendFormat('<table class="{0}"', $Table.Style.ToLower());
                }
                $styleElements = @();
                if ($Table.Width -gt 0) {
                    $styleElements += 'width:{0}%;' -f $Table.Width;
                }
                if ($Table.ColumnWidths) {
                    $styleElements += 'table-layout: fixed;';
                    $styleElements += 'word-break: break-word;'
                }
                if ($styleElements.Count -gt 0) {
                    [ref] $null = $divBuilder.AppendFormat(' style="{0}">', [String]::Join(' ', $styleElements));
                }
                else {
                    [ref] $null = $divBuilder.Append('>');
                }
                return $divBuilder.ToString();

            }
        } #end function GetHtmlTableDiv


        function GetHtmlTableColGroup {
        <#
            .SYNOPSIS
                Generates Html <colgroup> tags based on table column widths
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table
            )
            process {

                $colGroupBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                if ($Table.ColumnWidths) {
                    [ref] $null = $colGroupBuilder.Append('<colgroup>');
                    foreach ($columnWidth in $Table.ColumnWidths) {
                        if ($null -eq $columnWidth) {
                            [ref] $null = $colGroupBuilder.Append('<col />');
                        }
                        else {
                            [ref] $null = $colGroupBuilder.AppendFormat('<col style="max-width:{0}%; min-width:{0}%; width:{0}%" />', $columnWidth);
                        }
                    }
                    [ref] $null = $colGroupBuilder.AppendLine('</colgroup>');
                }
                return $colGroupBuilder.ToString();

            }
        } #end function GetHtmlTableDiv


        function OutHtmlTOC {
        <#
            .SYNOPSIS
                Generates Html table of contents.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $TOC
            )
            process {

                $tocBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                [ref] $null = $tocBuilder.AppendFormat('<h1 class="TOC">{0}</h1>', $TOC.Name);
                #[ref] $null = $tocBuilder.AppendLine('<table style="width: 100%;">');
                [ref] $null = $tocBuilder.AppendLine('<table>');
                foreach ($tocEntry in $Document.TOC) {
                    $sectionNumberIndent = '&nbsp;&nbsp;&nbsp;' * $tocEntry.Level;
                    if ($Document.Options['EnableSectionNumbering']) {
                        [ref] $null = $tocBuilder.AppendFormat('<tr><td>{0}</td><td>{1}<a href="#{2}" style="text-decoration: none;">{3}</a></td></tr>', $tocEntry.Number, $sectionNumberIndent, $tocEntry.Id, $tocEntry.Name).AppendLine();
                    }
                    else {
                        [ref] $null = $tocBuilder.AppendFormat('<tr><td>{0}<a href="#{1}" style="text-decoration: none;">{2}</a></td></tr>', $sectionNumberIndent, $tocEntry.Id, $tocEntry.Name).AppendLine();
                    }
                }
                [ref] $null = $tocBuilder.AppendLine('</table>');
                return $tocBuilder.ToString();

            } #end process
        } #end function OutHtmlTOC


        function OutHtmlBlankLine {
        <#
            .SYNOPSIS
                Outputs html PScribo.Blankline.
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
                    [ref] $null = $blankLineBuilder.Append('<br />');
                }
                return $blankLineBuilder.ToString();

            } #end process
        } #end function OutHtmlBlankLine


        function OutHtmlStyle {
        <#
            .SYNOPSIS
                Generates an in-line HTML CSS stylesheet from a PScribo document styles and table styles.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                ## PScribo document styles
                [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
                [System.Collections.Hashtable] $Styles,

                ## PScribo document tables styles
                [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
                [System.Collections.Hashtable] $TableStyles,

                ## Suppress page layout styling
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $NoPageLayoutStyle
            )
            process {

                $stylesBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                [ref] $null = $stylesBuilder.AppendLine('<style type="text/css">');
                if (-not $NoPageLayoutStyle) {
                    ## Add HTML page layout styling options, e.g. when emailing HTML documents
                    [ref] $null = $stylesBuilder.AppendLine('html { height: 100%; -webkit-background-size: cover; -moz-background-size: cover; -o-background-size: cover; background-size: cover; background: #f8f8f8; }');
                    [ref] $null = $stylesBuilder.Append("page { background: white; width: $($Document.Options['PageWidth'])mm; display: block; margin-top: 1em; margin-left: auto; margin-right: auto; margin-bottom: 1em; ");
                    [ref] $null = $stylesBuilder.AppendLine('border-style: solid; border-width: 1px; border-color: #c6c6c6; }');
                    [ref] $null = $stylesBuilder.AppendLine('@media print { body, page { margin: 0; box-shadow: 0; } }');
                    [ref] $null = $stylesBuilder.AppendLine('hr { margin-top: 1.0em; }');
                }
                foreach ($style in $Styles.Keys) {
                    ## Build style
                    $htmlStyle = GetHtmlStyle -Style $Styles[$style];
                    [ref] $null = $stylesBuilder.AppendFormat(' .{0} {{{1} }}', $Styles[$style].Id, $htmlStyle).AppendLine();
                }
                foreach ($tableStyle in $TableStyles.Keys) {
                    $tStyle = $TableStyles[$tableStyle];
                    $tableStyleId = $tStyle.Id.ToLower();
                    $htmlTableStyle = GetHtmlTableStyle -TableStyle $tStyle;
                    $htmlHeaderStyle = GetHtmlStyle -Style $Styles[$tStyle.HeaderStyle];
                    $htmlRowStyle = GetHtmlStyle -Style $Styles[$tStyle.RowStyle];
                    $htmlAlternateRowStyle = GetHtmlStyle -Style $Styles[$tStyle.AlternateRowStyle];
                    ## Generate Standard table styles
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0} {{{1} }}', $tableStyleId, $htmlTableStyle).AppendLine();
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0} th {{{1}{2} }}', $tableStyleId, $htmlHeaderStyle, $htmlTableStyle).AppendLine();
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0} tr:nth-child(odd) td {{{1}{2} }}', $tableStyleId, $htmlRowStyle, $htmlTableStyle).AppendLine();
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0} tr:nth-child(even) td {{{1}{2} }}', $tableStyleId, $htmlAlternateRowStyle, $htmlTableStyle).AppendLine();
                    ## Generate List table styles
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0}-list {{{1} }}', $tableStyleId, $htmlTableStyle).AppendLine();
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0}-list td:nth-child(1) {{{1}{2} }}', $tableStyleId, $htmlHeaderStyle, $htmlTableStyle).AppendLine();
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0}-list td:nth-child(2) {{{1}{2} }}', $tableStyleId, $htmlRowStyle, $htmlTableStyle).AppendLine();
                } #end foreach style
                [ref] $null = $stylesBuilder.AppendLine('</style>');
                return $stylesBuilder.ToString().TrimEnd();

            } #end process
        } #end function OutHtmlStyle


        function OutHtmlSection {
        <#
            .SYNOPSIS
                Output formatted Html section.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                ## Section to output
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Section
            )
            process {

                [System.Text.StringBuilder] $sectionBuilder = New-Object System.Text.StringBuilder;
                $encodedSectionName = [System.Net.WebUtility]::HtmlEncode($Section.Name);
                if ($Document.Options['EnableSectionNumbering']) { [string] $sectionName = '{0} {1}' -f $Section.Number, $encodedSectionName; }
                else { [string] $sectionName = '{0}' -f $encodedSectionName; }
                [int] $headerLevel = $Section.Number.Split('.').Count;
                ## Html <h5> is the maximum supported level
                if ($headerLevel -ge 5) {
                    WriteLog -Message $localized.MaxHeadingLevelWarning -IsWarning;
                    $headerLevel = 5;
                }
                if ([string]::IsNullOrEmpty($Section.Style)) { $className = $Document.DefaultStyle; }
                else { $className = $Section.Style; }
                [ref] $null = $sectionBuilder.AppendFormat('<a name="{0}"><h{1} class="{2}">{3}</h{1}></a>', $Section.Id, $headerLevel, $className, $sectionName.TrimStart());
                foreach ($s in $Section.Sections.GetEnumerator()) {
                    if ($s.Id.Length -gt 40) { $sectionId = '{0}[..]' -f $s.Id.Substring(0,36); }
                    else { $sectionId = $s.Id; }
                    $currentIndentationLevel = 1;
                    if ($null -ne $s.PSObject.Properties['Level']) { $currentIndentationLevel = $s.Level +1; }
                    WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
                    switch ($s.Type) {
                        'PScribo.Section' { [ref] $null = $sectionBuilder.Append((OutHtmlSection -Section $s)); }
                        'PScribo.Paragraph' { [ref] $null = $sectionBuilder.Append((OutHtmlParagraph -Paragraph $s)); }
                        'PScribo.LineBreak' { [ref] $null = $sectionBuilder.Append((OutHtmlLineBreak)); }
                        'PScribo.PageBreak' { [ref] $null = $sectionBuilder.Append((OutHtmlPageBreak)); }
                        'PScribo.Table' { [ref] $null = $sectionBuilder.Append((OutHtmlTable -Table $s)); }
                        'PScribo.BlankLine' { [ref] $null = $sectionBuilder.Append((OutHtmlBlankLine -BlankLine $s)); }
                        Default { WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning; }
                    } #end switch
                } #end foreach
                return $sectionBuilder.ToString();

            } #end process
        } # end function OutHtmlSection


        function GetHtmlParagraphStyle {
        <#
            .SYNOPSIS
                Generates html style attribute from PScribo paragraph style overrides.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()] [System.Object] $Paragraph
            )
            process {

                $paragraphStyleBuilder = New-Object -TypeName System.Text.StringBuilder;
                if ($Paragraph.Tabs -gt 0) {
                    ## Default to 1/2in tab spacing
                    $tabEm = ConvertMmToEm -Millimeter (12.7 * $Paragraph.Tabs);
                    [ref] $null = $paragraphStyleBuilder.AppendFormat(' margin-left: {0}em;', $tabEm);
                }
                if ($Paragraph.Font) { [ref] $null = $paragraphStyleBuilder.AppendFormat(" font-family: '{0}';", $Paragraph.Font -Join "','"); }
                if ($Paragraph.Size -gt 0) {
                    ## Create culture invariant decimal https://github.com/iainbrighton/PScribo/issues/6
                    $invariantParagraphSize = ($Paragraph.Size / 12).ToString('f2', [System.Globalization.CultureInfo]::InvariantCulture);
                    [ref] $null = $paragraphStyleBuilder.AppendFormat(' font-size: {0}em;', $invariantParagraphSize);
                }
                if ($Paragraph.Bold -eq $true) { [ref] $null = $paragraphStyleBuilder.Append(' font-weight: bold;'); }
                if ($Paragraph.Italic -eq $true) { [ref] $null = $paragraphStyleBuilder.Append(' font-style: italic;'); }
                if ($Paragraph.Underline -eq $true) { [ref] $null = $paragraphStyleBuilder.Append(' text-decoration: underline;'); }
                if (-not [System.String]::IsNullOrEmpty($Paragraph.Color) -and $Paragraph.Color.StartsWith('#')) {
                    [ref] $null = $paragraphStyleBuilder.AppendFormat(' color: {0};', $Paragraph.Color.ToLower());
                }
                elseif (-not [System.String]::IsNullOrEmpty($Paragraph.Color)) {
                    [ref] $null = $paragraphStyleBuilder.AppendFormat(' color: #{0};', $Paragraph.Color.ToLower());
                }
                return $paragraphStyleBuilder.ToString().TrimStart();

            } #end process
        } #end function GetHtmlParagraphStyle


        function OutHtmlParagraph {
        <#
            .SYNOPSIS
                Output formatted Html paragraph.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Paragraph
            )
            process {

                [System.Text.StringBuilder] $paragraphBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                $text = [System.Net.WebUtility]::HtmlEncode($Paragraph.Text);
                if ([System.String]::IsNullOrEmpty($text)) {
                    $text = [System.Net.WebUtility]::HtmlEncode($Paragraph.Id);
                }
                $customStyle = GetHtmlParagraphStyle -Paragraph $Paragraph;
                if ([System.String]::IsNullOrEmpty($Paragraph.Style) -and [System.String]::IsNullOrEmpty($customStyle)) {
                    [ref] $null = $paragraphBuilder.AppendFormat('<div>{0}</div>', $text);
                }
                elseif ([System.String]::IsNullOrEmpty($customStyle)) {
                    [ref] $null = $paragraphBuilder.AppendFormat('<div class="{0}">{1}</div>', $Paragraph.Style, $text);
                }
                else {
                    [ref] $null = $paragraphBuilder.AppendFormat('<div style="{1}">{2}</div>', $Paragraph.Style, $customStyle, $text);
                }
                return $paragraphBuilder.ToString();

            } #end process
        } #end OutHtmlParagraph


        function GetHtmlTableList {
        <#
            .SYNOPSIS
                Generates list html <table> from a PScribo.Table row object.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table,

                [Parameter(Mandatory)]
                [System.Object] $Row
            )
            process {

                $listTableBuilder = New-Object -TypeName System.Text.StringBuilder;
                [ref] $null = $listTableBuilder.Append((GetHtmlTableDiv -Table $Table));
                [ref] $null = $listTableBuilder.Append((GetHtmlTableColGroup -Table $Table));
                [ref] $null = $listTableBuilder.Append('<tbody>');

                for ($i = 0; $i -lt $Table.Columns.Count; $i++) {
                    $propertyName = $Table.Columns[$i];
                    [ref] $null = $listTableBuilder.AppendFormat('<tr><td>{0}</td>', $propertyName);
                    $propertyStyle = '{0}__Style' -f $propertyName;

                    if ($row.PSObject.Properties[$propertyStyle]) {
                        $propertyStyleHtml = (GetHtmlStyle -Style $Document.Styles[$Row.$propertyStyle]);
                        if ([string]::IsNullOrEmpty($Row.$propertyName)) {
                            [ref] $null = $listTableBuilder.AppendFormat('<td style="{0}">&nbsp;</td></tr>', $propertyStyleHtml);
                        }
                        else {
                            [ref] $null = $listTableBuilder.AppendFormat('<td style="{0}">{1}</td></tr>', $propertyStyleHtml, $Row.($propertyName));
                        }
                    }
                    else {
                        if ([string]::IsNullOrEmpty($Row.$propertyName)) {
                            [ref] $null = $listTableBuilder.Append('<td>&nbsp;</td></tr>');
                        }
                        else {
                            [ref] $null = $listTableBuilder.AppendFormat('<td>{0}</td></tr>', $Row.$propertyName);
                        }
                    }
                } #end for each property
                [ref] $null = $listTableBuilder.AppendLine('</tbody></table></div>');
                return $listTableBuilder.ToString();

            } #end process
        } #end function GetHtmlTableList


        function GetHtmlTable {
        <#
            .SYNOPSIS
                Generates html <table> from a PScribo.Table object.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table
            )
            process {

                $standardTableBuilder = New-Object -TypeName System.Text.StringBuilder;
                [ref] $null = $standardTableBuilder.Append((GetHtmlTableDiv -Table $Table));
                [ref] $null = $standardTableBuilder.Append((GetHtmlTableColGroup -Table $Table));

                ## Table headers
                [ref] $null = $standardTableBuilder.Append('<thead><tr>');
                for ($i = 0; $i -lt $Table.Columns.Count; $i++) {
                    [ref] $null = $standardTableBuilder.AppendFormat('<th>{0}</th>', $Table.Columns[$i]);
                }
                [ref] $null = $standardTableBuilder.Append('</tr></thead>');

                ## Table body
                [ref] $null = $standardTableBuilder.AppendLine('<tbody>');
                foreach ($row in $Table.Rows) {
                    [ref] $null = $standardTableBuilder.Append('<tr>');
                    foreach ($propertyName in $Table.Columns) {
                        $propertyStyle = '{0}__Style' -f $propertyName;
                        $encodedHtmlContent = [System.Net.WebUtility]::HtmlEncode($row.$propertyName);
                        if ($row.PSObject.Properties[$propertyStyle]) {
                            ## Cell styles override row styles
                            $propertyStyleHtml = (GetHtmlStyle -Style $Document.Styles[$row.$propertyStyle]).Trim();
                            [ref] $null = $standardTableBuilder.AppendFormat('<td style="{0}">{1}</td>', $propertyStyleHtml, $encodedHtmlContent);
                        }
                        elseif (($row.PSObject.Properties['__Style']) -and (-not [System.String]::IsNullOrEmpty($row.__Style))) {
                            ## We have a row style
                            $rowStyleHtml = (GetHtmlStyle -Style $Document.Styles[$row.__Style]).Trim();
                            [ref] $null = $standardTableBuilder.AppendFormat('<td style="{0}">{1}</td>', $rowStyleHtml, $encodedHtmlContent);
                        }
                        else {
                            if ($null -ne $row.$propertyName) {
                                ## Check that the property has a value
                                [ref] $null = $standardTableBuilder.AppendFormat('<td>{0}</td>', $encodedHtmlContent);
                            }
                            else {
                                [ref] $null = $standardTableBuilder.Append('<td>&nbsp</td>');
                            }
                        } #end if $row.PropertyStyle
                    } #end foreach property
                    [ref] $null = $standardTableBuilder.AppendLine('</tr>');
                } #end foreach row
                [ref] $null = $standardTableBuilder.AppendLine('</tbody></table></div>');
                return $standardTableBuilder.ToString();

            } #end process
        } #end function GetHtmlTableList


        function OutHtmlTable {
        <#
            .SYNOPSIS
                Output formatted Html <table> from PScribo.Table object.
            .NOTES
                One table is output per table row with the -List parameter.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()] [System.Object] $Table
            )
            process {

                [System.Text.StringBuilder] $tableBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                if ($Table.List) {
                    ## Create a table for each row
                    for ($r = 0; $r -lt $Table.Rows.Count; $r++) {
                        $row = $Table.Rows[$r];
                        if ($r -gt 0) {
                            ## Add a space between each table to mirror Word output rendering
                            [ref] $null = $tableBuilder.AppendLine('<p />');
                        }
                        [ref] $null = $tableBuilder.Append((GetHtmlTableList -Table $Table -Row $row));

                    } #end foreach row
                }
                else {
                    [ref] $null = $tableBuilder.Append((GetHtmlTable -Table $Table));
                } #end if
                return $tableBuilder.ToString();
                #Write-Output ($tableBuilder.ToString()) -NoEnumerate;

            } #end process
        } #end function outhtmltable


        function OutHtmlLineBreak {
        <#
            .SYNOPSIS
                Output formatted Html line break.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param ( )
            process {

                return '<hr />';

            }
        } #end function OutHtmlLineBreak


        function OutHtmlPageBreak {
        <#
            .SYNOPSIS
                Output formatted Html page break.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param ( )
            process {

                [System.Text.StringBuilder] $pageBreakBuilder = New-Object 'System.Text.StringBuilder';
                [ref] $null = $pageBreakBuilder.Append('</div></page>');
                $topMargin = ConvertMmToEm $Document.Options['MarginTop'];
                $leftMargin = ConvertMmToEm $Document.Options['MarginLeft'];
                $bottomMargin = ConvertMmToEm $Document.Options['MarginBottom'];
                $rightMargin = ConvertMmToEm $Document.Options['MarginRight'];
                [ref] $null = $pageBreakBuilder.AppendFormat('<page><div class="{0}" style="padding-top: {1}em; padding-left: {2}em; padding-bottom: {3}em; padding-right: {4}em;">', $Document.DefaultStyle, $topMargin, $leftMargin, $bottomMargin, $rightMargin).AppendLine();
                return $pageBreakBuilder.ToString();

            }
        } #end function OutHtmlPageBreak

        #endregion OutHtml Private Functions

# SIG # Begin signature block
# MIIXtwYJKoZIhvcNAQcCoIIXqDCCF6QCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUz5R5c5nep1G8qzi8/w9dtKh+
# T4ugghLqMIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUsdDOeZMWYDRMvYh4uFPvSQ3QTfQwDQYJKoZI
# hvcNAQEBBQAEggEAfWJsk0rVlT91YTbJ8kP1JrAvEJxdxuNf6ei/3XfxLSRzhue0
# gW3n3mn6AEIMwWTw86EBBNVZyfa7FofJWIif0rpYnMaII9DOZ5uJtjx/tppOf/E6
# WFIbrjFP6QhnFwu7lSAVmU5yu7xYRaeTtt3kF8/iG74F1VLkmi3yuWoqMbb2yxVF
# L67/aO+IwO4VFgh/pOxgLpvgUDF+6jRBfGOvkf61fe0W3xnM5wilxZCoDXoRh92w
# sGL5mfgZziCPfpsUGakDPdkGJS0fASD2wLQXF3c/7H8+2DD2LHpzO6qiQzrQZ8rk
# X5cj16Fnx59ALI45iwd+vnzFZKBQMDvq/UOOM6GCAgswggIHBgkqhkiG9w0BCQYx
# ggH4MIIB9AIBATByMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBD
# b3Jwb3JhdGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2
# aWNlcyBDQSAtIEcyAhAOz/Q4yP6/NW4E2GqYGxpQMAkGBSsOAwIaBQCgXTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA4MDEyMTQ2
# MTRaMCMGCSqGSIb3DQEJBDEWBBSEVKAoxmltI95TOdJPvWmdrvPYkTANBgkqhkiG
# 9w0BAQEFAASCAQBS/wSw9BmCrGi8tjTQBodi2wpg4QeV+6rlq4c14SJO808kI2Im
# vyMepUXqJVKb8weXZRjnz1N34QQEYmjTYfeShlj+LAx2U6JBuZ1N02yWNvQjUMeJ
# FjjRrNoYjj7vOcIBSBuCj/w5pCVN22ow4XCDVwdrMZohutAdHOHDhVDbe3tVQ+fz
# sibS0u/IxM1DfU19asXnEWpNIDJDkFg01TlSOJ3GbAWvmDQt3o8DN6nuicZxzziz
# 8S7pKUsNhL4T7erDshS2b9My/YufrzmXD1ctWi7ZN9xuB6VAbjRwtcanCc3C8gnD
# eCBU5mwpjLXDO9RjNbVacvQb43gJsVS6MsI0
# SIG # End signature block
