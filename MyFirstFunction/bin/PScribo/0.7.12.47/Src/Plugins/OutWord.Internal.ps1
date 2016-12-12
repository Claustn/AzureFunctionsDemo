        #region OutWord Private Functions

        function ConvertToWordColor {
        <#
            .SYNOPSIS
                Converts an HTML color to RRGGBB value as Word does not support short Html color codes
        #>
            [CmdletBinding()]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.String] $Color
            )
            process {

                $Color = $Color.TrimStart('#');
                if ($Color.Length -eq 3) {
                    $Color = '{0}{0}{1}{1}{2}{2}' -f $Color[0], $Color[1],$Color[2];
                }
                return $Color.ToUpper();

            }
        } #end function ConvertToWordColor


        function OutWordSection {
        <#
            .SYNOPSIS
                Output formatted Word section (paragraph).
        #>
            [CmdletBinding()]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Section,

                [Parameter(Mandatory)]
                [System.Xml.XmlElement] $RootElement,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {

                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

                $p = $RootElement.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                if (-not [System.String]::IsNullOrEmpty($Section.Style)) {
                    #if (-not $Section.IsExcluded) {
                        ## If it's excluded we need a non-Heading style :( Could explicitly set the style on the run?
                        $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain));
                        [ref] $null = $pStyle.SetAttribute('val', $xmlnsMain, $Section.Style);
                    #}
                }
                $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain));
                ## Increment heading spacing by 2pt for each section level, starting at 8pt for level 0, 10pt for level 1 etc
                $spacingPt = (($Section.Level * 2) + 8) * 20;
                [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, $spacingPt);
                [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, $spacingPt);
                $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                if ($Document.Options['EnableSectionNumbering']) { [string] $sectionName = '{0} {1}' -f $Section.Number, $Section.Name; }
                else { [string] $sectionName = '{0}' -f $Section.Name; }
                [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($sectionName));

                foreach ($s in $Section.Sections.GetEnumerator()) {
                    if ($s.Id.Length -gt 40) { $sectionId = '{0}[..]' -f $s.Id.Substring(0,36); }
                    else { $sectionId = $s.Id; }
                    $currentIndentationLevel = 1;
                    if ($null -ne $s.PSObject.Properties['Level']) { $currentIndentationLevel = $s.Level +1; }
                    WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
                    switch ($s.Type) {
                        'PScribo.Section' { $s | OutWordSection -RootElement $RootElement -XmlDocument $XmlDocument; }
                        'PScribo.Paragraph' { [ref] $null = $RootElement.AppendChild((OutWordParagraph -Paragraph $s -XmlDocument $XmlDocument)); }
                        'PScribo.PageBreak' { [ref] $null = $RootElement.AppendChild((OutWordPageBreak -PageBreak $s -XmlDocument $xmlDocument)); }
                        'PScribo.LineBreak' { [ref] $null = $RootElement.AppendChild((OutWordLineBreak -LineBreak $s -XmlDocument $xmlDocument)); }
                        'PScribo.Table' { OutWordTable -Table $s -XmlDocument $xmlDocument -Element $RootElement; }
                        'PScribo.BlankLine' { OutWordBlankLine -BlankLine $s -XmlDocument $xmlDocument -Element $RootElement; }
                        Default { WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning; }
                    } #end switch
                } #end foreach

            } #end process
        } #end function OutWordSection


        function OutWordParagraph {
        <#
            .SYNOPSIS
                Output formatted Word paragraph.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Paragraph,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {

                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

                $p = $XmlDocument.CreateElement('w', 'p', $xmlnsMain);
                $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                if ($Paragraph.Tabs -gt 0) {
                    $ind = $pPr.AppendChild($XmlDocument.CreateElement('w', 'ind', $xmlnsMain));
                    [ref] $null = $ind.SetAttribute('left', $xmlnsMain, (720 * $Paragraph.Tabs));
                }
                if (-not [System.String]::IsNullOrEmpty($Paragraph.Style)) {
                    $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain));
                    [ref] $null = $pStyle.SetAttribute('val', $xmlnsMain, $Paragraph.Style);
                }
                $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain));
                [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 0);
                [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 0);

                $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                $rPr = $r.AppendChild($XmlDocument.CreateElement('w', 'rPr', $xmlnsMain));
                ## Apply custom paragraph styles to the run..
                if ($Paragraph.Font) {
                    $rFonts = $rPr.AppendChild($XmlDocument.CreateElement('w', 'rFonts', $xmlnsMain));
                    [ref] $null = $rFonts.SetAttribute('ascii', $xmlnsMain, $Paragraph.Font[0]);
                    [ref] $null = $rFonts.SetAttribute('hAnsi', $xmlnsMain, $Paragraph.Font[0]);
                }
                if ($Paragraph.Size -gt 0) {
                    $sz = $rPr.AppendChild($XmlDocument.CreateElement('w', 'sz', $xmlnsMain));
                    [ref] $null = $sz.SetAttribute('val', $xmlnsMain, $Paragraph.Size * 2);
                }
                if ($Paragraph.Bold -eq $true) {
                    [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'b', $xmlnsMain));
                }
                if ($Paragraph.Italic -eq $true) {
                    [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'i', $xmlnsMain));
                }
                if ($Paragraph.Underline -eq $true) {
                    $u = $rPr.AppendChild($XmlDocument.CreateElement('w', 'u', $xmlnsMain));
                    [ref] $null = $u.SetAttribute('val', $xmlnsMain, 'single');
                }
                if (-not [System.String]::IsNullOrEmpty($Paragraph.Color)) {
                    $color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color', $xmlnsMain));
                    [ref] $null = $color.SetAttribute('val', $xmlnsMain, (ConvertToWordColor -Color $Paragraph.Color));
                }

                $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                [ref] $null = $t.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve'); ## needs to be xml:space="preserve" NOT w:space...
                if ([System.String]::IsNullOrEmpty($Paragraph.Text)) {
                    [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($Paragraph.Id));
                }
                else {
                    [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($Paragraph.Text));
                }
                return $p;

            } #end process
        } #end function OutWordParagraph


        function OutWordPageBreak {
            <#
            .SYNOPSIS
                Output formatted Word page break.
            #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $PageBreak,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {

                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $p = $XmlDocument.CreateElement('w', 'p', $xmlnsMain);
                $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                $br = $r.AppendChild($XmlDocument.CreateElement('w', 'br', $xmlnsMain));
                [ref] $null = $br.SetAttribute('type', $xmlnsMain, 'page');
                return $p;

            }
        } #end function OutWordPageBreak


        function OutWordLineBreak {
        <#
            .SYNOPSIS
                Output formatted Word line break.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $LineBreak,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {

                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $p = $XmlDocument.CreateElement('w', 'p', $xmlnsMain);
                $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                $pBdr = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pBdr', $xmlnsMain));
                $bottom = $pBdr.AppendChild($XmlDocument.CreateElement('w', 'bottom', $xmlnsMain));
                [ref] $null = $bottom.SetAttribute('val', $xmlnsMain, 'single');
                [ref] $null = $bottom.SetAttribute('sz', $xmlnsMain, 6);
                [ref] $null = $bottom.SetAttribute('space', $xmlnsMain, 1);
                [ref] $null = $bottom.SetAttribute('color', $xmlnsMain, 'auto');
                return $p;

            }
        } #end function OutWordLineBreak


        function GetWordTable {
        <#
            .SYNOPSIS
                Creates a scaffold Word <w:tbl> element
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {

                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $tableStyle = $Document.TableStyles[$Table.Style];
                $tbl = $XmlDocument.CreateElement('w', 'tbl', $xmlnsMain);
                $tblPr = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tblPr', $xmlnsMain));
                if ($Table.Tabs -gt 0) {
                    $tblInd = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblInd', $xmlnsMain));
                    [ref] $null = $tblInd.SetAttribute('w', $xmlnsMain, (720 * $Table.Tabs));
                }
                if ($Table.ColumnWidths) {
                    $tblLayout = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblLayout', $xmlnsMain));
                    [ref] $null = $tblLayout.SetAttribute('type', $xmlnsMain, 'fixed');
                }
                elseif ($Table.Width -eq 0) {
                    $tblLayout = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblLayout', $xmlnsMain));
                    [ref] $null = $tblLayout.SetAttribute('type', $xmlnsMain, 'autofit');
                }

                if ($Table.Width -gt 0) {
                    $tblW = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblW', $xmlnsMain));
                    [ref] $null = $tblW.SetAttribute('type', $xmlnsMain, 'pct');
                    $tableWidthRenderPct = $Table.Width;

                    if ($Table.Tabs -gt 0) {
                        ## We now need to deal with tables being pushed outside the page margin
                        $pageWidthMm = $Document.Options['PageWidth'] - ($Document.Options['PageMarginLeft'] + $Document.Options['PageMarginRight']);
                        $indentWidthMm = ConvertPtToMm -Point ($Table.Tabs * 36);
                        $tableRenderMm = (($pageWidthMm / 100) * $Table.Width) + $indentWidthMm;
                        if ($tableRenderMm -gt $pageWidthMm) {
                            ## We've over-flowed so need to work out the maximum percentage
                            $maxTableWidthMm = $pageWidthMm - $indentWidthMm;
                            $tableWidthRenderPct = [System.Math]::Round(($maxTableWidthMm / $pageWidthMm) * 100, 2);
                            WriteLog -Message ($localized.TableWidthOverflowWarning -f $tableWidthRenderPct) -IsWarning;
                        }
                    }
                    [ref] $null = $tblW.SetAttribute('w', $xmlnsMain, $tableWidthRenderPct * 50);
                }

                $spacing = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain));
                [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 72);
                [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 72);

                #$tblLook = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblLook', $xmlnsMain));
                #[ref] $null = $tblLook.SetAttribute('val', $xmlnsMain, '04A0');
                #[ref] $null = $tblLook.SetAttribute('firstRow', $xmlnsMain, 1);
                ## <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
                #$tblStyle = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblStyle', $xmlnsMain));
                #[ref] $null = $tblStyle.SetAttribute('val', $xmlnsMain, $Table.Style);

                if ($tableStyle.BorderWidth -gt 0) {
                    $tblBorders = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblBorders', $xmlnsMain));
                    foreach ($border in @('top','bottom','start','end','insideH','insideV')) {
                        $b = $tblBorders.AppendChild($XmlDocument.CreateElement('w', $border, $xmlnsMain));
                        [ref] $null = $b.SetAttribute('sz', $xmlnsMain, (ConvertMmToOctips $tableStyle.BorderWidth));
                        [ref] $null = $b.SetAttribute('val', $xmlnsMain, 'single');
                        [ref] $null = $b.SetAttribute('color', $xmlnsMain, (ConvertToWordColor -Color $tableStyle.BorderColor));
                    }
                }

                $tblCellMar = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblCellMar', $xmlnsMain));
                $top = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'top', $xmlnsMain));
                [ref] $null = $top.SetAttribute('w', $xmlnsMain, (ConvertMmToTwips $tableStyle.PaddingTop));
                [ref] $null = $top.SetAttribute('type', $xmlnsMain, 'dxa');
                $left = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'start', $xmlnsMain));
                [ref] $null = $left.SetAttribute('w', $xmlnsMain, (ConvertMmToTwips $tableStyle.PaddingLeft));
                [ref] $null = $left.SetAttribute('type', $xmlnsMain, 'dxa');
                $bottom = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'bottom', $xmlnsMain));
                [ref] $null = $bottom.SetAttribute('w', $xmlnsMain, (ConvertMmToTwips $tableStyle.PaddingBottom));
                [ref] $null = $bottom.SetAttribute('type', $xmlnsMain, 'dxa');
                $right = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'end', $xmlnsMain));
                [ref] $null = $right.SetAttribute('w', $xmlnsMain, (ConvertMmToTwips $tableStyle.PaddingRight));
                [ref] $null = $right.SetAttribute('type', $xmlnsMain, 'dxa');

                $tblGrid = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tblGrid', $xmlnsMain));
                $columnCount = $Table.Columns.Count;
                if ($Table.List) {
                    $columnCount = 2;
                }
                for ($i = 0; $i -lt $Table.Columns.Count; $i++) {
                    $gridCol = $tblGrid.AppendChild($XmlDocument.CreateElement('w', 'gridCol', $xmlnsMain));
                }

                return $tbl;

            } #end process
        } #end function GetWordTable


        function OutWordTable {
        <#
            .SYNOPSIS
                Output formatted Word table.
            .NOTES
                Specifies that the current row should be repeated at the top each new page on which the table is displayed. E.g, <w:tblHeader />.
        #>
            [CmdletBinding()]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table,

                ## Root element to append the table(s) to. List view will create multiple tables
                [Parameter(Mandatory)]
                [ValidateNotNull()]
                [System.Xml.XmlElement] $Element,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {

                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $tableStyle = $Document.TableStyles[$Table.Style];
                $headerStyle = $Document.Styles[$tableStyle.HeaderStyle];

                if ($Table.List) {

                    for ($r = 0; $r -lt $Table.Rows.Count; $r++) {
                        $row = $Table.Rows[$r];
                        if ($r -gt 0) {
                            ## Add a space between each table as Word renders them together..
                            [ref] $null = $Element.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                        }

                        ## Create <tr><tc></tc></tr> for each property
                        $tbl = $Element.AppendChild((GetWordTable -Table $Table -XmlDocument $XmlDocument));

                        $properties = @($row.PSObject.Properties);
                        for ($i = 0; $i -lt $properties.Count; $i++) {
                            $propertyName = $properties[$i].Name;
                            ## Ignore __Style properties
                            if (-not $propertyName.EndsWith('__Style')) {

                                $tr = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tr', $xmlnsMain));
                                $tc1 = $tr.AppendChild($XmlDocument.CreateElement('w', 'tc', $xmlnsMain));
                                $tcPr1 = $tc1.AppendChild($XmlDocument.CreateElement('w', 'tcPr', $xmlnsMain));

                                if ($null -ne $Table.ColumnWidths) {
                                    ## TODO: Refactor out
                                    $columnWidthTwips = ConvertMmToTwips -Millimeter $Table.ColumnWidths[0];
                                    $tcW1 = $tcPr1.AppendChild($XmlDocument.CreateElement('w', 'tcW', $xmlnsMain));
                                    [ref] $null = $tcW1.SetAttribute('w', $xmlnsMain, $Table.ColumnWidths[0] * 50);
                                    [ref] $null = $tcW1.SetAttribute('type', $xmlnsMain, 'pct');
                                }
                                if ($headerStyle.BackgroundColor) {
                                    [ref] $null = $tc1.AppendChild((GetWordTableStyleCellPr -Style $headerStyle -XmlDocument $XmlDocument));
                                }
                                $p1 = $tc1.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                                $pPr1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                                $pStyle1 = $pPr1.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain));
                                [ref] $null = $pStyle1.SetAttribute('val', $xmlnsMain, $tableStyle.HeaderStyle);
                                $r1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                                $t1 = $r1.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                                [ref] $null = $t1.AppendChild($XmlDocument.CreateTextNode($propertyName));

                                $tc2 = $tr.AppendChild($XmlDocument.CreateElement('w', 'tc', $xmlnsMain));
                                $tcPr2 = $tc2.AppendChild($XmlDocument.CreateElement('w', 'tcPr', $xmlnsMain));

                                if ($null -ne $Table.ColumnWidths) {
                                    ## TODO: Refactor out
                                    $tcW2 = $tcPr2.AppendChild($XmlDocument.CreateElement('w', 'tcW', $xmlnsMain));
                                    [ref] $null = $tcW2.SetAttribute('w', $xmlnsMain, $Table.ColumnWidths[1] * 50);
                                    [ref] $null = $tcW2.SetAttribute('type', $xmlnsMain, 'pct');
                                }

                                $p2 = $tc2.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                                $cellPropertyStyle = '{0}__Style' -f $propertyName;
                                if ($row.PSObject.Properties[$cellPropertyStyle]) {
                                    if (-not (Test-Path -Path Variable:\cellStyle)) {
                                        $cellStyle = $Document.Styles[$row.$cellPropertyStyle];
                                    }
                                    elseif ($cellStyle.Id -ne $row.$cellPropertyStyle) {
                                        ## Retrieve the style if we don't already have it
                                        $cellStyle = $Document.Styles[$row.$cellPropertyStyle];
                                    }
                                    if ($cellStyle.BackgroundColor) {
                                        [ref] $null = $tc2.AppendChild((GetWordTableStyleCellPr -Style $cellStyle -XmlDocument $XmlDocument));
                                    }
                                    if ($row.$cellPropertyStyle) {
                                        $pPr2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                                        $pStyle2 = $pPr2.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain));
                                        [ref] $null = $pStyle2.SetAttribute('val', $xmlnsMain, $row.$cellPropertyStyle);
                                    }
                                }

                                $r2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                                $t2 = $r2.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                                [ref] $null = $t2.AppendChild($XmlDocument.CreateTextNode($row.($propertyName)));
                            }
                        } #end for each property
                     } #end foreach row
                } #end if Table.List
                else {

                    $tbl = $Element.AppendChild((GetWordTable -Table $Table -XmlDocument $XmlDocument));

                    $tr = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tr', $xmlnsMain));
                    $trPr = $tr.AppendChild($XmlDocument.CreateElement('w', 'trPr', $xmlnsMain));
                    [ref] $rblHeader = $trPr.AppendChild($XmlDocument.CreateElement('w', 'tblHeader', $xmlnsMain)); ## Flow headers across pages
                    for ($i = 0; $i -lt $Table.Columns.Count; $i++) {
                        $tc = $tr.AppendChild($XmlDocument.CreateElement('w', 'tc', $xmlnsMain));
                        if ($headerStyle.BackgroundColor) {
                            $tcPr = $tc.AppendChild((GetWordTableStyleCellPr -Style $headerStyle -XmlDocument $XmlDocument));
                        }
                        else {
                            $tcPr = $tc.AppendChild($XmlDocument.CreateElement('w', 'tcPr', $xmlnsMain));
                        }
                        $tcW = $tcPr.AppendChild($XmlDocument.CreateElement('w', 'tcW', $xmlnsMain));

                        if (($Table.ColumnWidths -ne $null) -and ($Table.ColumnWidths[$i] -ne $null)) {
                            [ref] $null = $tcW.SetAttribute('w', $xmlnsMain, $Table.ColumnWidths[$i] * 50);
                            [ref] $null = $tcW.SetAttribute('type', $xmlnsMain, 'pct');
                        }
                        else {
                            [ref] $null = $tcW.SetAttribute('w', $xmlnsMain, 0);
                            [ref] $null = $tcW.SetAttribute('type', $xmlnsMain, 'auto');
                        }

                        $p = $tc.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                        $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                        $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain));
                        [ref] $null = $pStyle.SetAttribute('val', $xmlnsMain, $tableStyle.HeaderStyle);
                        $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                        $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                        [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($Table.Columns[$i]));
                    } #end for Table.Columns

                    $isAlternatingRow = $false;
                    foreach ($row in $Table.Rows) {
                        $tr = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tr', $xmlnsMain));
                        foreach ($propertyName in $Table.Columns) {
                            $cellPropertyStyle = '{0}__Style' -f $propertyName;
                            if ($row.PSObject.Properties[$cellPropertyStyle]) {
                                ## Cell style overrides row/default styles
                                $cellStyleName = $row.$cellPropertyStyle;
                            }
                            elseif (-not [System.String]::IsNullOrEmpty($row.__Style)) {
                                ## Row style overrides default style
                                $cellStyleName = $row.__Style;
                            }
                            else {
                                ## Use the table row/alternating style..
                                $cellStyleName = $tableStyle.RowStyle;
                                if ($isAlternatingRow) {
                                    $cellStyleName = $tableStyle.AlternateRowStyle;
                                }
                            }

                            if (-not (Test-Path -Path Variable:\cellStyle)) {
                                $cellStyle = $Document.Styles[$cellStyleName];
                            }
                            elseif ($cellStyle.Id -ne $cellStyleName) {
                                ## Retrieve the style if we don't already have it
                                $cellStyle = $Document.Styles[$cellStyleName];
                            }

                            $tc = $tr.AppendChild($XmlDocument.CreateElement('w', 'tc', $xmlnsMain));
                            if ($cellStyle.BackgroundColor) {
                                [ref] $null = $tc.AppendChild((GetWordTableStyleCellPr -Style $cellStyle -XmlDocument $XmlDocument));
                            }
                            $p = $tc.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                            $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                            $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain));
                            [ref] $null = $pStyle.SetAttribute('val', $xmlnsMain, $cellStyleName);
                            $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                            $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                            [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($row.($propertyName)));
                        } #end foreach property
                        $isAlternatingRow = !$isAlternatingRow;
                    } #end foreach row
                } #end if not Table.List

            } #end process
        } #end function OutWordTable


        function OutWordTOC {
        <#
            .SYNOPSIS
                 Output formatted Word table of contents.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $TOC,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {

                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $sdt = $XmlDocument.CreateElement('w', 'sdt', $xmlnsMain);
                $sdtPr = $sdt.AppendChild($XmlDocument.CreateElement('w', 'sdtPr', $xmlnsMain));
                $docPartObj = $sdtPr.AppendChild($XmlDocument.CreateElement('w', 'docPartObj', $xmlnsMain));
                $docObjectGallery = $docPartObj.AppendChild($XmlDocument.CreateElement('w', 'docPartGallery', $xmlnsMain));
                [ref] $null = $docObjectGallery.SetAttribute('val', $xmlnsMain, 'Table of Contents');
                [ref] $null = $docPartObj.AppendChild($XmlDocument.CreateElement('w', 'docPartUnique', $xmlnsMain));
                $sdtEndPr = $sdt.AppendChild($XmlDocument.CreateElement('w', 'stdEndPr', $xmlnsMain));

                $sdtContent = $sdt.AppendChild($XmlDocument.CreateElement('w', 'stdContent', $xmlnsMain));
                $p1 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
	            $pPr1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                $pStyle1 = $pPr1.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain));
                [ref] $null = $pStyle1.SetAttribute('val', $xmlnsMain, 'TOC');
                $r1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                $t1 = $r1.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                [ref] $null = $t1.AppendChild($XmlDocument.CreateTextNode($TOC.Name));

                $p2 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
	            $pPr2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                $tabs2 = $pPr2.AppendChild($XmlDocument.CreateElement('w', 'tabs', $xmlnsMain));
                $tab2 = $tabs2.AppendChild($XmlDocument.CreateElement('w', 'tab', $xmlnsMain));
                [ref] $null = $tab2.SetAttribute('val', $xmlnsMain, 'right');
                [ref] $null = $tab2.SetAttribute('leader', $xmlnsMain, 'dot');
                [ref] $null = $tab2.SetAttribute('pos', $xmlnsMain, '9016'); #10790?!
                $r2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                ##TODO: Refactor duplicate code
                $fldChar1 = $r2.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlnsMain));
                [ref] $null = $fldChar1.SetAttribute('fldCharType', $xmlnsMain, 'begin');

                $r3 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                $instrText = $r3.AppendChild($XmlDocument.CreateElement('w', 'instrText', $xmlnsMain));
                [ref] $null = $instrText.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve');
                [ref] $null = $instrText.AppendChild($XmlDocument.CreateTextNode(' TOC \o "1-3" \h \z \u '));

                $r4 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                $fldChar2 = $r4.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlnsMain));
                [ref] $null = $fldChar2.SetAttribute('fldCharType', $xmlnsMain, 'separate');

                $p3 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                $r5 = $p3.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
	            #$rPr3 = $r3.AppendChild($XmlDocument.CreateElement('w', 'rPr', $xmlnsMain));
                $fldChar3 = $r5.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlnsMain));
                [ref] $null = $fldChar3.SetAttribute('fldCharType', $xmlnsMain, 'end');

                return $sdt;

            } #end process
        } #end function OutWordTOC


        function OutWordBlankLine {
        <#
            .SYNOPSIS
                Output formatted Word xml blank line (paragraph).
        #>
            [CmdletBinding()]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $BlankLine,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument,

                [Parameter(Mandatory)]
                [System.Xml.XmlElement] $Element
            )
            process {

                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                for ($i = 0; $i -lt $BlankLine.LineCount; $i++) {
                    [ref] $null = $Element.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                }

            }
        } #end function OutWordLineBreak


        function GetWordStyle {
        <#
            .SYNOPSIS
                Generates Word Xml style element from a PScribo document style.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                ## PScribo document style
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Style,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument,

                [Parameter(Mandatory)]
                [ValidateSet('Paragraph','Character')]
                [System.String] $Type
            )
            process {

                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                if ($Type -eq 'Paragraph') {
                    $styleId = $Style.Id;
                    $styleName = $Style.Name;
                    $linkId = '{0}Char' -f $Style.Id;
                }
                else {
                    $styleId = '{0}Char' -f $Style.Id;
                    $styleName = '{0} Char' -f $Style.Name;
                    $linkId = $Style.Id;
                }


                $documentStyle = $XmlDocument.CreateElement('w', 'style', $xmlnsMain);
                [ref] $null = $documentStyle.SetAttribute('type', $xmlnsMain, $Type.ToLower());
                if ($Style.Id -eq $Document.DefaultStyle) {
                    ## Set as default style
                    [ref] $null = $documentStyle.SetAttribute('default', $xmlnsMain, 1);
                    $uiPriority = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'uiPriority', $xmlnsMain));
                    [ref] $null = $uiPriority.SetAttribute('val', $xmlnsMain, 1);
                }
                elseif (($Style.Id -eq 'Footer') -or ($Style.Id -eq 'Header')) {
                    ## Semi hide the styles named Footer and Header
                    [ref] $null = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'semiHidden', $xmlnsMain));
                }
                elseif (($document.TableStyles.Values | ForEach-Object { $_.HeaderStyle; $_.RowStyle; $_.AlternateRowStyle; }) -contains $Style.Id) {
                    ## Semi hide styles behind table styles (except default style!)
                    [ref] $null = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'semiHidden', $xmlnsMain));
                }

                [ref] $null = $documentStyle.SetAttribute('styleId', $xmlnsMain, $styleId);
                $documentStyleName = $documentStyle.AppendChild($xmlDocument.CreateElement('w', 'name', $xmlnsMain));
                [ref] $null = $documentStyleName.SetAttribute('val', $xmlnsMain, $styleName);
                $basedOn = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'basedOn', $xmlnsMain));
                [ref] $null = $basedOn.SetAttribute('val', $XmlnsMain, 'Normal');
                $link = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'link', $xmlnsMain));
                [ref] $null = $link.SetAttribute('val', $XmlnsMain, $linkId);
                $next = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'next', $xmlnsMain));
                [ref] $null = $next.SetAttribute('val', $xmlnsMain, 'Normal');
                $qFormat = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'qFormat', $xmlnsMain));
                $pPr = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                $keepNext = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepNext', $xmlnsMain));
                $keepLines = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepLines', $xmlnsMain));
                $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain));
                [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 0);
                [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 0);
                ## Set the <w:jc> (justification) element
                $jc = $pPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmlnsMain));
                if ($Style.Align.ToLower() -eq 'justify') {
                    [ref] $null = $jc.SetAttribute('val', $xmlnsMain, 'distribute');
                }
                else {
                    [ref] $null = $jc.SetAttribute('val', $xmlnsMain, $Style.Align.ToLower());
                }
                if ($Style.BackgroundColor) {
                    $shd = $pPr.AppendChild($XmlDocument.CreateElement('w', 'shd', $xmlnsMain));
                    [ref] $null = $shd.SetAttribute('val', $xmlnsMain, 'clear');
                    [ref] $null = $shd.SetAttribute('color', $xmlnsMain, 'auto');
                    [ref] $null = $shd.SetAttribute('fill', $xmlnsMain, (ConvertToWordColor -Color $Style.BackgroundColor));
                }
                [ref] $null = $documentStyle.AppendChild((GetWordStyleRunPr -Style $Style -XmlDocument $XmlDocument));

                return $documentStyle;

            } #end process
        } #end function GetWordStyle


        function GetWordTableStyle {
        <#
            .SYNOPSIS
                Generates Word Xml table style element from a PScribo document table style.
        #>
            [CmdletBinding()]
            param (
                ## PScribo document style
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $TableStyle,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {

                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $style = $XmlDocument.CreateElement('w', 'style', $xmlnsMain);
                [ref] $null = $style.SetAttribute('type', $xmlnsMain, 'table');
                [ref] $null = $style.SetAttribute('styleId', $xmlnsMain, $TableStyle.Id);
                $name = $style.AppendChild($XmlDocument.CreateElement('w', 'name', $xmlnsMain));
                [ref] $null = $name.SetAttribute('val', $xmlnsMain, $TableStyle.Id);
                $tblPr = $style.AppendChild($XmlDocument.CreateElement('w', 'tblPr', $xmlnsMain));
                $tblStyleRowBandSize = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblStyleRowBandSize', $xmlnsMain));
                [ref] $null = $tblStyleRowBandSize.SetAttribute('val', $xmlnsMain, 1);
                if ($tableStyle.BorderWidth -gt 0) {
                    $tblBorders = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblBorders', $xmlnsMain));
                    foreach ($border in @('top','bottom','start','end','insideH','insideV')) {
                        $b = $tblBorders.AppendChild($XmlDocument.CreateElement('w', $border, $xmlnsMain));
                        [ref] $null = $b.SetAttribute('sz', $xmlnsMain, (ConvertMmToOctips $tableStyle.BorderWidth));
                        [ref] $null = $b.SetAttribute('val', $xmlnsMain, 'single');
                        [ref] $null = $b.SetAttribute('color', $xmlnsMain, (ConvertToWordColor -Color $tableStyle.BorderColor));
                    }
                }
                [ref] $null = $style.AppendChild((GetWordTableStylePr -Style $Document.Styles[$TableStyle.HeaderStyle] -Type Header -XmlDocument $XmlDocument));
                [ref] $null = $style.AppendChild((GetWordTableStylePr -Style $Document.Styles[$TableStyle.RowStyle] -Type Row -XmlDocument $XmlDocument));
                [ref] $null = $style.AppendChild((GetWordTableStylePr -Style $Document.Styles[$TableStyle.AlternateRowStyle] -Type AlternateRow -XmlDocument $XmlDocument));
                return $style;

            }
        } #end function GetWordTableStyle


        function GetWordStyleParagraphPr {
        <#
            .SYNOPSIS
                Generates Word paragraph (pPr) formatting properties
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory)]
                [System.Object] $Style,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {

                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $pPr = $XmlDocument.CreateElement('w', 'pPr', $xmlnsMain);
                $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain));
                [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 0);
                [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 0);
                $keepNext = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepNext', $xmlnsMain));
                $keepLines = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepLines', $xmlnsMain));
                $jc = $pPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmlnsMain));
                if ($Style.Align.ToLower() -eq 'justify') { [ref] $null = $jc.SetAttribute('val', $xmlnsMain, 'distribute'); }
                else { [ref] $null = $jc.SetAttribute('val', $xmlnsMain, $Style.Align.ToLower()); }
                return $pPr;

            } #end process
        } #end function GetWordTableCellPr


        function GetWordStyleRunPrColor {
        <#
            .SYNOPSIS
                Generates Word run (rPr) text colour formatting property only.
            .NOTES
                This is only required to override the text colour in table rows/headers
                as I can't get this (yet) applied via the table style?
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory)]
                [System.Object] $Style,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {

                $rPr = $XmlDocument.CreateElement('w', 'rPr', $xmlnsMain);
                $color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color', $xmlnsMain));
                [ref] $null = $color.SetAttribute('val', $xmlnsMain, (ConvertToWordColor -Color $Style.Color));
                return $rPr;

            }
        } #end function GetWordStyleRunPrColor


        function GetWordStyleRunPr {
        <#
            .SYNOPSIS
                Generates Word run (rPr) formatting properties
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory)]
                [System.Object] $Style,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {

                $rPr = $XmlDocument.CreateElement('w', 'rPr', $xmlnsMain);
                $rFonts = $rPr.AppendChild($XmlDocument.CreateElement('w', 'rFonts', $xmlnsMain));
                [ref] $null = $rFonts.SetAttribute('ascii', $xmlnsMain, $Style.Font[0]);
                [ref] $null = $rFonts.SetAttribute('hAnsi', $xmlnsMain, $Style.Font[0]);
                if ($Style.Bold) {
                    [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'b', $xmlnsMain));
                }
                if ($Style.Underline) {
                    [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'u', $xmlnsMain));
                }
                if ($Style.Italic) {
                    [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'i', $xmlnsMain));
                }
                $color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color', $xmlnsMain));
                [ref] $null = $color.SetAttribute('val', $xmlnsMain, (ConvertToWordColor -Color $Style.Color));
                $sz = $rPr.AppendChild($XmlDocument.CreateElement('w', 'sz', $xmlnsMain));
                [ref] $null = $sz.SetAttribute('val', $xmlnsMain, $Style.Size * 2);
                return $rPr;

            } #end process
        } #end function GetWordStyleRunPr


        function GetWordTableStyleCellPr {
        <#
            .SYNOPSIS
                Generates Word table cell (tcPr) formatting properties
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory)]
                [System.Object] $Style,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {

                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $tcPr = $XmlDocument.CreateElement('w', 'tcPr', $xmlnsMain);
                if ($Style.BackgroundColor) {
                    $shd = $tcPr.AppendChild($XmlDocument.CreateElement('w', 'shd', $xmlnsMain));
                    [ref] $null = $shd.SetAttribute('val', $xmlnsMain, 'clear');
                    [ref] $null = $shd.SetAttribute('color', $xmlnsMain, 'auto');
                    [ref] $null = $shd.SetAttribute('fill', $xmlnsMain, (ConvertToWordColor -Color $Style.BackgroundColor));
                }
                return $tcPr;

            } #end process
        } #end function GetWordTableCellPr


        function GetWordTableStylePr {
        <#
            .SYNOPSIS
                Generates Word table style (tblStylePr) formatting properties for specified table style type
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory)]
                [System.Object] $Style,

                [Parameter(Mandatory)]
                [ValidateSet('Header','Row','AlternateRow')]
                [System.String] $Type,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {

                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $tblStylePr = $XmlDocument.CreateElement('w', 'tblStylePr', $xmlnsMain);
                $tblPr = $tblStylePr.AppendChild($XmlDocument.CreateElement('w', 'tblPr', $xmlnsMain));
                switch ($Type) {
                    'Header' { $tblStylePrType = 'firstRow'; }
                    'Row' { $tblStylePrType = 'band2Horz'; }
                    'AlternateRow' { $tblStylePrType = 'band1Horz'; }
                }
                [ref] $null = $tblStylePr.SetAttribute('type', $xmlnsMain, $tblStylePrType);
                [ref] $null = $tblStylePr.AppendChild((GetWordStyleParagraphPr -Style $Style -XmlDocument $XmlDocument));
                [ref] $null = $tblStylePr.AppendChild((GetWordStyleRunPr -Style $Style -XmlDocument $XmlDocument));
                [ref] $null = $tblStylePr.AppendChild((GetWordTableStyleCellPr -Style $Style -XmlDocument $XmlDocument));
                return $tblStylePr;

            } #end process
        } #end function GetWordTableStylePr

        function GetWordSectionPr {
        <#
            .SYNOPSIS
                Outputs Office Open XML section element to set page size and margins.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory)]
                [System.Single] $PageWidth,

                [Parameter(Mandatory)]
                [System.Single] $PageHeight,

                [Parameter(Mandatory)]
                [System.Single] $PageMarginTop,

                [Parameter(Mandatory)]
                [System.Single] $PageMarginLeft,

                [Parameter(Mandatory)]
                [System.Single] $PageMarginBottom,

                [Parameter(Mandatory)]
                [System.Single] $PageMarginRight,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {

                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $sectPr = $XmlDocument.CreateElement('w', 'sectPr', $xmlnsMain);
                $pgSz = $sectPr.AppendChild($XmlDocument.CreateElement('w', 'pgSz', $xmlnsMain));
                [ref] $null = $pgSz.SetAttribute('w', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageWidth));
                [ref] $null = $pgSz.SetAttribute('h', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageHeight));
                [ref] $null = $pgSz.SetAttribute('orient', $xmlnsMain, 'portrait');
                $pgMar = $sectPr.AppendChild($XmlDocument.CreateElement('w', 'pgMar', $xmlnsMain));
                [ref] $null = $pgMar.SetAttribute('top', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageMarginTop));
                [ref] $null = $pgMar.SetAttribute('bottom', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageMarginBottom));
                [ref] $null = $pgMar.SetAttribute('left', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageMarginLeft));
                [ref] $null = $pgMar.SetAttribute('right', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageMarginRight));
                return $sectPr;

            } #end process
        } #end GetWordSectionPr


        function OutWordStylesDocument {
        <#
            .SYNOPSIS
                Outputs Office Open XML style document
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlDocument])]
            param (
                ## PScribo document styles
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Collections.Hashtable] $Styles,

                ## PScribo document tables styles
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Collections.Hashtable] $TableStyles
            )
            process {

                ## Create the Style.xml document
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $xmlDocument = New-Object -TypeName 'System.Xml.XmlDocument';
                [ref] $null = $xmlDocument.AppendChild($xmlDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'));
                $documentStyles = $xmlDocument.AppendChild($xmlDocument.CreateElement('w', 'styles', $xmlnsMain));

                ## Create default style
                $defaultStyle = $documentStyles.AppendChild($xmlDocument.CreateElement('w', 'style', $xmlnsMain));
                [ref] $null = $defaultStyle.SetAttribute('type', $xmlnsMain, 'paragraph');
                [ref] $null = $defaultStyle.SetAttribute('default', $xmlnsMain, '1');
                [ref] $null = $defaultStyle.SetAttribute('styleId', $xmlnsMain, 'Normal');
                $defaultStyleName = $defaultStyle.AppendChild($xmlDocument.CreateElement('w', 'name', $xmlnsMain));
                [ref] $null = $defaultStyleName.SetAttribute('val', $xmlnsMain, 'Normal');
                [ref] $null = $defaultStyle.AppendChild($xmlDocument.CreateElement('w', 'qFormat', $xmlnsMain));

                foreach ($style in $Styles.Values) {
                    $documentParagraphStyle = GetWordStyle -Style $style -XmlDocument $xmlDocument -Type Paragraph;
                    [ref] $null = $documentStyles.AppendChild($documentParagraphStyle);
                    $documentCharacterStyle = GetWordStyle -Style $style -XmlDocument $xmlDocument -Type Character;
                    [ref] $null = $documentStyles.AppendChild($documentCharacterStyle);
                }
                foreach ($tableStyle in $TableStyles.Values) {
                    $documentTableStyle = GetWordTableStyle -TableStyle $tableStyle -XmlDocument $xmlDocument;
                    [ref] $null = $documentStyles.AppendChild($documentTableStyle);
                }
                return $xmlDocument;

            } #end process
        } #end function OutWordStyleDocument


        function OutWordSettingsDocument {
        <#
            .SYNOPSIS
                Outputs Office Open XML settings document
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlDocument])]
            param (
                [Parameter()]
                [System.Management.Automation.SwitchParameter] $UpdateFields
            )
            process {

                ## Create the Style.xml document
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                # <w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                # xmlns:o="urn:schemas-microsoft-com:office:office"
                # xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                # xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
                # xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word"
                # xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                # xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                # xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
                # xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"
                # mc:Ignorable="w14 w15">
                $settingsDocument = New-Object -TypeName 'System.Xml.XmlDocument';
                [ref] $null = $settingsDocument.AppendChild($settingsDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'));
                $settings = $settingsDocument.AppendChild($settingsDocument.CreateElement('w', 'settings', $xmlnsMain));
                ## Set compatibility mode to Word 2013
                $compat = $settings.AppendChild($settingsDocument.CreateElement('w', 'compat', $xmlnsMain));
                $compatSetting = $compat.AppendChild($settingsDocument.CreateElement('w', 'compatSetting', $xmlnsMain));
                [ref] $null = $compatSetting.SetAttribute('name', $xmlnsMain, 'compatibilityMode');
                [ref] $null = $compatSetting.SetAttribute('uri', $xmlnsMain, 'http://schemas.microsoft.com/office/word');
                [ref] $null = $compatSetting.SetAttribute('val', $xmlnsMain, 15);
                if ($UpdateFields) {
                    $wupdateFields = $settings.AppendChild($settingsDocument.CreateElement('w', 'updateFields', $xmlnsMain));
                    [ref] $null = $wupdateFields.SetAttribute('val', $xmlnsMain, 'true');
                }
                return $settingsDocument;

            } #end process
        } #end function OutWordSettingsDocument

        #endregion OutWord Private Functions
# SIG # Begin signature block
# MIIXtwYJKoZIhvcNAQcCoIIXqDCCF6QCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUavRl5fqgN6S401EePnTCYtwz
# EfWgghLqMIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUg63k3gk+Cq5IXzKlfnk5Er7eWmkwDQYJKoZI
# hvcNAQEBBQAEggEAZ5qOedzTYoNdPRpSzGaqgFRsbrfw2zB1Tewy4ddOMPhWwZy0
# OKPKRr3z7NfTP7CBc0+MVT6pUn1gi8yXUGXC0NBge1o8AEX6R0C9fazRMtZrfk6w
# Wkc6hJjn9c0Jrvi51+CrMqRpiR6WDvHxnLkE7jMmRD0EGy8BQNNXCGaqxW/LIwKy
# Ak3Sq/eUhT4Sh4mxOXisvSfvzS+yrXWBDxYT3K5SAQruEzKqUIkG8Q7uLK1ZuKQg
# y/1/2TcXZAdilAOstHKpijaeNq3Sn8X7IcuMG+4kkUbI9bSd5le7K0qS6iMUk6ZP
# q8gByfP2GrMs54rhYMp9YkdnamyTDNO7MEZ+DaGCAgswggIHBgkqhkiG9w0BCQYx
# ggH4MIIB9AIBATByMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBD
# b3Jwb3JhdGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2
# aWNlcyBDQSAtIEcyAhAOz/Q4yP6/NW4E2GqYGxpQMAkGBSsOAwIaBQCgXTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA4MDEyMTQ2
# MTVaMCMGCSqGSIb3DQEJBDEWBBSrlGP8yFsaRicbyfEi66dA905AmDANBgkqhkiG
# 9w0BAQEFAASCAQA2gmu19QQrHmVJ83BMHN1Y4WjDjC57WQMEBbRhwHaCCvv/0EDj
# /N1KXCR1sCSaIbPajtOPCq2NqyQ07yVJ7lMhRxeFyYl5CGRXv4itkEU4bKsqp6in
# jwI/a7T/v9STuess98akiI29t+tEwni6YRqLxAu0d0axZgQcpBtMSSqK7kU0xD0D
# zReY09my8dUATz+eiS4CiGYPke9PUJadzcVj9FY7T0oT+mehehoaslg5gOOY83R+
# cM6/+1OdaFLY6ECeWdbXun4SGYTIfeSKJbyXuFwcDVFVixcAt28bQFAn/NQjlxUl
# 75H4b5t2Fkm1S3bf5V7GrvwPf249OugrQLOy
# SIG # End signature block
