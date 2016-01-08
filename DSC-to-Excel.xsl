<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:math="http://www.w3.org/2005/xpath-functions/math"
    xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl" xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:x="urn:schemas-microsoft-com:office:excel"
    xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
    xmlns:html="http://www.w3.org/TR/REC-html40" xmlns:xlink="http://www.w3.org/1999/xlink"
    xmlns:mdc="http://mdc" xmlns="urn:isbn:1-931666-22-9" xmlns:ead="urn:isbn:1-931666-22-9"
    exclude-result-prefixes="xs math xd xlink mdc ead html" version="2.0">
    <xd:doc scope="stylesheet">
        <xd:desc>
            <xd:p><xd:b>Created on:</xd:b> Aug 16, 2014</xd:p>
            <xd:p><xd:b>Significantly revised on:</xd:b> August 2, 2015</xd:p>
            <xd:p><xd:b>Author:</xd:b> Mark Custer</xd:p>
            <xd:p>tested with Saxon-HE 9.6.0.5</xd:p>
        </xd:desc>
    </xd:doc>
    <xsl:output method="xml" encoding="UTF-8"/>

    <!--
        to do:
        
        should I add  ss:Hidden="1" to level-0 rows?  If so, they won't show up immediately in view
        but they will be grouped with their parents when sorted.
     -->
    <xsl:variable name="ead-copy-filename"
        select="
            if (ead:ead/ead:eadheader/ead:eadid[1]/normalize-space())
            then
                concat(ead:ead/ead:eadheader/ead:eadid[1], '.xml')
            else
                ''"/>

    <xsl:template match="/">
        <!-- storing a copy of the entire XML file so that it can be used later to re-create the collection-level information for roundtripping.
            also could use it to merge any unsupported features like bibliographies, links, etc., as long as @id attributes are present, but haven't dived that deep yet.-->
        <xsl:apply-templates/>
    </xsl:template>


    <xsl:template match="ead:ead">
        <!-- the "Excel-template" template does most of the work (and since it's so large, it's been pushed down to the end of this file);
            it also calls the archdesc/dsc section in order to create the primary worksheet for the container list-->
        <xsl:call-template name="Excel-Template"/>
    </xsl:template>

    <xsl:template
        match="
            ead:c | ead:c01 | ead:c02 | ead:c03 | ead:c04 | ead:c05 | ead:c06 | ead:c07 | ead:c08 | ead:c09
            | ead:c10 | ead:c11 | ead:c12"
        name="level-0">

        <xsl:param name="level-0" select="false()"/>
        <xsl:param name="current-position" select="1" as="xs:integer"/>

        <!-- 
            the following could be used if you want to limit the repeating fields to just 
            unidates (not bulk)
            container groups
            dao elements
            structured physdesc elements
            
            if so, you'll need to replace a number of 
            [position() eq $current-position] filters
            with for-each loops
            
        <xsl:variable name="depth-to-recurse" as="xs:integer">
            <xsl:sequence
                select="
                    if ($level-0 eq false())
                    then
                        max(
                        (count(ead:did/ead:unitdate[not(@type = 'bulk')])
                        ,
                        count(ead:did/ead:container[@id][not(@parent)])
                        ,
                        count(ead:did/ead:dao)
                        ,
                        count(ead:did/ead:physdesc/ead:extent[1]/matches(., '^\d'))
                        )
                        )
                    else 1"/>
        </xsl:variable>
        -->


        <xsl:variable name="did-items" as="item()*">
            <xsl:sequence select="ead:did/*[normalize-space()][not(self::ead:container)][not(self::ead:unitdate[@type='bulk'])][not(self::ead:physdesc[ead:extent[1]/matches(., '^\d')])][local-name() = following-sibling::*[normalize-space()]/local-name()]/local-name()"/>
        </xsl:variable>
        <xsl:variable name="non-did-items" as="item()*">
            <xsl:sequence select="ead:*[normalize-space()][not(local-name() = ('c', 'c01', 'c02', 'c03', 'c04', 'c05', 'c06', 'c07', 'c08', 'c09', 'c10', 'c11', 'c12'))][local-name() = following-sibling::*[normalize-space()]/local-name()]/local-name()"/>
        </xsl:variable>
        <xsl:variable name="depths" as="item()*">
            <xsl:sequence select="
                (
                    for $x
                    in distinct-values($did-items)
                    return count(index-of($did-items, $x))
                    , for $y 
                    in distinct-values($non-did-items)
                    return count(index-of($non-did-items, $y))
                    , count(ead:did/ead:container[@id][position() gt 1][not(@parent)])
                    , count(ead:did/ead:physdesc[position() gt 1]/ead:extent[1]/matches(., '^\d'))                
                )
                "/>
        </xsl:variable>
        <xsl:variable name="depth-to-recurse" select="max($depths)" as="xs:integer"/>


        <Row ss:Height="30" xmlns="urn:schemas-microsoft-com:office:spreadsheet">

            <xsl:if test="$level-0 eq true()">
                <xsl:attribute name="StyleID"
                    namespace="urn:schemas-microsoft-com:office:spreadsheet">
                    <xsl:text>s1</xsl:text>
                </xsl:attribute>
            </xsl:if>

            <!--this assumes that the YYYY-MM format will be used,
                not YYYYMM.
           also, if dateTimes are used, I could use those functions.-->
            <xsl:variable name="beginDates"
                select="
                    if (ead:did/ead:unitdate[not(@type = 'bulk')][position() eq $current-position]/contains(@normal, '/'))
                    then
                        tokenize(ead:did/ead:unitdate[not(@type = 'bulk')][position() eq $current-position]/substring-before(@normal, '/'), '-')
                    else
                        tokenize(ead:did/ead:unitdate[not(@type = 'bulk')][position() eq $current-position]/@normal, '-')"/>
            <xsl:variable name="endDates"
                select="tokenize(ead:did/ead:unitdate[not(@type = 'bulk')][position() eq $current-position]/substring-after(@normal, '/'), '-')"/>
            <xsl:variable name="bulkBeginDates"
                select="
                    if (ead:did/ead:unitdate[@type = 'bulk'][position() eq $current-position]/contains(@normal, '/'))
                    then
                        tokenize(ead:did/ead:unitdate[@type = 'bulk'][position() eq $current-position]/substring-before(@normal, '/'), '-')
                    else
                        tokenize(ead:did/ead:unitdate[@type = 'bulk'][position() eq $current-position]/@normal, '-')"/>
            <xsl:variable name="bulkEndDates"
                select="tokenize(ead:did/ead:unitdate[@type = 'bulk'][position() eq $current-position]/substring-after(@normal, '/'), '-')"/>

            <xsl:variable name="begin-year" select="$beginDates[1]"/>
            <xsl:variable name="begin-month" select="$beginDates[2]"/>
            <xsl:variable name="begin-day" select="$beginDates[3]"/>
            <xsl:variable name="end-year" select="$endDates[1]"/>
            <xsl:variable name="end-month" select="$endDates[2]"/>
            <xsl:variable name="end-day" select="$endDates[3]"/>

            <xsl:variable name="bulk-begin-year" select="$bulkBeginDates[1]"/>
            <xsl:variable name="bulk-begin-month" select="$bulkBeginDates[2]"/>
            <xsl:variable name="bulk-begin-day" select="$bulkBeginDates[3]"/>
            <xsl:variable name="bulk-end-year" select="$bulkEndDates[1]"/>
            <xsl:variable name="bulk-end-month" select="$bulkEndDates[2]"/>
            <xsl:variable name="bulk-end-day" select="$bulkEndDates[3]"/>

            <xsl:variable name="extent-number"
                select="
                    if (ead:did/ead:physdesc[position() eq $current-position]/ead:extent[1]/matches(., '^\d'))
                    then
                        number(substring-before(normalize-space(ead:did/ead:physdesc[position() eq $current-position]/ead:extent[1]), ' '))
                    else
                        ''"/>

            <xsl:variable name="extent-type"
                select="
                    if (ead:did/ead:physdesc[position() eq $current-position]/ead:extent[1]/matches(., '^\d'))
                    then
                        substring-after(normalize-space(ead:did/ead:physdesc[position() eq $current-position]/ead:extent[1]), ' ')
                    else
                        ''"/>


            <!-- column 1 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="Number">
                    <xsl:value-of
                        select="
                            if ($level-0 eq true()) then
                                0
                            else
                                count(ancestor::*) - 2"
                    />
                </Data>
            </Cell>
            <!-- column 2 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of
                        select="
                            if ($level-0 eq true()) then ''
                            else if (@level='otherlevel') then 'accession'
                            else if (@level) then @level
                            else 'file'"
                    />
                </Data>
            </Cell>
            <!-- column 3 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:did/ead:unitid[position() eq $current-position]"/>
                </Data>
            </Cell>
            <!-- column 4 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:did/ead:unittitle[position() eq $current-position]"/>
                </Data>
            </Cell>
            <!-- column 5 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates
                        select="ead:did/ead:unitdate[not(@type = 'bulk')][position() eq $current-position]"/>
                    <!--any repeating unidate values will be copied into the immediate following rows-->
                </Data>
                <NamedCell ss:Name="date_expression"/>
            </Cell>
            <!-- column 6 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of select="$begin-year"/>
                </Data>
                <NamedCell ss:Name="year_begin"/>
            </Cell>
            <!-- column 7 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of select="$begin-month"/>
                </Data>
                <NamedCell ss:Name="month_begin"/>
            </Cell>
            <!-- column 8 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of select="$begin-day"/>
                </Data>
                <NamedCell ss:Name="day_begin"/>
            </Cell>
            <!-- column 9 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of select="$end-year"/>
                </Data>
                <NamedCell ss:Name="year_end"/>
            </Cell>
            <!-- column 10 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of select="$end-month"/>
                </Data>
                <NamedCell ss:Name="month_end"/>
            </Cell>
            <!-- column 11 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of select="$end-day"/>
                </Data>
                <NamedCell ss:Name="day_end"/>
            </Cell>

            <!-- column 12 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of select="$bulk-begin-year"/>
                </Data>
                <NamedCell ss:Name="bulk_year_begin"/>
            </Cell>
            <!-- column 13 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of select="$bulk-begin-month"/>
                    <NamedCell ss:Name="bulk_month_begin"/>
                </Data>
            </Cell>
            <!-- column 14 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of select="$bulk-begin-day"/>
                </Data>
                <NamedCell ss:Name="bulk_day_begin"/>
            </Cell>
            <!-- column 15 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of select="$bulk-end-year"/>
                </Data>
                <NamedCell ss:Name="bulk_year_end"/>
            </Cell>
            <!-- column 16 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of select="$bulk-end-month"/>
                </Data>
                <NamedCell ss:Name="bulk_month_end"/>
            </Cell>
            <!-- column 17 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of select="$bulk-end-day"/>
                </Data>
                <NamedCell ss:Name="bulk_day_end"/>
            </Cell>


            <!-- column 18 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of
                        select="
                            if (ead:did/ead:container[@label][@id][not(@parent)][position() eq $current-position]/contains(@label, '['))
                            then
                                ead:did/ead:container[@label][@id][not(@parent)][position() eq $current-position]/substring-before(@label, ' [')
                            else
                                ead:did/ead:container[@label][@id][not(@parent)][position() eq $current-position]/@label"
                    />
                </Data>
                <NamedCell ss:Name="instance_type"/>
            </Cell>
            <!-- column 19 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates
                        select="ead:did/ead:container[@id][not(@parent)][position() eq $current-position]/@type"
                    />
                </Data>
                <NamedCell ss:Name="container_1_type"/>
            </Cell>
            <!-- column 20 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates
                        select="ead:did/ead:container[@id][not(@parent)][position() eq $current-position]/@altrender"
                    />
                </Data>
                <NamedCell ss:Name="container_profile"/>
            </Cell>
            <!-- column 21 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of
                        select="ead:did/ead:container[@id][not(@parent)][position() eq $current-position]/substring-after(substring-before(@label, ']'), '[')"
                    />
                </Data>
                <NamedCell ss:Name="barcode"/>
            </Cell>
            <!-- column 22 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates
                        select="ead:did/ead:container[@id][not(@parent)][position() eq $current-position]"
                    />
                </Data>
            </Cell>

            <!-- column 23 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates
                        select="ead:did/ead:container[@id][not(@parent)][position() eq $current-position]/following-sibling::ead:container[1][@parent]/@type"
                    />
                </Data>
                <NamedCell ss:Name="container_2_type"/>
            </Cell>
            <!-- column 24 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates
                        select="ead:did/ead:container[@id][not(@parent)][position() eq $current-position]/following-sibling::ead:container[1][@parent]"
                    />
                </Data>
            </Cell>

            <!-- column 25 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates
                        select="ead:did/ead:container[@id][not(@parent)][position() eq $current-position]/following-sibling::ead:container[2][@parent]/@type"
                    />
                </Data>
                <NamedCell ss:Name="container_3_type"/>
            </Cell>
            <!-- column 26 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates
                        select="ead:did/ead:container[@id][not(@parent)][position() eq $current-position]/following-sibling::ead:container[2][@parent]"
                    />
                </Data>
            </Cell>

            <!-- column 27 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of select="$extent-number"/>
                </Data>
                <NamedCell ss:Name="extent_number"/>
            </Cell>

            <!-- column 28 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of select="$extent-type"/>
                </Data>
                <NamedCell ss:Name="extent_value"/>
            </Cell>
            <!-- column 29 -->
            <!-- just change position of physdesc, not extent, if level 0 works-->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:did/ead:physdesc[position() eq $current-position]/ead:extent[2]"
                    />
                </Data>
                <NamedCell ss:Name="generic_extent"/>
            </Cell>

            <!-- column 30 -->
            <Cell ss:StyleID="s3">
                <!-- test! -->
                <Data ss:Type="String">
                    <xsl:apply-templates
                        select="ead:did/ead:physdesc[position() eq $current-position][not(ead:extent) or matches(ead:extent[1], '^\D')]"
                    />
                </Data>
            </Cell>

            <!-- column 31 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:did/ead:origination[position() eq $current-position]"/>
                </Data>
            </Cell>

            <!-- column 32 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:bioghist[position() eq $current-position]"/>
                </Data>
            </Cell>

            <!-- column 33 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:scopecontent[position() eq $current-position]"/>
                </Data>
            </Cell>

            <!-- column 34 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:arrangement[position() eq $current-position]"/>
                </Data>
            </Cell>

            <!-- column 35 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:accessrestrict[position() eq $current-position]"/>
                </Data>
            </Cell>

            <!-- column 36 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:phystech[position() eq $current-position]"/>
                </Data>
            </Cell>

            <!-- column 37 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:physloc[position() eq $current-position]"/>
                </Data>
            </Cell>

            <!-- column 38 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:userestrict[position() eq $current-position]"/>
                </Data>
            </Cell>

            <!-- column 39 -->
            <!-- just one, since that's all that the AT and ASpace will allow....  but perhaps better not to include at all?-->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of
                        select="
                            if ($level-0 eq false())
                            then
                                ead:did/ead:langmaterial/ead:language[1]/@langcode
                            else
                                ''"
                    />
                </Data>
            </Cell>

            <!-- column 40 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:langmaterial[position() eq $current-position]"/>
                </Data>
            </Cell>

            <!-- column 41 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:otherfindaid[position() eq $current-position]"/>
                </Data>
            </Cell>
            <!-- column 42 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:custodhist[position() eq $current-position]"/>
                </Data>
            </Cell>
            <!-- column 43 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:acqinfo[position() eq $current-position]"/>
                </Data>
            </Cell>
            <!-- column 44 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:appraisal[position() eq $current-position]"/>
                </Data>
            </Cell>
            <!-- column 45 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:accruals[position() eq $current-position]"/>
                </Data>
            </Cell>
            <!-- column 46 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:originalsloc[position() eq $current-position]"/>
                </Data>
            </Cell>
            <!-- column 47 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:altformavail[position() eq $current-position]"/>
                </Data>
            </Cell>
            <!-- column 48 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:relatedmaterial[position() eq $current-position]"/>
                </Data>
            </Cell>
            <!-- column 49 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:separatedmaterial[position() eq $current-position]"/>
                </Data>
            </Cell>
            <!-- column 50 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:prefercite[position() eq $current-position]"/>
                </Data>
            </Cell>
            <!-- column 51 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:processinfo[position() eq $current-position]"/>
                </Data>
            </Cell>

            <!-- column 52 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:apply-templates select="ead:controlaccess[position() eq $current-position]"/>
                </Data>
            </Cell>
            <!-- column 53 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of
                        select="
                            if ($level-0 eq false())
                            then
                                @id
                            else
                                ''"
                    />
                </Data>
                <NamedCell ss:Name="component_id"/>
            </Cell>

            <!-- column 54 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of
                        select="ead:did/ead:dao[position() eq $current-position]/@xlink:href"/>
                </Data>
            </Cell>
            <!-- column 55 -->
            <Cell ss:StyleID="s3">
                <Data ss:Type="String">
                    <xsl:value-of
                        select="ead:did/ead:dao[position() eq $current-position]/@xlink:title"/>
                </Data>
            </Cell>
        </Row>


        <!--
            recurisve template to handle repeating fields
        -->
        <xsl:if test="$current-position &lt;= $depth-to-recurse">
            <xsl:call-template name="level-0">
                <xsl:with-param name="current-position" select="$current-position + 1" as="xs:integer"/>
                <xsl:with-param name="level-0" select="true()"/>
            </xsl:call-template>
        </xsl:if>

        <xsl:if test="$level-0 eq false()">
            <xsl:apply-templates select="
                ead:c | ead:c02 | ead:c03 | ead:c04 | ead:c05 | ead:c06 | ead:c07 | ead:c08 | ead:c09
                | ead:c10 | ead:c11 | ead:c12"/>
        </xsl:if>
    </xsl:template>

    
    
    <!-- formatting templates -->
    <xsl:template match="ead:emph[@render = 'bold']">
        <html:B>
            <xsl:apply-templates/>
        </html:B>
    </xsl:template>
    <xsl:template match="ead:emph[@render = 'boldunderline']">
        <html:B>
            <html:U>
                <xsl:apply-templates/>
            </html:U>
        </html:B>
    </xsl:template>
    <xsl:template match="ead:emph[@render = 'bolditalic']">
        <html:B>
            <html:I>
                <xsl:apply-templates/>
            </html:I>
        </html:B>
    </xsl:template>
    <xsl:template match="ead:emph[@render = 'boldsmcaps']">
        <html:B>
            <html:Font html:Size="8">
                <xsl:apply-templates/>
            </html:Font>
        </html:B>
    </xsl:template>
    <xsl:template match="ead:emph[@render = 'italic']">
        <html:I>
            <xsl:apply-templates/>
        </html:I>
    </xsl:template>
    <xsl:template match="ead:emph[@render = 'underline']">
        <html:U>
            <xsl:apply-templates/>
        </html:U>
    </xsl:template>
    <xsl:template match="ead:emph[@render = 'super']">
        <html:Sup>
            <xsl:apply-templates/>
        </html:Sup>
    </xsl:template>
    <xsl:template match="ead:emph[@render = 'sub']">
        <html:Sub>
            <xsl:apply-templates/>
        </html:Sub>
    </xsl:template>
    <xsl:template match="ead:emph[@render = 'nonproport']">
        <html:Font html:Face='Courier New'>
            <xsl:apply-templates/>
        </html:Font>
    </xsl:template>
    <xsl:template match="ead:emph[@render = 'smcaps']">
        <html:Font html:Size="8">
            <xsl:apply-templates/>
        </html:Font>
    </xsl:template>
    <xsl:template match="ead:controlaccess">
        <xsl:for-each select="*">
            <xsl:apply-templates select="."/>
            <xsl:if test="position() ne last()">
                <xsl:text disable-output-escaping="yes">&amp;#10;</xsl:text>
            </xsl:if>
        </xsl:for-each>
    </xsl:template>
    <!-- in rainbows -->
    <xsl:template match="ead:corpname">
        <html:Font html:Color='#0070C0'>
            <xsl:apply-templates/>
        </html:Font>
    </xsl:template>
    <xsl:template match="ead:persname">
        <html:Font html:Color='#7030A0'>
            <xsl:apply-templates/>
        </html:Font>
    </xsl:template>
    <xsl:template match="ead:famname">
        <html:Font html:Color='#ED7D31'>
            <xsl:apply-templates/>
        </html:Font>
    </xsl:template>
    <xsl:template match="ead:geogname">
        <html:Font html:Color='#44546A'>
            <xsl:apply-templates/>
        </html:Font>
    </xsl:template>
    <xsl:template match="ead:genreform">
        <html:Font html:Color='#00B050'>
            <xsl:apply-templates/>
        </html:Font>
    </xsl:template>
    <xsl:template match="ead:subject">
        <html:Font html:Color='#00B0F0'>
            <xsl:apply-templates/>
        </html:Font>
    </xsl:template>
    <xsl:template match="ead:occupation">
        <html:Font html:Color='#FFC000'>
            <xsl:apply-templates/>
        </html:Font>
    </xsl:template>
    <xsl:template match="ead:function">
        <html:Font html:Color='#FF00FF'>
            <xsl:apply-templates/>
        </html:Font>
    </xsl:template>
    <xsl:template match="ead:controlaccess/ead:name">
        <html:Font html:Color='#000000'>
            <xsl:apply-templates/>
        </html:Font>
    </xsl:template>
    
    <xsl:template match="ead:title[not(@render)]">
        <html:Font html:Color='#FF0000'>
            <xsl:apply-templates/>
        </html:Font>
    </xsl:template>
    <xsl:template match="ead:title[@render='bolditalic']">
        <html:B>
            <html:I>
                <html:Font html:Color='#FF0000'>
                    <xsl:apply-templates/>
                </html:Font>
            </html:I>
        </html:B>
    </xsl:template>
    <xsl:template match="ead:title[@render='boldunderline']">
        <html:B>
            <html:U>
                <html:Font html:Color='#FF0000'>
                    <xsl:apply-templates/>
                </html:Font>
            </html:U>
        </html:B>
    </xsl:template>
    <xsl:template match="ead:title[@render='bold']">
        <html:B>
            <html:Font html:Color='#FF0000'>
                <xsl:apply-templates/>
            </html:Font>
        </html:B>
    </xsl:template>
    <xsl:template match="ead:title[@render='italic']">
        <html:I>
            <html:Font html:Color='#FF0000'>
                <xsl:apply-templates/>
            </html:Font>
        </html:I>
    </xsl:template>
    <xsl:template match="ead:title[@render='underline']">
        <html:U>
            <html:Font html:Color='#FF0000'>
                <xsl:apply-templates/>
            </html:Font>
        </html:U>
    </xsl:template>
    
    <xsl:template match="ead:p">
        <xsl:apply-templates/>
        <xsl:if test="position() ne last()">
            <xsl:text disable-output-escaping="yes">&amp;#10;&amp;#10;</xsl:text>
        </xsl:if>
    </xsl:template>
    <xsl:template match="ead:lb">
        <xsl:text disable-output-escaping="yes">&amp;#10;</xsl:text>
    </xsl:template>
    
    <!--pre process to remove any unnnecessary headers added by ASpace (e.g. Scope and Contents note)-->
    <xsl:template match="ead:head">
        <html:Font html:Size="14">
            <xsl:apply-templates/>
        </html:Font>
        <xsl:text disable-output-escaping="yes">&amp;#10;</xsl:text>
    </xsl:template>
    
    <xsl:template match="text()">
        <xsl:value-of select="normalize-unicode(replace(., '\n|\s+', ' '))"/>
    </xsl:template>
    
    <!-- this template provides the framework for the main worksheet, including all of the column headers-->
    <xsl:template match="ead:dsc">
        <Worksheet ss:Name="ContainerList" xmlns="urn:schemas-microsoft-com:office:spreadsheet">
            <Names>
                <NamedRange ss:Name="_FilterDatabase" ss:RefersTo="=ContainerList!R1C1:R16C38"
                    ss:Hidden="1"/>
            </Names>
            <Table ss:ExpandedColumnCount="55" x:FullColumns="1"
                x:FullRows="1" ss:DefaultRowHeight="15">
                <Column ss:AutoFitWidth="0" ss:Width="76"/>
                <Column ss:Width="52" ss:Span="1"/>
                <Column ss:Index="4" ss:AutoFitWidth="0" ss:Width="190"/>
                <Column ss:AutoFitWidth="0" ss:Width="100"/>
                <Column ss:AutoFitWidth="0" ss:Width="62"/>
                <Column ss:AutoFitWidth="0" ss:Width="70"/>
                <Column ss:AutoFitWidth="0" ss:Width="58"/>
                <Column ss:AutoFitWidth="0" ss:Width="70"/>
                <Column ss:AutoFitWidth="0" ss:Width="80"/>
                <Column ss:AutoFitWidth="0" ss:Width="58"/>
                <Column ss:AutoFitWidth="0" ss:Width="90"/>
                <Column ss:AutoFitWidth="0" ss:Width="90"/>
                <Column ss:AutoFitWidth="0" ss:Width="90"/>
                <Column ss:AutoFitWidth="0" ss:Width="70"/>
                <Column ss:AutoFitWidth="0" ss:Width="80"/>
                <Column ss:AutoFitWidth="0" ss:Width="60"/>
                <Column ss:AutoFitWidth="0" ss:Width="70"/>
                <Column ss:AutoFitWidth="0" ss:Width="85" ss:Span="1"/>
                <Column ss:Index="21" ss:AutoFitWidth="0" ss:Width="130"/>
                <Column ss:AutoFitWidth="0" ss:Width="120"/>
                <Column ss:AutoFitWidth="0" ss:Width="88"/>
                <Column ss:AutoFitWidth="0" ss:Width="125"/>
                <Column ss:AutoFitWidth="0" ss:Width="85"/>
                <Column ss:AutoFitWidth="0" ss:Width="100"/>
                <Column ss:AutoFitWidth="0" ss:Width="100"/>
                <Column ss:AutoFitWidth="0" ss:Width="100" ss:Span="4"/>
                <Column ss:Index="33" ss:AutoFitWidth="0" ss:Width="170" ss:Span="1"/>
                <Column ss:Index="35" ss:AutoFitWidth="0" ss:Width="150"/>
                <Column ss:AutoFitWidth="0" ss:Width="110"/>
                <Column ss:AutoFitWidth="0" ss:Width="120"/>
                <Column ss:AutoFitWidth="0" ss:Width="75"/>
                <Column ss:AutoFitWidth="0" ss:Width="105"/>
                <Column ss:AutoFitWidth="0" ss:Width="105"/>
                <Column ss:AutoFitWidth="0" ss:Width="85"/>
                <Column ss:AutoFitWidth="0" ss:Width="115"/>
                <Column ss:AutoFitWidth="0" ss:Width="65"/>
                <Column ss:Index="46" ss:AutoFitWidth="0" ss:Width="60"/>
                <Column ss:AutoFitWidth="0" ss:Width="120"/>
                <Column ss:AutoFitWidth="0" ss:Width="80"/>
                <Column ss:AutoFitWidth="0" ss:Width="95"/>
                <Column ss:AutoFitWidth="0" ss:Width="110"/>
                <Column ss:AutoFitWidth="0" ss:Width="110"/>
                <Column ss:AutoFitWidth="0" ss:Width="165"/>
                <Column ss:AutoFitWidth="0" ss:Width="280"/>
                <!--column headers-->
                <Row ss:AutoFitHeight="0" ss:StyleID="s2">
                    <Cell>
                        <Data ss:Type="String">level number</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">level type</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">unitid</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">title</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">date expression</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">year begin</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="year_begin"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">month begin</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="month_begin"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">day begin</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="day_begin"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">year end</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="year_end"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">month end</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="month_end"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">day end</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="day_end"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">bulk year begin</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="bulk_year_begin"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">bulk month begin</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="bulk_month_begin"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">bulk day begin</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="bulk_day_begin"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">bulk year end</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="bulk_year_end"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">bulk month end</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="bulk_month_end"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">bulk day end</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="bulk_day_end"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">instance type</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="instance_type"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">container 1 type</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="container_1_type"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">container profile</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="container_profile"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">barcode</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="barcode"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">container 1 value / BOX by default</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">container 2 type</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="container_2_type"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">container 2 value / FOLDER by default</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">container 3 type</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="container_3_type"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">container 3 value</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">extent number</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="extent_number"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">extent value</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="extent_value"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">generic extent</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                        <NamedCell ss:Name="generic_extent"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">generic physdesc</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">origination</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">bioghist</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">scope and content note</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">arrangement note</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">access restrictions</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">phystech</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">physloc (location note)</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">use restrictions</Data>
                        <NamedCell ss:Name="_FilterDatabase"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">language code</Data>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">langmaterial note</Data>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">other finding aid</Data>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">custodhist</Data>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">acqinfo</Data>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">appraisal</Data>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">accruals</Data>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">originalsloc</Data>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">alternative form available</Data>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">related material</Data>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">separated material</Data>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">preferred citation</Data>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">processing information</Data>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">control access headings</Data>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">component @id (leave blank, unless value already present)</Data>
                        <NamedCell ss:Name="component_id"/>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">dao link</Data>
                    </Cell>
                    <Cell>
                        <Data ss:Type="String">dao title</Data>
                    </Cell>
                </Row>

                <!-- apply templates for all the components-->

                <xsl:apply-templates select="ead:c | ead:c01"/>

            </Table>
            <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
                <Unsynced/>
                <Print>
                    <ValidPrinterInfo/>
                    <HorizontalResolution>600</HorizontalResolution>
                    <VerticalResolution>600</VerticalResolution>
                </Print>
                <FreezePanes/>
                <FrozenNoSplit/>
                <SplitHorizontal>1</SplitHorizontal>
                <TopRowBottomPane>1</TopRowBottomPane>
                <ActivePane>2</ActivePane>
                <Panes>
                    <Pane>
                        <Number>3</Number>
                    </Pane>
                    <Pane>
                        <Number>2</Number>
                        <ActiveRow>0</ActiveRow>
                    </Pane>
                </Panes>
                <ProtectObjects>False</ProtectObjects>
                <ProtectScenarios>False</ProtectScenarios>
            </WorksheetOptions>
            <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
                <Range>R2C2:R1048576C2</Range>
                <Type>List</Type>
                <Value>LevelValues</Value>
            </DataValidation>
            <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
                <Range>R2C18:R1048576C18</Range>
                <Type>List</Type>
                <Value>InstanceValues</Value>
            </DataValidation>
            <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
                <Range>R2C17:R1048576C17,R2C8:R1048576C8,C11,R2C14:R1048576C14</Range>
                <Type>Whole</Type>
                <Min>1</Min>
                <Max>31</Max>
            </DataValidation>
            <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
                <Range>R2C28:R1048576C28</Range>
                <Type>List</Type>
                <Value>ExtentValues</Value>
            </DataValidation>
            <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
                <Range>R2C39:R1048576C39</Range>
                <Type>List</Type>
                <Value>LanguageCodes</Value>
            </DataValidation>
            <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
                <Range>R2C6:R1048576C6,R2C9:R1048576C9,R2C12:R1048576C12,R2C15:R1048576C15</Range>
                <Type>Whole</Type>
                <Min>0</Min>
                <Max>9999</Max>
            </DataValidation>
            <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
                <Range>R2C7:R1048576C7,R2C10:R1048576C10,R2C13:R1048576C13,R2C16:R1048576C16</Range>
                <Type>Whole</Type>
                <Min>1</Min>
                <Max>12</Max>
            </DataValidation>
            <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
                <Range>R2C23:R1048576C23,R2C25:R1048576C25,R2C19:R1048576C19</Range>
                <Type>List</Type>
                <Value>ContainerValues</Value>
            </DataValidation>
            <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">
                <Range>R2C1:R1048576C1</Range>
                <Type>Whole</Type>
                <Min>0</Min>
                <Max>12</Max>
            </DataValidation>
        </Worksheet>
    </xsl:template>


    <xsl:template name="Excel-Template">
        <xsl:processing-instruction name="mso-application">progid="Excel.Sheet"</xsl:processing-instruction>
        <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:x="urn:schemas-microsoft-com:office:excel"
            xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
            xmlns:html="http://www.w3.org/TR/REC-html40">
            <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">


                <Created>
                    <xsl:value-of select="current-dateTime()"/>
                    <xsl:comment>does the above dateTime show up in the right format? e.g.:
                        2013-03-09T16:16:59Z
                    </xsl:comment>
                </Created>

            </DocumentProperties>
            <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office"/>
            <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
                <WindowHeight>9000</WindowHeight>
                <WindowWidth>23000</WindowWidth>
                <WindowTopX>0</WindowTopX>
                <WindowTopY>0</WindowTopY>
                <ProtectStructure>False</ProtectStructure>
                <ProtectWindows>False</ProtectWindows>
            </ExcelWorkbook>
            <Styles>
                <Style ss:ID="Default" ss:Name="Normal">
                    <Alignment ss:Vertical="Bottom"/>
                    <Borders/>
                    <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
                    <Interior/>
                    <NumberFormat/>
                    <Protection/>
                </Style>
                <!-- gray for level 0 -->
                <Style ss:ID="s1">
                    <Interior ss:Color="#E7E6E6" ss:Pattern="Solid"/>
                </Style>
                <!-- bold styling, for headers -->
                <Style ss:ID="s2">
                    <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"
                        ss:Bold="1"/>
                </Style>
                <!-- generic styling, for all other rows -->
                <Style ss:ID="s3">
                    <Alignment ss:Horizontal="Left" ss:Vertical="Top" ss:WrapText="1"/>
                    <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11"/>
                </Style>
                <!-- add other styles here, or just provide those inline?  hopefully the latter will work-->
            </Styles>
            <Names>
                <NamedRange ss:Name="barcode" ss:RefersTo="=ContainerList!C21"/>
                <NamedRange ss:Name="bulk_day_begin" ss:RefersTo="=ContainerList!C14"/>
                <NamedRange ss:Name="bulk_day_end" ss:RefersTo="=ContainerList!C17"/>
                <NamedRange ss:Name="bulk_month_begin" ss:RefersTo="=ContainerList!C13"/>
                <NamedRange ss:Name="bulk_month_end" ss:RefersTo="=ContainerList!C16"/>
                <NamedRange ss:Name="bulk_year_begin" ss:RefersTo="=ContainerList!C12"/>
                <NamedRange ss:Name="bulk_year_end" ss:RefersTo="=ContainerList!C15"/>
                <NamedRange ss:Name="component_id" ss:RefersTo="=ContainerList!C53"/>
                <NamedRange ss:Name="container_1_type" ss:RefersTo="=ContainerList!C19"/>
                <NamedRange ss:Name="container_2_type" ss:RefersTo="=ContainerList!C23"/>
                <NamedRange ss:Name="container_3_type" ss:RefersTo="=ContainerList!C25"/>
                <NamedRange ss:Name="container_profile" ss:RefersTo="=ContainerList!C20"/>
                <NamedRange ss:Name="ContainerValues" ss:RefersTo="=ControlledVocab!R2C3:R10C3"/>
                <NamedRange ss:Name="day_begin" ss:RefersTo="=ContainerList!C8"/>
                <NamedRange ss:Name="day_end" ss:RefersTo="=ContainerList!C11"/>
                <NamedRange ss:Name="extent_number" ss:RefersTo="=ContainerList!C27"/>
                <NamedRange ss:Name="extent_value" ss:RefersTo="=ContainerList!C28"/>
                <NamedRange ss:Name="ExtentValues" ss:RefersTo="=ControlledVocab!R2C4:R36C4"/>
                <NamedRange ss:Name="generic_extent" ss:RefersTo="=ContainerList!C29"/>
                <NamedRange ss:Name="instance_type" ss:RefersTo="=ContainerList!C18"/>
                <NamedRange ss:Name="InstanceValues" ss:RefersTo="=ControlledVocab!R2C2:R11C2"/>
                <NamedRange ss:Name="LanguageCodes" ss:RefersTo="=ControlledVocab!R2C5:R486C5"/>
                <NamedRange ss:Name="LevelValues" ss:RefersTo="=ControlledVocab!R2C1:R5C1"/>
                <NamedRange ss:Name="month_begin" ss:RefersTo="=ContainerList!C7"/>
                <NamedRange ss:Name="month_end" ss:RefersTo="=ContainerList!C10"/>
                <NamedRange ss:Name="year_begin" ss:RefersTo="=ContainerList!C6"/>
                <NamedRange ss:Name="year_end" ss:RefersTo="=ContainerList!C9"/>
                <NamedRange ss:Name="date_expression" ss:RefersTo="=ContainerList!C5"/>
            </Names>

            <!-- 1st worksheet is created by the description in the DSC-->
            <xsl:apply-templates select="ead:archdesc/ead:dsc[1]"/>

            <!-- 2nd worksheet -->
             <Worksheet ss:Name="ControlledVocab">
     <Table ss:ExpandedColumnCount="6" ss:ExpandedRowCount="486" x:FullColumns="1"
         x:FullRows="1" ss:DefaultRowHeight="15">
      <Column ss:AutoFitWidth="0" ss:Width="112"/>
      <Column ss:AutoFitWidth="0" ss:Width="100"/>
      <Column ss:AutoFitWidth="0" ss:Width="88"/>
      <Column ss:AutoFitWidth="0" ss:Width="80"/>
      <Column ss:AutoFitWidth="0" ss:Width="75"/>
      <Column ss:AutoFitWidth="0" ss:Width="60"/>
   <Row ss:AutoFitHeight="0">
    <Cell ss:StyleID="s2"><Data ss:Type="String">Level values</Data></Cell>
    <Cell ss:StyleID="s2"><Data ss:Type="String">Instance values</Data></Cell>
    <Cell ss:StyleID="s2"><Data ss:Type="String">Container values</Data></Cell>
    <Cell ss:StyleID="s2"><Data ss:Type="String">Extent values</Data></Cell>
    <Cell ss:StyleID="s2"><Data ss:Type="String">Language codes</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="String">series</Data><NamedCell ss:Name="LevelValues"/></Cell>
    <Cell><Data ss:Type="String">Audio</Data><NamedCell ss:Name="InstanceValues"/></Cell>
    <Cell><Data ss:Type="String">Box</Data><NamedCell ss:Name="ContainerValues"/></Cell>
    <Cell><Data ss:Type="String">linear feet</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">aar - Afar</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="String">subseries</Data><NamedCell ss:Name="LevelValues"/></Cell>
    <Cell><Data ss:Type="String">Books</Data><NamedCell ss:Name="InstanceValues"/></Cell>
    <Cell><Data ss:Type="String">Folder</Data><NamedCell ss:Name="ContainerValues"/></Cell>
    <Cell><Data ss:Type="String">gigabytes</Data><NamedCell ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">abk - Abkhazian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="String">file</Data><NamedCell ss:Name="LevelValues"/></Cell>
    <Cell><Data ss:Type="String">Computer disks / tapes</Data><NamedCell
      ss:Name="InstanceValues"/></Cell>
    <Cell><Data ss:Type="String">Reel</Data><NamedCell ss:Name="ContainerValues"/></Cell>
    <Cell><Data ss:Type="String">computer storage media</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">ace - Achinese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="String">item</Data><NamedCell ss:Name="LevelValues"/></Cell>
    <Cell><Data ss:Type="String">Maps</Data><NamedCell ss:Name="InstanceValues"/></Cell>
    <Cell><Data ss:Type="String">Frame</Data><NamedCell ss:Name="ContainerValues"/></Cell>
    <Cell><Data ss:Type="String">computer files</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">ach - Acoli</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="2"><Data ss:Type="String">Microform</Data><NamedCell
      ss:Name="InstanceValues"/></Cell>
    <Cell><Data ss:Type="String">Volume</Data><NamedCell ss:Name="ContainerValues"/></Cell>
    <Cell><Data ss:Type="String">audio cylinders</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">ada - Adangme</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="2"><Data ss:Type="String">Graphic materials</Data><NamedCell
      ss:Name="InstanceValues"/></Cell>
    <Cell><Data ss:Type="String">Oversize Box</Data><NamedCell
      ss:Name="ContainerValues"/></Cell>
    <Cell><Data ss:Type="String">audio discs (CD) </Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">ady - Adyghe; Adygei</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="2"><Data ss:Type="String">Mixed materials</Data><NamedCell
      ss:Name="InstanceValues"/></Cell>
    <Cell><Data ss:Type="String">Oversize Folder</Data><NamedCell
      ss:Name="ContainerValues"/></Cell>
    <Cell><Data ss:Type="String">audio wire reels</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">afa - Afro-Asiatic languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="2"><Data ss:Type="String">Moving images</Data><NamedCell
      ss:Name="InstanceValues"/></Cell>
    <Cell><Data ss:Type="String">Carton</Data><NamedCell ss:Name="ContainerValues"/></Cell>
    <Cell><Data ss:Type="String">audiocassettes</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">afh - Afrihili</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="2"><Data ss:Type="String">Realia</Data><NamedCell
      ss:Name="InstanceValues"/></Cell>
    <Cell><Data ss:Type="String">Case</Data><NamedCell ss:Name="ContainerValues"/></Cell>
    <Cell><Data ss:Type="String">audiotape reels</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">afr - Afrikaans</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="2"><Data ss:Type="String">Text</Data><NamedCell
      ss:Name="InstanceValues"/></Cell>
    <Cell ss:Index="4"><Data ss:Type="String">film cartridges</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">ain - Ainu</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">film cassettes</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">aka - Akan</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="14.4375">
    <Cell ss:Index="4"><Data ss:Type="String">film loops</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">akk - Akkadian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">film reels</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">alb - Albanian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">film reels (8 mm)</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">ale - Aleut</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">film reels (16 mm)</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">alg - Algonquian languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">phonograph records</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">alt - Southern Altai</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">sound track film reels</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">amh - Amharic</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">sound cartridges</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">ang - English, Old (ca.450-1100)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">videocartridges</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">anp - Angika</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">videocassettes</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">apa - Apache languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">videocassettes (VHS)</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">ara - Arabic</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">videocassettes (U-matic)</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">arc - Official Aramaic (700-300 BCE); Imperial Aramaic (700-300 BCE)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">videocassettes (Betacam)</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">arg - Aragonese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">videocassettes (BetacamSP)</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">arm - Armenian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">videocassettes (BetacamSP L)</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">arn - Mapudungun; Mapuche</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">videocassettes (Betamax)</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">arp - Arapaho</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">videocassettes (Video 8)</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">art - Artificial languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">videocassettes (Hi8)</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">arw - Arawak</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">videocassettes (Digital Betacam)</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">asm - Assamese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">videocassettes (MiniDV)</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">ast - Asturian; Bable; Leonese; Asturleonese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">videocassettes (HDCAM)</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">ath - Athapascan languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">videocassettes (DVCAM)</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">aus - Australian languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">videodiscs</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">ava - Avaric</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">videoreels</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">ave - Avestan</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="4"><Data ss:Type="String">see container summary</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">awa - Awadhi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">aym - Aymara</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">aze - Azerbaijani</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bad - Banda languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bai - Bamileke languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bak - Bashkir</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bal - Baluchi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bam - Bambara</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ban - Balinese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">baq - Basque</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bas - Basa</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bat - Baltic languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bej - Beja; Bedawiyet</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bel - Belarusian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bem - Bemba</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ben - Bengali</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ber - Berber languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bho - Bhojpuri</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bih - Bihari languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bik - Bikol</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bin - Bini; Edo</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bis - Bislama</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bla - Siksika</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bnt - Bantu (Other)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bos - Bosnian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bra - Braj</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bre - Breton</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">btk - Batak languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bua - Buriat</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bug - Buginese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bul - Bulgarian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">bur - Burmese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">byn - Blin; Bilin</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cad - Caddo</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cai - Central American Indian languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">car - Galibi Carib</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cat - Catalan; Valencian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cau - Caucasian languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ceb - Cebuano</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cel - Celtic languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cha - Chamorro</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">chb - Chibcha</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">che - Chechen</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">chg - Chagatai</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">chi - Chinese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">chk - Chuukese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">chm - Mari</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">chn - Chinook jargon</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cho - Choctaw</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">chp - Chipewyan; Dene Suline</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">chr - Cherokee</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">chu - Church Slavic; Old Slavonic; Church Slavonic; Old Bulgarian; Old Church Slavonic</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">chv - Chuvash</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">chy - Cheyenne</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cmc - Chamic languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cop - Coptic</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cor - Cornish</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cos - Corsican</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cpe - Creoles and pidgins, English based</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cpf - Creoles and pidgins, French-based </Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cpp - Creoles and pidgins, Portuguese-based </Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cre - Cree</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">crh - Crimean Tatar; Crimean Turkish</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">crp - Creoles and pidgins </Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">csb - Kashubian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cus - Cushitic languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">cze - Czech</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">dak - Dakota</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">dan - Danish</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">dar - Dargwa</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">day - Land Dayak languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">del - Delaware</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">den - Slave (Athapascan)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">dgr - Dogrib</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">din - Dinka</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">div - Divehi; Dhivehi; Maldivian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">doi - Dogri</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">dra - Dravidian languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">dsb - Lower Sorbian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">dua - Duala</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">dum - Dutch, Middle (ca.1050-1350)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">dut - Dutch; Flemish</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">dyu - Dyula</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">dzo - Dzongkha</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">efi - Efik</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">egy - Egyptian (Ancient)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">eka - Ekajuk</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">elx - Elamite</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">eng - English</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">enm - English, Middle (1100-1500)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">epo - Esperanto</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">est - Estonian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ewe - Ewe</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ewo - Ewondo</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">fan - Fang</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">fao - Faroese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">fat - Fanti</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">fij - Fijian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">fil - Filipino; Pilipino</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">fin - Finnish</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">fiu - Finno-Ugrian languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">fon - Fon</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">fre - French</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">frm - French, Middle (ca.1400-1600)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">fro - French, Old (842-ca.1400)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">frr - Northern Frisian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">frs - Eastern Frisian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">fry - Western Frisian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ful - Fulah</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">fur - Friulian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">gaa - Ga</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">gay - Gayo</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">gba - Gbaya</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">gem - Germanic languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">geo - Georgian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ger - German</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">gez - Geez</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">gil - Gilbertese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">gla - Gaelic; Scottish Gaelic</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">gle - Irish</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">glg - Galician</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">glv - Manx</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">gmh - German, Middle High (ca.1050-1500)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">goh - German, Old High (ca.750-1050)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">gon - Gondi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">gor - Gorontalo</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">got - Gothic</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">grb - Grebo</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">grc - Greek, Ancient (to 1453)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">gre - Greek, Modern (1453-)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">grn - Guarani</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">gsw - Swiss German; Alemannic; Alsatian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">guj - Gujarati</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">gwi - Gwich'in</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">hai - Haida</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">hat - Haitian; Haitian Creole</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">hau - Hausa</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">haw - Hawaiian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">heb - Hebrew</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">her - Herero</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">hil - Hiligaynon</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">him - Himachali languages; Western Pahari languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">hin - Hindi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">hit - Hittite</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">hmn - Hmong; Mong</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">hmo - Hiri Motu</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">hrv - Croatian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">hsb - Upper Sorbian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">hun - Hungarian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">hup - Hupa</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">iba - Iban</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ibo - Igbo</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ice - Icelandic</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ido - Ido</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">iii - Sichuan Yi; Nuosu</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ijo - Ijo languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">iku - Inuktitut</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ile - Interlingue; Occidental</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ilo - Iloko</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ina - Interlingua (International Auxiliary Language Association)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">inc - Indic languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ind - Indonesian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ine - Indo-European languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">inh - Ingush</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ipk - Inupiaq</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ira - Iranian languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">iro - Iroquoian languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ita - Italian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">jav - Javanese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">jbo - Lojban</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">jpn - Japanese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">jpr - Judeo-Persian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">jrb - Judeo-Arabic</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kaa - Kara-Kalpak</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kab - Kabyle</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kac - Kachin; Jingpho</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kal - Kalaallisut; Greenlandic</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kam - Kamba</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kan - Kannada</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kar - Karen languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kas - Kashmiri</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kau - Kanuri</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kaw - Kawi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kaz - Kazakh</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kbd - Kabardian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kha - Khasi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">khi - Khoisan languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">khm - Central Khmer</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kho - Khotanese; Sakan</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kik - Kikuyu; Gikuyu</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kin - Kinyarwanda</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kir - Kirghiz; Kyrgyz</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kmb - Kimbundu</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kok - Konkani</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kom - Komi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kon - Kongo</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kor - Korean</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kos - Kosraean</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kpe - Kpelle</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">krc - Karachay-Balkar</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">krl - Karelian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kro - Kru languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kru - Kurukh</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kua - Kuanyama; Kwanyama</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kum - Kumyk</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kur - Kurdish</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">kut - Kutenai</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lad - Ladino</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lah - Lahnda</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lam - Lamba</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lao - Lao</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lat - Latin</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lav - Latvian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lez - Lezghian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lim - Limburgan; Limburger; Limburgish</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lin - Lingala</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lit - Lithuanian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lol - Mongo</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">loz - Lozi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ltz - Luxembourgish; Letzeburgesch</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lua - Luba-Lulua</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lub - Luba-Katanga</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lug - Ganda</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lui - Luiseno</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lun - Lunda</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">luo - Luo (Kenya and Tanzania)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">lus - Lushai</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mac - Macedonian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mad - Madurese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mag - Magahi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mah - Marshallese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mai - Maithili</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mak - Makasar</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mal - Malayalam</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">man - Mandingo</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mao - Maori</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">map - Austronesian languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mar - Marathi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mas - Masai</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">may - Malay</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mdf - Moksha</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mdr - Mandar</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">men - Mende</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mga - Irish, Middle (900-1200)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mic - Mi'kmaq; Micmac</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">min - Minangkabau</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mis - Uncoded languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mkh - Mon-Khmer languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mlg - Malagasy</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mlt - Maltese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mnc - Manchu</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mni - Manipuri</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mno - Manobo languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">moh - Mohawk</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mon - Mongolian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mos - Mossi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mul - Multiple languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mun - Munda languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mus - Creek</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mwl - Mirandese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">mwr - Marwari</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">myn - Mayan languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">myv - Erzya</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nah - Nahuatl languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nai - North American Indian languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nap - Neapolitan</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nau - Nauru</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nav - Navajo; Navaho</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nbl - Ndebele, South; South Ndebele</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nde - Ndebele, North; North Ndebele</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ndo - Ndonga</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nds - Low German; Low Saxon; German, Low; Saxon, Low</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nep - Nepali</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">new - Nepal Bhasa; Newari</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nia - Nias</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nic - Niger-Kordofanian languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">niu - Niuean</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nno - Norwegian Nynorsk; Nynorsk, Norwegian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nob - Bokmål, Norwegian; Norwegian Bokmål</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nog - Nogai</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">non - Norse, Old</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nor - Norwegian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nqo - N'Ko</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nso - Pedi; Sepedi; Northern Sotho</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nub - Nubian languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nwc - Classical Newari; Old Newari; Classical Nepal Bhasa</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nya - Chichewa; Chewa; Nyanja</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nym - Nyamwezi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nyn - Nyankole</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nyo - Nyoro</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">nzi - Nzima</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">oci - Occitan (post 1500); Provençal</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">oji - Ojibwa</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ori - Oriya</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">orm - Oromo</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">osa - Osage</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">oss - Ossetian; Ossetic</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ota - Turkish, Ottoman (1500-1928)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">oto - Otomian languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">paa - Papuan languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">pag - Pangasinan</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">pal - Pahlavi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">pam - Pampanga; Kapampangan</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">pan - Panjabi; Punjabi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">pap - Papiamento</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">pau - Palauan</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">peo - Persian, Old (ca.600-400 B.C.)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">per - Persian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">phi - Philippine languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">phn - Phoenician</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">pli - Pali</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">pol - Polish</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">pon - Pohnpeian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">por - Portuguese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">pra - Prakrit languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">pro - Provençal, Old (to 1500)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">pus - Pushto; Pashto</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">que - Quechua</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">raj - Rajasthani</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">rap - Rapanui</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">rar - Rarotongan; Cook Islands Maori</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">roa - Romance languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">roh - Romansh</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">rom - Romany</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">rum - Romanian; Moldavian; Moldovan</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">run - Rundi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">rup - Aromanian; Arumanian; Macedo-Romanian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">rus - Russian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sad - Sandawe</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sag - Sango</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sah - Yakut</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sai - South American Indian (Other)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sal - Salishan languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sam - Samaritan Aramaic</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">san - Sanskrit</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sas - Sasak</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sat - Santali</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">scn - Sicilian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sco - Scots</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sel - Selkup</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sem - Semitic languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sga - Irish, Old (to 900)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sgn - Sign Languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">shn - Shan</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sid - Sidamo</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sin - Sinhala; Sinhalese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sio - Siouan languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sit - Sino-Tibetan languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sla - Slavic languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">slo - Slovak</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">slv - Slovenian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sma - Southern Sami</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sme - Northern Sami</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">smi - Sami languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">smj - Lule Sami</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">smn - Inari Sami</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">smo - Samoan</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sms - Skolt Sami</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sna - Shona</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">snd - Sindhi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">snk - Soninke</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sog - Sogdian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">som - Somali</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">son - Songhai languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sot - Sotho, Southern</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">spa - Spanish; Castilian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">srd - Sardinian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">srn - Sranan Tongo</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">srp - Serbian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">srr - Serer</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ssa - Nilo-Saharan languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ssw - Swati</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">suk - Sukuma</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sun - Sundanese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sus - Susu</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">sux - Sumerian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">swa - Swahili</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">swe - Swedish</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">syc - Classical Syriac</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">syr - Syriac</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tah - Tahitian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tai - Tai languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tam - Tamil</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tat - Tatar</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tel - Telugu</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tem - Timne</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ter - Tereno</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tet - Tetum</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tgk - Tajik</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tgl - Tagalog</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tha - Thai</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tib - Tibetan</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tig - Tigre</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tir - Tigrinya</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tiv - Tiv</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tkl - Tokelau</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tlh - Klingon; tlhIngan-Hol</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tli - Tlingit</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tmh - Tamashek</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tog - Tonga (Nyasa)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ton - Tonga (Tonga Islands)</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tpi - Tok Pisin</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tsi - Tsimshian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tsn - Tswana</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tso - Tsonga</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tuk - Turkmen</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tum - Tumbuka</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tup - Tupi languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tur - Turkish</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tut - Altaic languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tvl - Tuvalu</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">twi - Twi</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">tyv - Tuvinian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">udm - Udmurt</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">uga - Ugaritic</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">uig - Uighur; Uyghur</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ukr - Ukrainian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">umb - Umbundu</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">und - Undetermined</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">urd - Urdu</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">uzb - Uzbek</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">vai - Vai</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ven - Venda</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">vie - Vietnamese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">vol - Volapük</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">vot - Votic</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">wak - Wakashan languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">wal - Walamo</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">war - Waray</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">was - Washo</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">wel - Welsh</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">wen - Sorbian languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">wln - Walloon</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">wol - Wolof</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">xal - Kalmyk; Oirat</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">xho - Xhosa</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">yao - Yao</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">yap - Yapese</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">yid - Yiddish</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">yor - Yoruba</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">ypk - Yupik languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">zap - Zapotec</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">zbl - Blissymbols; Blissymbolics; Bliss</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">zen - Zenaga</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">zgh - Standard Moroccan Tamazight</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">zha - Zhuang; Chuang</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">znd - Zande languages</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">zul - Zulu</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">zun - Zuni</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">zxx - No linguistic content; Not applicable</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="5"><Data ss:Type="String">zza - Zaza; Dimili; Dimli; Kirdki; Kirmanjki; Zazaki</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Unsynced/>
   <FreezePanes/>
   <FrozenNoSplit/>
   <SplitHorizontal>1</SplitHorizontal>
   <TopRowBottomPane>1</TopRowBottomPane>
   <ActivePane>2</ActivePane>
   <Panes>
    <Pane>
     <Number>3</Number>
    </Pane>
    <Pane>
     <Number>2</Number>
     <ActiveRow>0</ActiveRow>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>

            
            <!-- 3rd worksheet -->
            <Worksheet ss:Name="Original-EAD">
                <Table ss:ExpandedColumnCount="1" ss:ExpandedRowCount="1" x:FullColumns="1"
                    x:FullRows="1" ss:DefaultRowHeight="25">
                    <Column ss:AutoFitWidth="0" ss:Width="800"/>
                    <Row ss:AutoFitHeight="0">
                        <Cell>
                            <Data ss:Type="String">
                                <xsl:value-of select="$ead-copy-filename"/>
                            </Data>
                        </Cell>
                    </Row>
                </Table>
                <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
                    <Unsynced/>
                    <Print>
                        <ValidPrinterInfo/>
                        <HorizontalResolution>600</HorizontalResolution>
                        <VerticalResolution>600</VerticalResolution>
                    </Print>
                    <ProtectObjects>False</ProtectObjects>
                    <ProtectScenarios>False</ProtectScenarios>
                </WorksheetOptions>
            </Worksheet>
        </Workbook>
    </xsl:template>




</xsl:stylesheet>
