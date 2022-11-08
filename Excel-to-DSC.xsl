<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:math="http://www.w3.org/2005/xpath-functions/math"
    xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl" xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:x="urn:schemas-microsoft-com:office:excel"
    xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
    xmlns:html="http://www.w3.org/TR/REC-html40" xmlns:xlink="http://www.w3.org/1999/xlink"
    xmlns:ead="urn:isbn:1-931666-22-9" xmlns:mdc="http://mdc" xmlns="urn:isbn:1-931666-22-9"
    xpath-default-namespace="urn:isbn:1-931666-22-9"
    exclude-result-prefixes="xs math xd o x ss html xlink ead mdc" version="2.0">
    <xd:doc scope="stylesheet">
        <xd:desc>
            <xd:p><xd:b>Created on:</xd:b> December 19, 2013</xd:p>
            <xd:p><xd:b>Significantly revised on:</xd:b> August 18, 2020</xd:p>
            <xd:p><xd:b>Author:</xd:b> Mark Custer</xd:p>
            <xd:p>tested with Saxon-HE 9.6.0.5</xd:p>
        </xd:desc>
    </xd:doc>

    <!-- to do:
  
  recheck how dates are parsed

recheck how origination names are parsed (multiples AND font colors)
        
        update how physdesc mixed content is handled?  (allow genreform, dimensions???)
          
        -->
    <xsl:key name="style-ids_match-for-color" match="ss:Style" use="@ss:ID"/>
    <!-- will probably want to change how this works, but right now you can create mixed content with the following font colors (which results in a pretty hideous rainbow):
        (when there's a second color, that's to deal with the issue of 
        converting a file from XML to XLSX, at which point MS Excel changes the color)
        #FF0000 = title
        #0070C0, #0066CC = corpname
        #7030A0, #666699 = persname  e.g. =('#666699', '#7030A0')
        #ED7D31, #FF6600 = famname
        #44546A, #339966 = geogname
        #00B050, #008080 = genreform     
        #00B0F0, #00CCFF = subject
        #FFC000, #FFCC00 = occupation
        #FF00FF = function
        #000000 = name, but only in the controlaccess column.
        
        italics, underline, bold, etc. (all the emph options) are handled with the other font controls in Excel (e.g. bold -> emph render='bold', etc.)
      -->

    <xsl:output method="xml" indent="yes" encoding="UTF-8"/>
    <xsl:strip-space elements="*"/>
    <xsl:preserve-space elements="*:Data *:Font"/>
    <!--   (1 - 56 / A - BD), columns in Excel
        1   - level number (no default..  requires at least one level-1 value; level-0 values are used for repeating values wihtin the same component; e.g. multiple unitdate expressions)
        2   - level type  (if no value, the level will = "file")
        3   - unitid (ex: 1 (for "Series 1").  if blank, should the transformation auto-number the series and subseries??? add paramters for whether to auto-number, roman vs. arabic numerals, etc.) (did)
        4   - title (did)
  
        5   - date expression (did)
        6   - year begin (did)
        7   - month begin
        8   - day begin
        9   - year end (did)
        10 - month end
        11 - day end
        
        12 - bulk year begin (did) [hidden]
        13 - bulk month begin
        14 - bulk day begin
        15 - bulk year end (did) [hidden]
        16 - bulk month end
        17 - bulk day end
         
        18 - instance type (mixed materials by default) (did)
        19 - container 1 type ("Box" is used if a value is present and the type is blank) (did)
        20 - container profile (did)
        21 - barcode (did)
        22 - container 1 value (did)
        
        23 - container 2 type ("Folder" is used if a value is present and the type is blank) (did)
        24 - container 2 value (did)
        25 - container 3 type ("Carton" is used if a value is present and the type is blank) (did)
        26 - container 3 value (did)
        
        27 - extent number (did)
        28 - extent value (did)
        29 - generic extent statement (did)
        
        30 - generic physdesc statement (allow subelements like dimensions, genreform, and physfacet???) (did)
        
        31 - origination (fyi: @role is NOT supported here) (did)
        32 - bioghist note
        33 - scope and content note 
        34 - arrangement
        
        35 - access restrictions
        36 - phystech
        37 - physloc (did)
        38 - use restrictions
        39 - language code (only 1 allowed, according to AT and ASpace models, and no @script attribute) (did)
        40 - langmaterial (just supports text for now) (did)
        41 - other finding aid
             
        42 - custodial history <custodhist>
        43 - immediate source of aquisition <acqinfo>
        44 - appraisal 
        45 - accruals
             
        46 - location of originals <originalsloc>
        47 - alternative form available
        48 - related material
        49 - separated material 

        50 - preferred citation (discourage use, since it should be automated / inherited)
        51 - process information
        52 - control access (see color coding in the "style-ids_match-for-color" key)
                        
        53 - @id  (herbie, the love bug)
        
        54 - dao link (did) 
        55 - dao title (did)
        
        56 - system id (where we'll park the ASpace URI fragment)
         
         EAD elements/attributes that are NOT currently supported include:
         - daogrp
         - bibliography
         - fileplan
         - index
         - odd (though I should add this somewhere)
         - note
         - @role (in origination/* elements)
         - @script 
         - and a whole lot of other attributes, like @calendar, @certainty, etc.!
         
         ref, extref, etc.  still need to add support for these.
         lists (including chronlist, eventgrp, etc.).
         
         Note, though, that all of these EAD features will still be copied over from the collection-level description, which is where they occur more frequently.  
         These features just aren't supproted currently in the Excel DSC worksheet.

        -->
    
    <xsl:param name="keep-unpublished" select="true()"/>

    <xsl:param name="default-extent-number" select="0" as="xs:decimal"/>
    <xsl:param name="default-extent-type" select="'linear_feet'" as="xs:string"/>
    
    <xsl:variable name="ead-copy-filename"
        select="ss:Workbook/ss:Worksheet[@ss:Name = 'Original-EAD']/ss:Table/ss:Row[1]/ss:Cell/ss:Data"/>

    <xsl:function name="mdc:get-column-number" as="xs:integer">
        <xsl:param name="position"/>
        <xsl:param name="current-index"/>
        <xsl:param name="previous-index"/>
        <xsl:param name="cells-before-previous-index"/>
        <xsl:sequence
            select="
                if ($current-index) then
                    $current-index
                else
                    if ($previous-index) then
                        $cells-before-previous-index + $previous-index + 1
                    else
                        $position"
        />
    </xsl:function>


    <xsl:template match="ss:Workbook">
        <!-- should I change an existing archdesc to be unpublished if the new param is set?? -->
        <xsl:param name="workbook" select="." as="node()"/>
        <xsl:if test="not(ss:Worksheet[@ss:Name = 'ContainerList'])">
            <xsl:message terminate="yes">
               <xsl:text>Woops.  No Worksheet named ContainerList, so we can't run this file.</xsl:text>
            </xsl:message>
        </xsl:if>
        <xsl:choose>
            <xsl:when test="$ead-copy-filename ne ''">
                <xsl:for-each select="document($ead-copy-filename)">
                    <xsl:apply-templates select="@* | node()" mode="ead-copy">
                        <xsl:with-param name="workbook" select="$workbook" tunnel="yes"/>
                    </xsl:apply-templates>
                </xsl:for-each>
            </xsl:when>
            <xsl:otherwise>
                <ead xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                    xsi:schemaLocation="urn:isbn:1-931666-22-9 http://www.loc.gov/ead/ead.xsd">
                    <eadheader>
                        <eadid/>
                        <filedesc>
                            <titlestmt>
                                <titleproper/>
                            </titlestmt>
                        </filedesc>
                    </eadheader>
                    <archdesc level="collection">
                        <xsl:if test="$keep-unpublished eq true()">
                            <xsl:attribute name="audience" select="'internal'"/>
                        </xsl:if>
                        <did>
                            <unitid>
                                <!--AT can only accept 20 characters as the unitid, so that's exactly what the following will provide-->
                                <xsl:value-of
                                    select="concat('temp', substring(string(current-dateTime()), 1, 16))"
                                />
                            </unitid>
                            <unitdate>undated</unitdate>
                            <unittitle>collection title</unittitle>
                            <physdesc>
                                <extent>
                                    <xsl:value-of select="concat($default-extent-number, ' ', $default-extent-type)"/>
                                </extent>
                            </physdesc>
                            <langmaterial>
                                <language langcode="eng"/>
                            </langmaterial>
                        </did>
                        <!-- right now, this will only process a worksheet that has a name of "ContainerList".  if you need multiple DSCs, this would help,
                    but it might be better to change the predicate in the following XPath expression to [1], thereby ensuring a single DSC...  and if someone
                    renamed the first worksheet, it would still be processed-->
                        <xsl:apply-templates
                            select="ss:Worksheet[@ss:Name = 'ContainerList']/ss:Table"/>
                    </archdesc>
                </ead>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <!-- adding the identity template, so we can use the source EAD files during roundtripping-->
    <xsl:template match="@* | node()" mode="ead-copy">
        <xsl:copy>
            <xsl:apply-templates select="@* | node()" mode="#current"/>
        </xsl:copy>
    </xsl:template>

    <xsl:template match="ead:archdesc" mode="ead-copy">
        <xsl:param name="workbook" as="node()" tunnel="yes"/>
        <xsl:copy>
            <xsl:apply-templates select="@* | node() except ead:dsc" mode="#current"/>
            <xsl:apply-templates
                select="$workbook/ss:Worksheet[@ss:Name = 'ContainerList']/ss:Table"/>
        </xsl:copy>
    </xsl:template>

    <xsl:template match="ss:Table">
        <!-- date_expression might be the only name that's really required, but let's be strict -->
        <xsl:variable name="named-cells-required" select="(
            'date_expression',
            'year_begin','month_begin','day_begin','year_end','month_end','day_end',
            'bulk_year_begin','bulk_month_begin','bulk_day_begin','bulk_year_end','bulk_month_end','bulk_day_end',
            'instance_type','container_1_type','container_profile','barcode',
            'container_2_type','container_3_type',
            'extent_number','extent_value','generic_extent',
            'component_id', 'system_id'
            )">
        </xsl:variable>
        <xsl:variable name="named-cells-present" select="ss:Row[1]/ss:Cell/ss:NamedCell[not(matches(@ss:Name, '^_'))]/@ss:Name/string()"/>
        <dsc>
            <!-- and also be strict about the order -->
            <xsl:if test="not(deep-equal($named-cells-present, $named-cells-required))">
                <xsl:message terminate="yes">
                    <xsl:text>The following named cells are NOT present in your Excel template: </xsl:text>
                    <xsl:sequence select="string-join(for $name in $named-cells-required return 
                        $name[not($name[. = ($named-cells-present)])], '; ')"/>
                </xsl:message>
            </xsl:if>
            <xsl:apply-templates select="ss:Row[ss:Cell[1]/ss:Data eq '1']"/>
        </dsc>
    </xsl:template>

    <xsl:template match="ss:Row[ss:Cell/ss:Data]">
        <xsl:param name="depth" select="ss:Cell[1]/ss:Data" as="xs:integer"/>
        <xsl:param name="following-depth"
            select="
                if (following-sibling::ss:Row[ss:Cell[1]/ss:Data ne '0'][1])
                then
                    following-sibling::ss:Row[ss:Cell[1]/ss:Data ne '0'][1]/ss:Cell[1]/ss:Data
                else
                    0"
            as="xs:integer"/>
        <xsl:param name="level"
            select="
                if (not(matches(ss:Cell[2]/ss:Data, '^(series|subseries|file|item|accession)$'))) then
                    'file'
                else
                    ss:Cell[2]/ss:Data/text()
                    (: in other words, if the second column of the row is blank, then 'file' will be used as the @level type by default :)"
            as="xs:string"/>

        <!-- so that rows will NOT be skipped in the case that they jump levels, e.g. 1, 3, rather than 1, 2, let's halt the whole thing -->
        <xsl:if test="$following-depth gt ($depth + 1)">
            <xsl:message terminate="yes">
                <xsl:text>This spreadsheet does not include a proper hierarchy, jumping from </xsl:text>
                <xsl:value-of select="$depth"/> 
                <xsl:text> to </xsl:text>
                <xsl:value-of select="$following-depth"/> 
                <xsl:text> at Row </xsl:text>
                <xsl:value-of select="count(preceding-sibling::ss:Row) + 2"/>
                <xsl:text>. In order to ensure that all rows are properly transformed, you must fix this issue before converting this Excel file to EAD.</xsl:text>
            </xsl:message>
        </xsl:if>
        
        <!-- should I add an option to use c elements OR ennumerated components?  this would be simple to do, but it would require a slightly longer style sheet.-->
        <c>
            <xsl:if test="$keep-unpublished eq true()">
                <xsl:attribute name="audience" select="'internal'"/>
            </xsl:if>
            <xsl:attribute name="level">
                <xsl:value-of
                    select="
                        if ($level = 'accession') then
                            'otherlevel'
                        else
                            $level"
                />
            </xsl:attribute>
            <xsl:if test="$level = 'accession'">
                <xsl:attribute name="otherlevel">
                    <xsl:text>accesssion</xsl:text>
                </xsl:attribute>
            </xsl:if>
            <!-- this next part grabs the @id attribute from column 53, if there is one-->
            <xsl:if
                test="ss:Cell[ss:NamedCell/@ss:Name = 'component_id'][ss:Data/normalize-space()]">
                <xsl:attribute name="id">
                    <xsl:value-of
                        select="ss:Cell[ss:NamedCell/@ss:Name = 'component_id'][1]/ss:Data/normalize-space()"
                    />
                </xsl:attribute>
            </xsl:if>
            <!-- this next part grabs the @altrender attribute from column 56, if there is one-->
            <xsl:if
                test="ss:Cell[ss:NamedCell/@ss:Name = 'system_id'][ss:Data/normalize-space()]">
                <xsl:attribute name="altrender">
                    <xsl:value-of
                        select="ss:Cell[ss:NamedCell/@ss:Name = 'system_id'][1]/ss:Data/normalize-space()"
                    />
                </xsl:attribute>
            </xsl:if>
            <did>
                <xsl:apply-templates mode="did"/>
                <!-- this grabs all of the fields that we allow to repeat via "level 0" in the did node.-->
                <xsl:if test="following-sibling::ss:Row[1][ss:Cell[1]/ss:Data eq '0']">
                    <xsl:for-each-group select="following-sibling::ss:Row[ss:Cell/ss:Data]"
                        group-adjacent="ss:Cell[1]/ss:Data eq '0'">
                        <xsl:variable name="group-position" select="position()"/>
                        <xsl:for-each select="current-group()">
                            <xsl:if test="$group-position eq 1">
                                <xsl:apply-templates select="." mode="did"/>
                            </xsl:if>
                        </xsl:for-each>
                    </xsl:for-each-group>
                </xsl:if>
            </did>
            <xsl:apply-templates mode="non-did"/>

            <!-- this grabs all of the fields that we allow to repeat via "level 0".-->
            <xsl:if test="following-sibling::ss:Row[1][ss:Cell[1]/ss:Data eq '0']">
                <xsl:for-each-group select="following-sibling::ss:Row[ss:Cell/ss:Data]"
                    group-adjacent="ss:Cell[1]/ss:Data eq '0'">
                    <xsl:variable name="group-position" select="position()"/>
                    <xsl:for-each select="current-group()">
                        <xsl:if test="$group-position eq 1">
                            <xsl:apply-templates select="." mode="non-did"/>
                        </xsl:if>
                    </xsl:for-each>
                </xsl:for-each-group>
            </xsl:if>

            <!-- I feel like I should be able to do this by group-ending-with the current depth,
            but i might've messed something up since i couldn't get it to work as expected. -->
            <xsl:if test="$following-depth eq $depth + 1">
                <!--
                    this works, in about 200 seconds, for one of the large test files.  
                    that's a big improvement, but let's try one more thing.
                    and it looks like the newest tactic, that doesn't use "except," only takes 2 seconds for a really large file! much better!
                    
                <xsl:for-each-group
                    select="
                        following-sibling::ss:Row[ss:Cell[1]/ss:Data[1][. ne '0']/normalize-space()]
                        except
                        following-sibling::ss:Row[ss:Cell[1]/ss:Data[xs:integer(.) eq $depth]]/following-sibling::ss:Row"
                    group-by="ss:Cell[1]/ss:Data[1]/xs:integer(.)">
                    <xsl:variable name="group-position" select="position()"/>
                    <xsl:for-each select="current-group()">
                        <xsl:if test="$group-position eq 1">
                            <xsl:apply-templates select="."/>
                        </xsl:if>
                    </xsl:for-each>
                </xsl:for-each-group>
                -->
                <xsl:variable name="depths-left"
                    select="following-sibling::ss:Row/ss:Cell[1]/ss:Data[1][. ne '0'][normalize-space()]/xs:integer(.)"/>
                <xsl:variable name="group-until"
                    select="
                        if (following-sibling::ss:Row/ss:Cell[1]/ss:Data[1][xs:integer(.) eq $depth]) then
                            index-of($depths-left, $depth)[1]
                        else
                            index-of($depths-left, $depth + 1)[last()]"/>
                <xsl:for-each-group
                    select="subsequence(following-sibling::ss:Row[ss:Cell[1]/ss:Data[1][. ne '0']/normalize-space()], 1, $group-until)"
                    group-by="ss:Cell[1]/ss:Data[1]/xs:integer(.)">
                    <xsl:variable name="group-position" select="position()"/>
                    <xsl:for-each select="current-group()">
                        <xsl:if test="$group-position eq 1">
                            <xsl:apply-templates select="."/>
                        </xsl:if>
                    </xsl:for-each>
                </xsl:for-each-group>
            </xsl:if>
        </c>
    </xsl:template>


    <xsl:template match="ss:Cell[ss:Data[normalize-space()]]" mode="did">
        <xsl:param name="style-id" select="@ss:StyleID"/>
        <xsl:param name="row-id" select="generate-id(..)"/>
        <xsl:variable name="position" select="position()" as="xs:integer"/>
        <xsl:variable name="current-index" select="xs:integer(@ss:Index)"/>
        <xsl:variable name="previous-index"
            select="xs:integer(preceding-sibling::ss:Cell[@ss:Index][1]/@ss:Index)"/>
        <xsl:variable name="cells-before-previous-index"
            select="count(preceding-sibling::ss:Cell[@ss:Index][1]/following-sibling::* intersect preceding-sibling::ss:Cell)"/>
        <xsl:variable name="column-number" as="xs:integer">
            <xsl:value-of
                select="mdc:get-column-number($position, $current-index, $previous-index, $cells-before-previous-index)"
            />
        </xsl:variable>

        <xsl:if
            test="
                $column-number = (3,
                4,
                12,
                22, (: right now, container 1 value is required, via column 22... but to match the ASpace data model, I should change this to container 1 value OR a barcode, which is stored in column 21 :)
                24,
                26,
                28,
                30,
                31,
                37,
                39,
                40,
                54)">
            <xsl:call-template name="did-stuff">
                <xsl:with-param name="column-number" select="$column-number" as="xs:integer"/>
                <xsl:with-param name="style-id" select="$style-id"/>
                <xsl:with-param name="row-id" select="$row-id"/>
            </xsl:call-template>
        </xsl:if>
        <xsl:choose>
            <!-- in other words, column number 5 must be blank (no Cell in the output at all)-->
            <xsl:when test="@ss:Index eq '6'">
                <xsl:call-template name="did-stuff">
                    <xsl:with-param name="column-number" select="$column-number" as="xs:integer"/>
                    <xsl:with-param name="style-id" select="$style-id"/>
                </xsl:call-template>
            </xsl:when>
            <!-- in other words, column 5 isn't blank (has Cell/Data) -->
            <xsl:when test="$column-number eq 5">
                <xsl:call-template name="did-stuff">
                    <xsl:with-param name="column-number" select="$column-number" as="xs:integer"/>
                    <xsl:with-param name="style-id" select="$style-id"/>
                </xsl:call-template>
            </xsl:when>
            <!-- in other words, column 5 isn't entirely blank (it has a Cell, but it doesn't have any Data), so we just use column 6 -->
            <!-- recheck this rule!!!! -->
            <xsl:when test="$column-number eq 6 and ss:NamedCell[@ss:Name = 'year_begin']">
                <xsl:call-template name="did-stuff">
                    <xsl:with-param name="column-number" select="$column-number" as="xs:integer"/>
                    <xsl:with-param name="style-id" select="$style-id"/>
                </xsl:call-template>
            </xsl:when>
        </xsl:choose>
    </xsl:template>


    <xsl:template match="ss:Cell[ss:Data[normalize-space()]]" mode="non-did">
        <xsl:param name="style-id" select="@ss:StyleID"/>
        <xsl:variable name="position" select="position()"/>
        <xsl:variable name="current-index" select="xs:integer(@ss:Index)"/>
        <xsl:variable name="previous-index"
            select="xs:integer(preceding-sibling::ss:Cell[@ss:Index][1]/@ss:Index)"/>
        <xsl:variable name="cells-before-previous-index"
            select="count(preceding-sibling::ss:Cell[@ss:Index][1]/following-sibling::* intersect preceding-sibling::ss:Cell)"/>
        <xsl:variable name="column-number" as="xs:integer">
            <xsl:value-of
                select="mdc:get-column-number($position, $current-index, $previous-index, $cells-before-previous-index)"
            />
        </xsl:variable>
        <xsl:if
            test="
                $column-number = (32 to 36,
                38,
                41 to 52)">
            <xsl:call-template name="non-did-stuff">
                <xsl:with-param name="column-number" select="$column-number" as="xs:integer"/>
                <xsl:with-param name="style-id" select="$style-id"/>
            </xsl:call-template>
        </xsl:if>
    </xsl:template>

    <xsl:template name="did-stuff">
        <xsl:param name="style-id"/>
        <xsl:param name="column-number" as="xs:integer"/>
        <xsl:param name="row-id"/>

        <xsl:choose>
            <xsl:when test="$column-number eq 3">
                <unitid>
                    <xsl:apply-templates/>
                </unitid>
            </xsl:when>
            <xsl:when test="$column-number eq 4">
                <unittitle>
                    <xsl:choose>
                        <!-- 1st test checks to see if the current Cell has a style ID that would indicate that the font is supposed to be red -->
                        <!-- the second test makes sure that the cell and the data don't both have the RED font color specified.  without the "not" statement, two nested title elements might appear in the output. -->
                        <xsl:when
                            test="
                                key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color = '#FF0000'
                                and
                                not(ss:Data//html:Font/@html:Color = '#FF0000')
                                and key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Underline">
                            <title render="underline">
                                <xsl:apply-templates/>
                            </title>
                        </xsl:when>
                        <xsl:when
                            test="
                                key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color = '#FF0000'
                                and
                                not(ss:Data//html:Font/@html:Color = '#FF0000')
                                and key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Italic">
                            <title render="italic">
                                <xsl:apply-templates/>
                            </title>
                        </xsl:when>
                        <xsl:when
                            test="
                                key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color = '#FF0000'
                                and
                                not(ss:Data//html:Font/@html:Color = '#FF0000')">
                            <title>
                                <xsl:apply-templates/>
                            </title>
                        </xsl:when>
                        <xsl:when
                            test="
                                key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color = ('#0070C0', '#0066CC')
                                and
                                not(ss:Data//html:Font/@html:Color = ('#0070C0', '#0066CC'))">
                            <corpname>
                                <xsl:apply-templates/>
                            </corpname>
                        </xsl:when>
                        <xsl:when
                            test="
                                key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color = ('#666699', '#7030A0')
                                and
                                not(ss:Data//html:Font/@html:Color = ('#666699', '#7030A0'))">
                            <persname>
                                <xsl:apply-templates/>
                            </persname>
                        </xsl:when>
                        <xsl:when
                            test="
                                key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color = ('#ED7D31', '#FF6600')
                                and
                                not(ss:Data//html:Font/@html:Color = ('#ED7D31', '#FF6600'))">
                            <famname>
                                <xsl:apply-templates/>
                            </famname>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:apply-templates/>
                        </xsl:otherwise>
                    </xsl:choose>
                </unittitle>
            </xsl:when>
            <!-- there should a better way to deal with dates / other grouped cells -->
            <xsl:when test="$column-number eq 5">
                <!--added some DateTime checking. Might want to add this to all of the date fields -->
                <xsl:variable name="year-begin"
                    select="
                        if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'year_begin'][ss:Data/@ss:Type = 'DateTime'])
                        then
                            following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'year_begin']/ss:Data/year-from-dateTime(.)
                        else
                            if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'year_begin'][ss:Data[normalize-space()]])
                            then
                                following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'year_begin']/format-number(., '0000')
                            else
                                ''"/>
                <xsl:variable name="month-begin"
                    select="
                        if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'month_begin'][ss:Data[normalize-space()]])
                        then
                            concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'month_begin']/format-number(., '00'))
                        else
                            ''"/>
                <xsl:variable name="day-begin"
                    select="
                        if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'day_begin'][ss:Data[normalize-space()]])
                        then
                            concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'day_begin']/format-number(., '00'))
                        else
                            ''"/>
                <xsl:variable name="year-end"
                    select="
                        if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'year_end'][ss:Data/@ss:Type = 'DateTime'])
                        then
                            following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'year_end']/ss:Data/year-from-dateTime(.)
                        else
                            if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'year_end'][ss:Data[normalize-space()]])
                            then
                                following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'year_end']/format-number(., '0000')
                            else
                                ''"/>
                <xsl:variable name="month-end"
                    select="
                        if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'month_end'][ss:Data[normalize-space()]])
                        then
                            concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'month_end']/format-number(., '00'))
                        else
                            ''"/>
                <xsl:variable name="day-end"
                    select="
                        if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'day_end'][ss:Data[normalize-space()]])
                        then
                            concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'day_end']/format-number(., '00'))
                        else
                            ''"/>
                <unitdate type="inclusive">
                    <xsl:if test="$year-begin ne ''">
                        <xsl:attribute name="normal">
                            <xsl:choose>
                                <xsl:when
                                    test="
                                        concat($year-begin, $month-begin, $day-begin) eq concat($year-end, $month-end, $day-end)
                                        or boolean($year-end) eq false()">
                                    <xsl:value-of
                                        select="concat($year-begin, $month-begin, $day-begin)"/>
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:value-of
                                        select="concat($year-begin, $month-begin, $day-begin, '/', $year-end, $month-end, $day-end)"
                                    />
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:attribute>
                    </xsl:if>
                    <xsl:value-of select="."/>
                </unitdate>
            </xsl:when>
            <xsl:when
                test="
                    $column-number eq 6 and
                    not(preceding-sibling::ss:Cell[1][ss:Data/normalize-space()]/ss:NamedCell[@ss:Name = 'date_expression'])">
                <xsl:variable name="year-begin"
                    select="
                        if (ss:Data[@ss:Type = 'DateTime'])
                        then
                            ss:Data/year-from-dateTime(.)
                        else
                            if (ss:Data[normalize-space()]) then
                                format-number(ss:Data, '0000')
                            else
                                ''"/>
                <xsl:variable name="month-begin"
                    select="
                        if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'month_begin'][ss:Data[normalize-space()]])
                        then
                            concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'month_begin']/format-number(., '00'))
                        else
                            ''"/>
                <xsl:variable name="day-begin"
                    select="
                        if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'day_begin'][ss:Data[normalize-space()]])
                        then
                            concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'day_begin']/format-number(., '00'))
                        else
                            ''"/>
                <xsl:variable name="year-end"
                    select="
                        if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'year_end'][ss:Data/@ss:Type = 'DateTime'])
                        then
                            following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'year_end']/ss:Data/year-from-dateTime(.)
                        else
                            if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'year_end'][ss:Data[normalize-space()]])
                            then
                                following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'year_end']/format-number(., '0000')
                            else
                                ''"/>
                <xsl:variable name="month-end"
                    select="
                        if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'month_end'][ss:Data[normalize-space()]])
                        then
                            concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'month_end']/format-number(., '00'))
                        else
                            ''"/>
                <xsl:variable name="day-end"
                    select="
                        if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'day_end'][ss:Data[normalize-space()]])
                        then
                            concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'day_end']/format-number(., '00'))
                        else
                            ''"/>
                <xsl:variable name="date-value">
                    <xsl:choose>
                        <xsl:when
                            test="
                                concat($year-begin, $month-begin, $day-begin) eq concat($year-end, $month-end, $day-end)
                                or boolean($year-end) eq false()">
                            <xsl:value-of select="concat($year-begin, $month-begin, $day-begin)"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:value-of
                                select="concat($year-begin, $month-begin, $day-begin, '/', $year-end, $month-end, $day-end)"
                            />
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:variable>
                <unitdate type="inclusive">
                    <xsl:attribute name="normal">
                        <xsl:value-of select="$date-value"/>
                    </xsl:attribute>
                    <!-- this shouldn't be required, but ASpace's EAD importer with version 1.3 has a bug in it
                        that results in values like "1912-1912" if you leave a date expression out!!! 
                      (in fact, rather than do this in this transformation, I'll add a second transformation that will add in the 
                      text nodes if those are missing.  That shouldn't be necessary, but until we can update the EAD importer, we'll need to 
                      do just that)
                    <xsl:value-of select="translate($date-value, '/', '-')"/>
                    -->
                </unitdate>
            </xsl:when>

            <xsl:when test="$column-number eq 12">
                <xsl:variable name="bulk-year-begin"
                    select="
                        if (ss:Data) then
                            format-number(., '0000')
                        else
                            ''"/>
                <xsl:variable name="bulk-month-begin"
                    select="
                        if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'bulk_month_begin'][ss:Data[normalize-space()]])
                        then
                            concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'bulk_month_begin']/format-number(., '00'))
                        else
                            ''"/>
                <xsl:variable name="bulk-day-begin"
                    select="
                        if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'bulk_day_begin'][ss:Data[normalize-space()]])
                        then
                            concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'bulk_day_begin']/format-number(., '00'))
                        else
                            ''"/>
                <xsl:variable name="bulk-year-end"
                    select="
                        if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'bulk_year_end'][ss:Data[normalize-space()]])
                        then
                            following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'bulk_year_end']/format-number(., '0000')
                        else
                            ''"/>
                <xsl:variable name="bulk-month-end"
                    select="
                        if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'bulk_month_end'][ss:Data[normalize-space()]])
                        then
                            concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'bulk_month_end']/format-number(., '00'))
                        else
                            ''"/>
                <xsl:variable name="bulk-day-end"
                    select="
                        if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'bulk_day_end'][ss:Data[normalize-space()]])
                        then
                            concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'bulk_day_end']/format-number(., '00'))
                        else
                            ''"/>
                <unitdate type="bulk">
                    <xsl:attribute name="normal">
                        <xsl:choose>
                            <xsl:when
                                test="
                                    concat($bulk-year-begin, $bulk-month-begin, $bulk-day-begin) eq concat($bulk-year-end, $bulk-month-end, $bulk-day-end)
                                    or boolean($bulk-year-end) eq false()">
                                <xsl:value-of
                                    select="concat($bulk-year-begin, $bulk-month-begin, $bulk-day-begin)"
                                />
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of
                                    select="concat($bulk-year-begin, $bulk-month-begin, $bulk-day-begin, '/', $bulk-year-end, $bulk-month-end, $bulk-day-end)"
                                />
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:attribute>
                </unitdate>
            </xsl:when>


            <xsl:when test="$column-number eq 22">
                <!-- label should be column 18.  If empty, though, just choose Mixed materials-->
                <xsl:variable name="instance_type"
                    select="
                        if (preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'instance_type'][ss:Data[normalize-space()]])
                        then
                            preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'instance_type']/ss:Data
                        else
                            'mixed_materials'"/>
                <xsl:variable name="barcode"
                    select="preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'barcode']/ss:Data"/>

                <container id="{$row-id}">
                    <xsl:attribute name="label">
                        <xsl:value-of
                            select="
                                if ($barcode ne '') then
                                    concat($instance_type, ' (', $barcode, ')')
                                else
                                    $instance_type"
                        />
                    </xsl:attribute>
                    <xsl:attribute name="type">
                        <xsl:value-of
                            select="
                                if (preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'container_1_type'][ss:Data[normalize-space()]])
                                then
                                    preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'container_1_type']/translate(lower-case(normalize-space(ss:Data)), ' ', '_')
                                else
                                    'box'"
                        />
                    </xsl:attribute>
                    <xsl:if
                        test="preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'container_profile'][ss:Data[normalize-space()]]">
                        <xsl:attribute name="altrender">
                            <xsl:value-of
                                select="preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'container_profile']/ss:Data"
                            />
                        </xsl:attribute>
                    </xsl:if>
                    <xsl:apply-templates/>
                </container>
            </xsl:when>

            <xsl:when test="$column-number eq 24">
                <container parent="{$row-id}">
                    <xsl:attribute name="type">
                        <xsl:value-of
                            select="
                                if (preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'container_2_type'][ss:Data[normalize-space()]])
                                then
                                    preceding-sibling::ss:Cell[1]/translate(lower-case(normalize-space(ss:Data)), ' ', '_')
                                else
                                    'folder'"
                        />
                    </xsl:attribute>
                    <xsl:apply-templates/>
                </container>
            </xsl:when>

            <xsl:when test="$column-number eq 26">
                <container parent="{$row-id}">
                    <xsl:attribute name="type">
                        <xsl:value-of
                            select="
                                if (preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'container_3_type'][ss:Data[normalize-space()]])
                                then
                                    preceding-sibling::ss:Cell[1]/translate(lower-case(normalize-space(ss:Data)), ' ', '_')
                                else
                                    'carton'"
                        />
                    </xsl:attribute>
                    <xsl:apply-templates/>
                </container>
            </xsl:when>

            <xsl:when test="$column-number eq 28">
                <physdesc>
                    <xsl:variable name="extent-number">
                        <xsl:value-of
                            select="
                                if (preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'extent_number'][ss:Data[normalize-space()]])
                                then
                                    preceding-sibling::ss:Cell[1]/ss:Data
                                else
                                    'noextent'"
                        />
                    </xsl:variable>
                    <extent>
                        <xsl:value-of
                            select="
                                if ($extent-number castable as xs:double) then
                                    concat(format-number($extent-number, '0.##'), ' ', .)
                                else
                                    if ($extent-number ne 'noextent') then
                                        concat($extent-number, ' ', .)
                                    else
                                        '0 See container summary'"
                        />
                    </extent>
                    <xsl:if
                        test="following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'generic_extent'][ss:Data[normalize-space()]]">
                        <extent>
                            <xsl:apply-templates
                                select="following-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'generic_extent']"
                            />
                        </extent>
                    </xsl:if>
                </physdesc>
            </xsl:when>

            <xsl:when test="$column-number eq 30">
                <physdesc>
                    <xsl:apply-templates/>
                </physdesc>
            </xsl:when>

            <xsl:when test="$column-number eq 31">
                <origination>
                    <xsl:choose>
                        <xsl:when
                            test="
                                key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color = ('#0070C0', '#0066CC')
                                and
                                not(ss:Data/html:Font/@html:Color = ('#0070C0', '#0066CC'))">
                            <corpname>
                                <xsl:apply-templates/>
                            </corpname>
                        </xsl:when>
                        <xsl:when
                            test="
                                key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color = ('#666699', '#7030A0')
                                and
                                not(ss:Data/html:Font/@html:Color = ('#666699', '#7030A0'))">
                            <persname>
                                <xsl:apply-templates/>
                            </persname>
                        </xsl:when>
                        <xsl:when
                            test="
                                key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color = ('#ED7D31', '#FF6600')
                                and
                                not(ss:Data/html:Font/@html:Color = ('#ED7D31', '#FF6600'))">
                            <famname>
                                <xsl:apply-templates/>
                            </famname>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:apply-templates/>
                        </xsl:otherwise>
                    </xsl:choose>
                </origination>
            </xsl:when>

            <xsl:when test="$column-number eq 37">
                <physloc>
                    <xsl:apply-templates/>
                </physloc>
            </xsl:when>

            <xsl:when test="$column-number eq 39">
                <langmaterial>
                    <language
                        langcode="{if (contains(., '-')) then substring-before(., ' -') else .}"/>
                </langmaterial>
            </xsl:when>

            <xsl:when test="$column-number eq 40">
                <langmaterial>
                    <xsl:apply-templates/>
                </langmaterial>
            </xsl:when>

            <!-- 54 and 55 -->
            <xsl:when test="$column-number eq 54">
                <dao xlink:type="simple">
                    <xsl:attribute name="href" namespace="http://www.w3.org/1999/xlink">
                        <xsl:value-of select="normalize-space()"/>
                    </xsl:attribute>
                    <xsl:if test="following-sibling::ss:Cell">
                        <xsl:attribute name="title" namespace="http://www.w3.org/1999/xlink">
                            <xsl:value-of select="following-sibling::ss:Cell[1]"/>
                        </xsl:attribute>
                    </xsl:if>
                </dao>
            </xsl:when>

        </xsl:choose>
    </xsl:template>

    <xsl:template name="non-did-stuff">
        <xsl:param name="column-number"/>
        <xsl:param name="style-id"/>
        <!-- 32 to 36, 38, 41 to 52 -->
        <xsl:variable name="element-name"
            select="
                if ($column-number eq 32) then
                    'bioghist'
                else
                    if ($column-number eq 33) then
                        'scopecontent'
                    else
                        if ($column-number eq 34) then
                            'arrangement'
                        else
                            if ($column-number eq 35) then
                                'accessrestrict'
                            else
                                if ($column-number eq 36) then
                                    'phystech'
                                else
                                    if ($column-number eq 38) then
                                        'userestrict'
                                    else
                                        if ($column-number eq 41) then
                                            'otherfindaid'
                                        else
                                            if ($column-number eq 42) then
                                                'custodhist'
                                            else
                                                if ($column-number eq 43) then
                                                    'acqinfo'
                                                else
                                                    if ($column-number eq 44) then
                                                        'appraisal'
                                                    else
                                                        if ($column-number eq 45) then
                                                            'accruals'
                                                        else
                                                            if ($column-number eq 46) then
                                                                'originalsloc'
                                                            else
                                                                if ($column-number eq 47) then
                                                                    'altformavail'
                                                                else
                                                                    if ($column-number eq 48) then
                                                                        'relatedmaterial'
                                                                    else
                                                                        if ($column-number eq 49) then
                                                                            'separatedmaterial'
                                                                        else
                                                                            if ($column-number eq 50) then
                                                                                'prefercite'
                                                                            else
                                                                                if ($column-number eq 51) then
                                                                                    'processinfo'
                                                                                else
                                                                                    if ($column-number eq 52) then
                                                                                        'controlaccess'
                                                                                    else
                                                                                        'nada'"/>
        <xsl:choose>
            <xsl:when test="$element-name eq 'nada' or normalize-space(.) eq ''"/>
            <xsl:otherwise>
                <xsl:element name="{$element-name}" namespace="urn:isbn:1-931666-22-9">
                    <xsl:apply-templates>
                        <xsl:with-param name="column-number" select="$column-number" as="xs:integer"/>
                        <xsl:with-param name="style-id" select="$style-id"/>
                    </xsl:apply-templates>
                </xsl:element>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>


    <xsl:template match="ss:Data">
        <xsl:param name="column-number"/>
        <xsl:param name="style-id"/>
        <xsl:choose>
            <!-- controlaccess stuff, when sub-elements like Font are present -->
            <xsl:when test="number($column-number) = (52) and *">
                <xsl:apply-templates select="*[normalize-space()]">
                    <xsl:with-param name="column-number" select="$column-number"/>
                </xsl:apply-templates>
            </xsl:when>

            <xsl:when test="number($column-number) = (52) and not(*)">
                <xsl:choose>
                    <xsl:when
                        test="
                        key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color = ('#00B0F0', '#00CCFF')
                        and
                        not(ss:Data/html:Font/@html:Color = ('#00B0F0', '#00CCFF'))">
                        <subject>
                            <xsl:apply-templates/>
                        </subject>
                    </xsl:when>
                    <xsl:when
                        test="
                            key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color = ('#0070C0', '#0066CC')
                            and
                            not(ss:Data/html:Font/@html:Color = ('#0070C0', '#0066CC'))">
                        <corpname>
                            <xsl:apply-templates/>
                        </corpname>
                    </xsl:when>
                    <xsl:when
                        test="
                            key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color = ('#666699', '#7030A0')
                            and
                            not(ss:Data/html:Font/@html:Color = ('#666699', '#7030A0'))">
                        <persname>
                            <xsl:apply-templates/>
                        </persname>
                    </xsl:when>
                    <xsl:when
                        test="
                            key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color = ('#ED7D31', '#FF6600')
                            and
                            not(ss:Data/html:Font/@html:Color = ('#ED7D31', '#FF6600'))">
                        <famname>
                            <xsl:apply-templates/>
                        </famname>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:apply-templates/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>

            <!-- hack way to deal with adding <head> elements for scope and content and other types of notes.-->
            <!-- also gotta check style ids, since if you re-save an Excel file, it'll strip the font element out and replace it with an ID :( -->
            
            <!-- pontential encoding:
                    <Cell ss:Index="42" ss:StyleID="s61">
                    <ss:Data ss:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40"><Font html:Size="14"
       html:Color="#000000">Provenance&#10;</Font>
       <Font html:Color="#000000">Has an ink stamp</Font>
       </ss:Data>
       </Cell>

  -->
            <xsl:when test="starts-with(*[2], '&#10;') and not(html:Font[1]/@html:Size eq '14')">
                <head>
                    <xsl:apply-templates select="*[1]"/>
                </head>
                <p>
                    <xsl:apply-templates select="node() except *[1]"/>
                </p>
            </xsl:when>

            <xsl:when test="contains(text()[1], '&#10;') and html:Font[1]/@html:Size eq '14'">
                <xsl:apply-templates select="*[1]"/>
                <p>
                    <xsl:apply-templates select="node() except *[1]"/>
                </p>
            </xsl:when>
            
            
            <xsl:when test="html:Font[1]/@html:Size eq '14'">
                <xsl:apply-templates select="*[1]"/>
                <p>
                    <xsl:apply-templates select="node() except *[1]"/>
                </p>
            </xsl:when>
            

            <xsl:when
                test="
                    key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Size eq '14'
                    and html:Font[@html:Size = '11'][1]/starts-with(., '&#10;')
                    and not(html:Font[1]/@html:Size eq '14')">
                <head>
                    <xsl:apply-templates select="text()[1]"/>
                </head>
                <p>
                    <xsl:apply-templates select="node() except text()[1]"/>
                </p>
            </xsl:when>

            <xsl:when test="starts-with(text()[1], '&#10;')">
                <xsl:apply-templates select="text()[1]"/>
                <p>
                    <xsl:apply-templates select="node() except text()[1]"/>
                </p>
            </xsl:when>

            <!-- 32 to 36, 38, 41 to 52 -->
            <xsl:when
                test="
                    number($column-number) = (32 to 36,
                    38,
                    41 to 51)">
                <p>
                    <xsl:apply-templates/>
                </p>
            </xsl:when>
            <xsl:when test="contains(., '&#10;&#10;')">
                <p>
                    <xsl:apply-templates/>
                </p>
            </xsl:when>
            <xsl:otherwise>
                <xsl:apply-templates/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <!-- Still need to ensure that ALL of the emph @render options work
        when that text is the only content of the Cell.
    
    render='nonproport' requires use of "Courier New"
    
   (why doesn't EAD have bolditalicunderline?)  
    
    -->
    <xsl:template match="html:B[not(*)][normalize-space()]">
        <emph render="bold">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:B[parent::html:U][not(*)][normalize-space()]" priority="3">
        <emph render="boldunderline">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:U[parent::html:B][not(*)][normalize-space()]" priority="2">
        <emph render="boldunderline">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:I[parent::html:B][not(*)][normalize-space()]" priority="3">
        <emph render="bolditalic">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:B[parent::html:I][not(*)][normalize-space()]" priority="2">
        <emph render="bolditalic">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:B[parent::html:Font/@html:Size = '8'][not(*)][normalize-space()]" priority="2">
        <emph render="boldsmcaps">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>


    <!-- also need to account for this, though:
        <I><Font html:Color="#000000">See also</Font></I> 
    -->
    <xsl:template match="html:I[not(*)][normalize-space()]">
        <emph render="italic">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:U[not(*)][normalize-space()]">
        <emph render="underline">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:Sup[normalize-space()]">
        <emph render="super">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:Sub[normalize-space()]">
        <emph render="sub">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:Font[@html:Face = 'Courier New'][normalize-space()]">
        <emph render="nonproport">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>


    <xsl:template match="html:Font[@html:Size = '8'][parent::html:B][not(*)][normalize-space()]" priority="2">
        <emph render="boldsmcaps">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:Font[@html:Size = '8'][not(*)][normalize-space()]">
        <emph render="smcaps">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:Font[@html:Size = '14'][normalize-space()]">
        <head>
            <xsl:apply-templates/>
        </head>
    </xsl:template>

    <xsl:template match="*:Font[@html:Color = ('#0070C0', '#0066CC')][normalize-space()]">
        <corpname>
            <xsl:apply-templates/>
        </corpname>
    </xsl:template>
    <xsl:template match="*:Font[@html:Color = ('#666699', '#7030A0')][normalize-space()]">
        <persname>
            <xsl:apply-templates/>
        </persname>
    </xsl:template>
    <xsl:template match="*:Font[@html:Color = ('#ED7D31', '#FF6600')][normalize-space()]">
        <famname>
            <xsl:apply-templates/>
        </famname>
    </xsl:template>
    <xsl:template match="*:Font[@html:Color = ('#44546A', '#339966')][normalize-space()]">
        <geogname>
            <xsl:apply-templates/>
        </geogname>
    </xsl:template>
    <xsl:template match="*:Font[@html:Color = ('#00B050', '#008080')][normalize-space()]">
        <genreform>
            <xsl:apply-templates/>
        </genreform>
    </xsl:template>
    <xsl:template match="*:Font[@html:Color = ('#00B0F0', '#00CCFF')][normalize-space()]">
        <subject>
            <xsl:apply-templates/>
        </subject>
    </xsl:template>
    <xsl:template match="*:Font[@html:Color = ('#FFC000', '#FFCC00')][normalize-space()]">
        <occupation>
            <xsl:apply-templates/>
        </occupation>
    </xsl:template>
    <xsl:template match="*:Font[@html:Color = '#FF00FF'][normalize-space()]">
        <function>
            <xsl:apply-templates/>
        </function>
    </xsl:template>
    <xsl:template match="*:Font[@html:Color = '#000000'][not(@html:Size = '14')][normalize-space()]" priority="2">
        <xsl:param name="column-number"/>
        <xsl:choose>
            <xsl:when test="number($column-number) eq 52">
                <name>
                    <xsl:apply-templates/>
                </name>
            </xsl:when>
            <xsl:otherwise>
                <xsl:apply-templates/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <xsl:template match="*:Font[@html:Color = '#FF0000'][normalize-space()]">
        <xsl:param name="column-number"/>
        <xsl:choose>
            <xsl:when test=".[parent::html:I/parent::html:B]">
                <title render="bolditalic">
                    <xsl:apply-templates/>
                </title>
            </xsl:when>
            <xsl:when test=".[parent::html:I]">
                <title render="italic">
                    <xsl:apply-templates/>
                </title>
            </xsl:when>
            <xsl:when test=".[parent::html:B]">
                <title render="bold">
                    <xsl:apply-templates/>
                </title>
            </xsl:when>
            <xsl:when test=".[parent::html:U]">
                <title render="underline">
                    <xsl:apply-templates/>
                </title>
            </xsl:when>
            <xsl:when test="number($column-number) eq 52">
                <title>
                    <xsl:apply-templates/>
                </title>
            </xsl:when>
            <xsl:otherwise>
                <title>
                    <xsl:apply-templates/>
                </title>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <!-- I don't like doing this, but I'm not sure of a better way to create multiple paragaphs right now -->
    <xsl:template match="text()">
        <xsl:choose>
            <xsl:when test="contains(., '&#10;&#10;')">
                <xsl:call-template name="create-paragraph-from-text"/>
            </xsl:when>
            <xsl:when test="contains(., '&#10;')">
                <xsl:call-template name="create-line-break-from-text"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="."/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <xsl:template name="create-paragraph-from-text">
        <xsl:for-each select="tokenize(., '&#10;&#10;')">
            <xsl:choose>
                <xsl:when test="contains(., '&#10;')">
                    <xsl:call-template name="create-line-break-from-text"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="."/>
                </xsl:otherwise>
            </xsl:choose>
            <xsl:if test="position() ne last()">
                <xsl:text disable-output-escaping="yes">&lt;/p&gt;&lt;p&gt;</xsl:text>
            </xsl:if>
        </xsl:for-each>
    </xsl:template>
    
    <xsl:template name="create-line-break-from-text">
        <xsl:for-each select="tokenize(., '&#10;')">
            <xsl:value-of select="."/>
            <xsl:if test="position() ne last()">
                <xsl:element name="lb" namespace="urn:isbn:1-931666-22-9"/>
            </xsl:if>
        </xsl:for-each>
    </xsl:template>
    
</xsl:stylesheet>
