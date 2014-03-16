<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:math="http://www.w3.org/2005/xpath-functions/math"
    xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl" xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:x="urn:schemas-microsoft-com:office:excel"
    xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
    xmlns:html="http://www.w3.org/TR/REC-html40" xmlns:xlink="http://www.w3.org/1999/xlink"
    xmlns:mdc="http://mdc" xmlns="urn:isbn:1-931666-22-9"
    exclude-result-prefixes="xs math xd o x ss html" version="2.0">
    <xd:doc scope="stylesheet">
        <xd:desc>
            <xd:p><xd:b>Created on:</xd:b> December 19, 2013</xd:p>
            <xd:p><xd:b>Author:</xd:b> Mark Custer</xd:p>
            <xd:p>tested with Saxon-HE 9.5.0.2</xd:p>
        </xd:desc>
    </xd:doc>

    <!-- to do:
   
        still need to update the following columns to make them work correctly:

        - langmaterial / column 38 (right now, it just outputs text without any <language> elements)
        - bibliography / column 48
        - index / column 49
        
               for now, I wouldn't put any data in these columns, but if you do, the EAD will still be valid.. it just won't be useful.
               (most of this is really overkill, though; what's most useful is the ability to use Excel to create a complex hierarchy before importing
               that container list into a database like the AT or ArchivesSpace, since you can bulk edit this information within Excel much more easily 
               and quickly than you can in the AT/AS interface.)
        
        I also might want to update:
        
        - how the style sheet handles unitids (ex: auto-number for series and subseries...  add parameters to turn this off and on and to choose the enumeration type)
        - what fields are repeatable (ex: unitdates are not repeatable right now)
        - to include barcodes (with an example of how this information can be migrated via a simple SQL script after import).
        - what else?
        -->
    <xsl:key name="style-ids_match-for-color" match="ss:Style" use="@ss:ID"/>
    <!-- will probably want to change how this works, but right now you can create mixed content with the following font colors (which results in a pretty hideous rainbow):  
        #FF0000 = title
        #0070C0 = corpname
        #7030A0 = persname
        #ED7D31 = famname
        #44546A = geogname
        #00B050 = genreform     
        #00B0F0 = subject
        #FFC000 = occupation
        #FF00FF = function
        
        italics, underline, bold, etc. (all the emph options) are handled with the other font controls in Excel (e.g. bold -> emph render='bold', etc.)
      -->

    <xsl:output method="xml" indent="yes" encoding="UTF-8"/>
    <xsl:strip-space elements="*"/>
    <!--   (1 - 52 / A - AZ), columns in Excel
        1   - level number (no default..  requires at least one level-"1" value)
        2   - level type  (if no value, the level will = "file")
        3   - unitid (ex: 1 (for "Series 1").  if blank, should the transformation auto-number the series and subseries??? add paramters for whether to auto-number, roman vs. arabic numerals, etc.) [hidden] (did)
        4   - title (did)
  
        5   - date expression / make repeatable? (did)
        6   - year begin / repeatable? (did)
        7   - month begin
        8   - day begin
        9   - year end / repeatable? (did)
        10 - month end
        11 - day end
        
        12 - bulk year begin (did) [hidden]
        13 - bulk month begin
        14 - bulk day begin
        15 - bulk year end (did) [hidden]
        16 - bulk month end
        17 - bulk day end
         
        18 - instance type (use mixed materials by default / allow 14-digit barcodes?) (did)
        19 - container 1 type ("Box" is used if a value is present and the type is blank) (did)
        20 - instance 1 value (did)
        21 - container 2 type ("Folder" is used if a value is present and the type is blank) (did)
        22 - container 2 value (did)
        23 - container 3 type ("Carton" is used if a value is present and the type is blank) (did)
        24 - container 3 value (did)
        
        25 - extent number (did)
        26 - extent value, from a controlled list (did)
        27 - generic extent statement (for AT)
        28 - generic physdesc statement (allow subelements like dimensions, genreform, and physfacet???)
        
        29 - origination
        30 - bioghist note
        31 - scope and content note 
        32 - arrangement
        
        33 - access restrictions
        34 - phystech
        35 - physloc (did)
        36 - use restrictions
        37 - language code  
        38 - langmaterial (note...  language in Lime, script in Scarlet?)
        39 - other finding aid
             
        40 - custodial history <custodhist>
        41 - immediate source of aquisition <acqinfo>
        42 - appraisal 
        43 - accruals
             
        44 - location of originals <originalsloc>
        45 - alternative form available
        46 - related material
        47 - separated material 
        48 - bibliography
        49 - index
        50 - preferred citation (discourage use, since it could be automated)
        51 - process information
        52 - control access (colors for each?..
                        subject
                        geogname
                        genreform
                        etc.)
        -->
    <xsl:function name="mdc:get-column-number" as="xs:integer">
        <xsl:param name="position"/>
        <xsl:param name="current-index"/>
        <xsl:param name="previous-index"/>
        <xsl:param name="cells-before-previous-index"/>
        <xsl:value-of
            select="if ($current-index) then $current-index else
            if ($previous-index) then 
            $cells-before-previous-index + $previous-index + 1
            else $position"
        />
    </xsl:function>

    <xsl:template match="ss:Workbook">
        <ead>
            <eadheader>
                <eadid/>
                <filedesc>
                    <titlestmt>
                        <titleproper/>
                    </titlestmt>
                </filedesc>
            </eadheader>
            <archdesc level="collection">
                <did>
                    <unitid>
                        <!--AT can only accept 20 characters as the unitid, so that's exactly what the following will provide-->
                        <xsl:value-of
                            select="concat('temp', substring(string(current-dateTime()), 1, 16))"/>
                    </unitid>
                    <unitdate>undated</unitdate>
                    <unittitle>collection title</unittitle>
                    <physdesc>
                        <extent>99 Linear feet</extent>
                    </physdesc>
                    <langmaterial>
                        <language langcode="eng"/>
                    </langmaterial>
                </did>
                <!-- right now, this will only process a worksheet that has a name of "ContainerList".  if you need multiple DSCs, this would help,
                    but it might be better to change the predicate in the following XPath expression to [1], thereby ensuring a single DSC...  and if someone
                    renamed the first worksheet, it would still be processed-->
                <xsl:apply-templates select="ss:Worksheet[@ss:Name='ContainerList']/ss:Table"/>
            </archdesc>
        </ead>
    </xsl:template>

    <xsl:template match="ss:Table">
        <dsc>
            <!-- only those rows that have a level number = 1 will be processed recursively, thanks to the ss:Row template -->
            <xsl:apply-templates select="ss:Row[ss:Cell[1]/ss:Data eq '1']"/>
        </dsc>
    </xsl:template>

    <xsl:template match="ss:Row">
        <xsl:param name="depth" select="ss:Cell[1]/ss:Data" as="xs:integer"/>
        <xsl:param name="following-depth"
            select="if (following-sibling::ss:Row) 
            then following-sibling::ss:Row[1]/ss:Cell[1]/ss:Data 
            else 0"
            as="xs:integer"/>
        <xsl:param name="level"
            select="if (not(matches(ss:Cell[2]/ss:Data, '^(series|subseries|file|item)$'))) then 'file' else ss:Cell[2]/ss:Data/text()
            (: in other words, if the second column of the row is blank, then 'file' will be used as the @level type by default :)"
            as="xs:string"/>

        <!-- should I add an option to use c elements OR ennumerated components?  this would be simple to do, but it would require a slightly longer style sheet.-->
        <c level="{$level}">
            <did>
                <xsl:apply-templates mode="did"/>
            </did>
            <xsl:apply-templates mode="non-did"/>
            <xsl:if test="$following-depth eq $depth + 1">
                <xsl:apply-templates
                    select="
                    following-sibling::ss:Row[ss:Cell[1]/ss:Data[xs:integer(.)  eq $depth + 1]] 
                    except 
                    following-sibling::ss:Row[ss:Cell[1]/ss:Data[xs:integer(.) eq $depth]]/following-sibling::ss:Row"
                />
            </xsl:if>
        </c>
    </xsl:template>

    <xsl:template match="ss:Cell" mode="did">
        <xsl:param name="style-id" select="@ss:StyleID"/>
        <xsl:variable name="position" select="position()"/>
        <xsl:variable name="current-index" select="xs:integer(@ss:Index)"/>
        <xsl:variable name="previous-index"
            select="xs:integer(preceding-sibling::ss:Cell[@ss:Index][1]/@ss:Index)"/>
        <xsl:variable name="cells-before-previous-index"
            select="count(preceding-sibling::ss:Cell[@ss:Index][1]/following-sibling::* intersect preceding-sibling::ss:Cell)"/>
        <xsl:variable name="column-number" as="xs:integer">
            <xsl:sequence
                select="mdc:get-column-number($position, $current-index, $previous-index, $cells-before-previous-index)"
            />
        </xsl:variable>
        <xsl:if test="$column-number = (3, 4, 12, 20, 22, 24, 26, 28, 29, 35, 37, 38)">
            <xsl:call-template name="did-stuff">
                <xsl:with-param name="column-number" select="$column-number" as="xs:integer"/>
                <xsl:with-param name="style-id" select="$style-id"/>
            </xsl:call-template>
        </xsl:if>
        <xsl:choose>
            <!-- in other words, column number 5 must be blank, and we don't have a textual description for the unitdate-->
            <xsl:when test="@ss:Index eq '6'">
                <xsl:call-template name="did-stuff">
                    <xsl:with-param name="column-number" select="$column-number" as="xs:integer"/>
                    <xsl:with-param name="style-id" select="$style-id"/>
                </xsl:call-template>
            </xsl:when>
            <!-- in other words, column 5 isn't blank -->
            <xsl:when test="$column-number eq 5">
                <xsl:call-template name="did-stuff">
                    <xsl:with-param name="column-number" select="$column-number" as="xs:integer"/>
                    <xsl:with-param name="style-id" select="$style-id"/>
                </xsl:call-template>
            </xsl:when>
        </xsl:choose>
    </xsl:template>

    <xsl:template match="ss:Cell" mode="non-did">
        <xsl:variable name="position" select="position()"/>
        <xsl:variable name="current-index" select="xs:integer(@ss:Index)"/>
        <xsl:variable name="previous-index"
            select="xs:integer(preceding-sibling::ss:Cell[@ss:Index][1]/@ss:Index)"/>
        <xsl:variable name="cells-before-previous-index"
            select="count(preceding-sibling::ss:Cell[@ss:Index][1]/following-sibling::* intersect preceding-sibling::ss:Cell)"/>
        <xsl:variable name="column-number" as="xs:integer">
            <xsl:sequence
                select="mdc:get-column-number($position, $current-index, $previous-index, $cells-before-previous-index)"
            />
        </xsl:variable>
        <xsl:if test="$column-number = (30 to 34, 36, 39 to 52)">
            <xsl:call-template name="non-did-stuff">
                <xsl:with-param name="column-number" select="$column-number" as="xs:integer"/>
            </xsl:call-template>
        </xsl:if>
    </xsl:template>

    <xsl:template name="did-stuff">
        <xsl:param name="style-id"/>
        <xsl:param name="column-number"/>
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
                        <xsl:when test="key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color='#FF0000' 
                            and 
                            not(ss:Data/html:Font/@html:Color='#FF0000')
                            and key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Underline">
                            <title render="underline">
                                <xsl:apply-templates/>
                            </title>
                        </xsl:when>
                        <xsl:when
                            test="
                            key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color='#FF0000' 
                            and 
                            not(ss:Data/html:Font/@html:Color='#FF0000')">
                            <title render="italic">
                                <xsl:apply-templates/>
                            </title>
                        </xsl:when>
                        <xsl:when
                            test="
                            key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color='#0070C0' 
                            and 
                            not(ss:Data/html:Font/@html:Color='#0070C0')">
                            <corpname>
                                <xsl:apply-templates/>
                            </corpname>
                        </xsl:when>
                        <xsl:when
                            test="
                            key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color='#7030A0' 
                            and 
                            not(ss:Data/html:Font/@html:Color='#7030A0')">
                            <persname>
                                <xsl:apply-templates/>
                            </persname>
                        </xsl:when>
                        <xsl:when
                            test="
                            key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color='#ED7D31' 
                            and 
                            not(ss:Data/html:Font/@html:Color='#ED7D31')">
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
                <xsl:variable name="year-begin"
                    select="if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name='year_begin']) 
                    then following-sibling::ss:Cell[ss:NamedCell/@ss:Name='year_begin']/format-number(., '0000')
                    else ''"/>
                <xsl:variable name="month-begin"
                    select="if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name='month_begin'])
                    then concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name='month_begin']/format-number(., '00'))
                    else ''"/>
                <xsl:variable name="day-begin"
                    select="if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name='day_begin'])
                    then concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name='day_begin']/format-number(., '00'))
                    else ''"/>
                <xsl:variable name="year-end"
                    select="if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name='year_end'])
                    then following-sibling::ss:Cell[ss:NamedCell/@ss:Name='year_end']/format-number(., '0000')
                    else ''"/>
                <xsl:variable name="month-end"
                    select="if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name='month_end'])
                    then concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name='month_end']/format-number(., '00'))
                    else ''"/>
                <xsl:variable name="day-end"
                    select="if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name='day_end'])
                    then concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name='day_end']/format-number(., '00'))
                    else ''"/>
                <unitdate type="inclusive">
                    <xsl:attribute name="normal">
                        <xsl:choose>
                            <xsl:when
                                test="concat($year-begin, $month-begin, $day-begin) eq concat($year-end, $month-end, $day-end) 
                                or boolean($year-end) eq false()">
                                <xsl:value-of select="concat($year-begin, $month-begin, $day-begin)"
                                />
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of
                                    select="concat($year-begin, $month-begin, $day-begin, '/', $year-end, $month-end, $day-end)"
                                />
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:attribute>
                    <xsl:value-of select="."/>
                </unitdate>
            </xsl:when>
            <xsl:when test="$column-number eq 6">
                <xsl:variable name="year-begin" select="format-number(., '0000')"/>
                <xsl:variable name="month-begin"
                    select="if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name='month_begin'])
                    then concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name='month_begin']/format-number(., '00'))
                    else ''"/>
                <xsl:variable name="day-begin"
                    select="if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name='day_begin'])
                    then concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name='day_begin']/format-number(., '00'))
                    else ''"/>
                <xsl:variable name="year-end"
                    select="if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name='year_end'])
                    then following-sibling::ss:Cell[ss:NamedCell/@ss:Name='year_end']/format-number(., '0000')
                    else ''"/>
                <xsl:variable name="month-end"
                    select="if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name='month_end'])
                    then concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name='month_end']/format-number(., '00'))
                    else ''"/>
                <xsl:variable name="day-end"
                    select="if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name='day_end'])
                    then concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name='day_end']/format-number(., '00'))
                    else ''"/>
                <unitdate type="inclusive">
                    <xsl:attribute name="normal">
                        <xsl:choose>
                            <xsl:when
                                test="concat($year-begin, $month-begin, $day-begin) eq concat($year-end, $month-end, $day-end) 
                                or boolean($year-end) eq false()">
                                <xsl:value-of select="concat($year-begin, $month-begin, $day-begin)"
                                />
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of
                                    select="concat($year-begin, $month-begin, $day-begin, '/', $year-end, $month-end, $day-end)"
                                />
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:attribute>
                </unitdate>
            </xsl:when>

            <xsl:when test="$column-number eq 12">
                <xsl:variable name="bulk-year-begin" select="format-number(., '0000')"/>
                <xsl:variable name="bulk-month-begin"
                    select="if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name='bulk_month_begin'])
                    then concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name='bulk_month_begin']/format-number(., '00'))
                    else ''"/>
                <xsl:variable name="bulk-day-begin"
                    select="if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name='bulk_day_begin'])
                    then concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name='bulk_day_begin']/format-number(., '00'))
                    else ''"/>
                <xsl:variable name="bulk-year-end"
                    select="if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name='bulk_year_end'])
                    then following-sibling::ss:Cell[ss:NamedCell/@ss:Name='bulk_year_end']/format-number(., '0000')
                    else ''"/>
                <xsl:variable name="bulk-month-end"
                    select="if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name='bulk_month_end'])
                    then concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name='bulk_month_end']/format-number(., '00'))
                    else ''"/>
                <xsl:variable name="bulk-day-end"
                    select="if (following-sibling::ss:Cell[ss:NamedCell/@ss:Name='bulk_day_end'])
                    then concat('-', following-sibling::ss:Cell[ss:NamedCell/@ss:Name='bulk_day_end']/format-number(., '00'))
                    else ''"/>
                <unitdate type="bulk">
                    <xsl:attribute name="normal">
                        <xsl:choose>
                            <xsl:when
                                test="concat($bulk-year-begin, $bulk-month-begin, $bulk-day-begin) eq concat($bulk-year-end, $bulk-month-end, $bulk-day-end) 
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

            <xsl:when test="$column-number eq 20">
                <!-- label should be column 7.  If empty, though, just choose Mixed materials-->
                <container>
                    <xsl:attribute name="label">
                        <xsl:value-of
                            select="if (preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name='instance_type'])
                            then preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name='instance_type']/ss:Data
                            else 'Mixed materials'"
                        />
                    </xsl:attribute>
                    <xsl:attribute name="type">
                        <xsl:value-of
                            select="if (preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name='container_1_type'])
                            then preceding-sibling::ss:Cell[1]/ss:Data
                            else 'Box'"
                        />
                    </xsl:attribute>
                    <xsl:apply-templates/>
                </container>
            </xsl:when>

            <xsl:when test="$column-number eq 22">
                <container>
                    <xsl:attribute name="type">
                        <xsl:value-of
                            select="if (preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name='container_2_type'])
                            then preceding-sibling::ss:Cell[1]/ss:Data
                            else 'Folder'"
                        />
                    </xsl:attribute>
                    <xsl:apply-templates/>
                </container>
            </xsl:when>

            <xsl:when test="$column-number eq 24">
                <container>
                    <xsl:attribute name="type">
                        <xsl:value-of
                            select="if (preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name='container_3_type'])
                            then preceding-sibling::ss:Cell[1]/ss:Data
                            else 'Carton'"
                        />
                    </xsl:attribute>
                    <xsl:apply-templates/>
                </container>
            </xsl:when>

            <xsl:when test="$column-number eq 26">
                <physdesc>
                    <xsl:variable name="extent-number">
                        <xsl:value-of
                            select="if (preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name='extent_number'])
                            then preceding-sibling::ss:Cell[1]/ss:Data
                            else 'noextent'"
                        />
                    </xsl:variable>
                    <extent>
                        <xsl:value-of
                            select="if ($extent-number ne 'noextent') then concat($extent-number, ' ', .) else ."
                        />
                    </extent>
                    <xsl:if
                        test="following-sibling::ss:Cell[ss:NamedCell/@ss:Name='generic_extent']">
                        <extent>
                            <xsl:apply-templates
                                select="following-sibling::ss:Cell[ss:NamedCell/@ss:Name='generic_extent']"
                            />
                        </extent>
                    </xsl:if>
                </physdesc>
            </xsl:when>

            <xsl:when test="$column-number eq 28">
                <physdesc>
                    <xsl:apply-templates/>
                </physdesc>
            </xsl:when>
            <xsl:when test="$column-number eq 29">
                <origination>
                   <xsl:choose>
                       <xsl:when
                           test="
                           key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color='#0070C0' 
                           and 
                           not(ss:Data/html:Font/@html:Color='#0070C0')">
                           <corpname>
                               <xsl:apply-templates/>
                           </corpname>
                       </xsl:when>
                       <xsl:when
                           test="
                           key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color='#7030A0' 
                           and 
                           not(ss:Data/html:Font/@html:Color='#7030A0')">
                           <persname>
                               <xsl:apply-templates/>
                           </persname>
                       </xsl:when>
                       <xsl:when
                           test="
                           key('style-ids_match-for-color', $style-id)/ss:Font/@ss:Color='#ED7D31' 
                           and 
                           not(ss:Data/html:Font/@html:Color='#ED7D31')">
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
            <xsl:when test="$column-number eq 35">
                <physloc>
                    <xsl:apply-templates/>
                </physloc>
            </xsl:when>
            <xsl:when test="$column-number eq 37">
                <langmaterial>
                    <language langcode="{.}"/>
                </langmaterial>
            </xsl:when>
            <xsl:when test="$column-number eq 38">
                <langmaterial>
                    <xsl:apply-templates/>
                </langmaterial>
            </xsl:when>
        </xsl:choose>
    </xsl:template>

    <xsl:template name="non-did-stuff">
        <xsl:param name="column-number"/>
        <!-- 30 to 34, 36, 39 to 52 -->
        <xsl:variable name="element-name"
            select="if ($column-number eq 30) then 'bioghist' else if ($column-number eq 31) then 'scopecontent' 
            else if ($column-number eq 32) then 'arrangement'
            else if ($column-number eq 33) then 'accessrestrict'
            else if ($column-number eq 34) then 'phystech'
            else if ($column-number eq 36) then 'userestrict'
            else if ($column-number eq 39) then 'otherfindaid'
            else if ($column-number eq 40) then 'custodhist'
            else if ($column-number eq 41) then 'acqinfo'
            else if ($column-number eq 42) then 'appraisal'
            else if ($column-number eq 43) then 'accruals'
            else if ($column-number eq 44) then 'originalsloc'
            else if ($column-number eq 45) then 'altformavail'
            else if ($column-number eq 46) then 'relatedmaterial'
            else if ($column-number eq 47) then 'separatedmaterial'
            else if ($column-number eq 48) then 'bibliography'
            else if ($column-number eq 49) then 'index'
            else if ($column-number eq 50) then 'prefercite'
            else if ($column-number eq 51) then 'processinfo'
            else if ($column-number eq 52) then 'controlaccess'
            else 'nada'"/>
        <xsl:choose>
            <xsl:when test="$element-name eq 'nada' or normalize-space(.) eq ''"/>
            <xsl:otherwise>
                <xsl:element name="{$element-name}" namespace="urn:isbn:1-931666-22-9">
                    <xsl:apply-templates>
                        <xsl:with-param name="column-number" select="$column-number" as="xs:integer"/>
                    </xsl:apply-templates>
                </xsl:element>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>


    <xsl:template match="ss:Data">
        <xsl:param name="column-number"/>
        <xsl:choose>
            <!-- hack way to deal with adding <head> elements for scope and content and other types of notes.-->
            <xsl:when test="starts-with(*[2], '&#10;')">
                <xsl:apply-templates select="*[1]"/>
                <p>
                    <xsl:apply-templates select="node() except *[1]"/>
                </p>
            </xsl:when>
            <!-- still need to fix how we handle bibliographies, indices, and controlaccess sections.
                for now, all three are commented out.-->
            <!-- index stuff (to fix later)-->
            <xsl:when test="number($column-number) = (49)">
                <indexentry>
                    <name>
                        <xsl:apply-templates/>
                    </name>
                </indexentry>
            </xsl:when>
            
            <!-- controlaccess stuff (to fix later)-->
            <xsl:when test="number($column-number) = (52)">
                <xsl:apply-templates select="*[normalize-space()]">
                        <xsl:with-param name="column-number" select="$column-number"/>
                    </xsl:apply-templates>
            </xsl:when>
          
            
            <xsl:when test="number($column-number) = (30 to 34, 36, 39 to 48, 50, 51)">
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
    <xsl:template match="html:B[not(*)]">
        <emph render="bold">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:B[parent::html:U][not(*)]">
        <emph render="boldunderline">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:U[parent::html:B][not(*)]">
        <emph render="boldunderline">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:I[parent::html:B][not(*)]">
        <emph render="bolditalic">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:B[parent::html:I][not(*)]">
        <emph render="bolditalic">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:B[parent::html:Font/@html:Size='8'][not(*)]">
        <emph render="boldsmcaps">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:I[not(*)]">
        <emph render="italic">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:U[not(*)]">
        <emph render="underline">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:Sup">
        <emph render="super">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:Sub">
        <emph render="sub">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:Font[@html:Face='Courier New']">
        <emph render="nonproport">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>


    <xsl:template match="html:Font[@html:Size='8'][parent::html:B][not(*)]">
        <emph render="boldsmcaps">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:Font[@html:Size='8'][not(*)]">
        <emph render="smcaps">
            <xsl:apply-templates/>
        </emph>
    </xsl:template>

    <xsl:template match="html:Font[@html:Size='14']">
        <head>
            <xsl:apply-templates/>
        </head>
    </xsl:template>
    
    <xsl:template match="*:Font[@html:Color='#0070C0']">
        <corpname>
            <xsl:apply-templates/>
        </corpname>
    </xsl:template>
    <xsl:template match="*:Font[@html:Color='#7030A0']">
        <persname>
            <xsl:apply-templates/>
        </persname>
    </xsl:template>
    <xsl:template match="*:Font[@html:Color='#ED7D31']">
        <famname>
            <xsl:apply-templates/>
        </famname>
    </xsl:template>
    <xsl:template match="*:Font[@html:Color='#44546A']">
        <geogname>
            <xsl:apply-templates/>
        </geogname>
    </xsl:template>
    <xsl:template match="*:Font[@html:Color='#00B050']">
        <genreform>
            <xsl:apply-templates/>
        </genreform>
    </xsl:template>
    <xsl:template match="*:Font[@html:Color='#00B0F0']">
        <subject>
            <xsl:apply-templates/>
        </subject>
    </xsl:template>
    <xsl:template match="*:Font[@html:Color='#FFC000']">
        <occupation>
            <xsl:apply-templates/>
        </occupation>
    </xsl:template>
    <xsl:template match="*:Font[@html:Color='#FF00FF']">
        <function>
            <xsl:apply-templates/>
        </function>
    </xsl:template>
    <xsl:template match="*:Font[@html:Color='#000000']">
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

    <xsl:template match="*:Font[@html:Color='#FF0000']">
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
                <title render="italic">
                    <xsl:apply-templates/>
                </title>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <!-- I don't like doing this, but I'm not sure of a better way to create multiple paragaphs right now -->
    <xsl:template match="text()">
        <xsl:choose>
            <xsl:when test="contains(., '&#10;&#10;')">
                <xsl:for-each select="tokenize(., '&#10;&#10;')">
                    <xsl:value-of select="normalize-space(.)"/>
                    <xsl:if test="position() ne last()">
                        <xsl:text disable-output-escaping="yes">&lt;/p&gt;
                            &lt;p&gt;</xsl:text>
                    </xsl:if>
                </xsl:for-each>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="."/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
</xsl:stylesheet>
