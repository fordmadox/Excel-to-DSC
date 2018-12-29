<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:mdc="http://www.local-functions/mdc"
    xmlns:xlink="http://www.w3.org/1999/xlink"
    xmlns:ead="urn:isbn:1-931666-22-9"
    exclude-result-prefixes="xs"
    
    version="2.0">
    
    <!-- remove this entire process once we can ingest EAD dates into ASapce as expected -->
    
    <xsl:output method="xml" indent="yes" encoding="UTF-8"/>
    
    <xsl:function name="mdc:iso-date-2-display-form" as="xs:string*">
        <xsl:param name="date" as="xs:string"/>
        <xsl:variable name="months"
            select="
            ('January',
            'February',
            'March',
            'April',
            'May',
            'June',
            'July',
            'August',
            'September',
            'October',
            'November',
            'December')"/>
        <xsl:analyze-string select="$date" flags="x" regex="(\d{{4}})(\d{{2}})?(\d{{2}})?">
            <xsl:matching-substring>
                <!-- year -->
                <xsl:value-of select="regex-group(1)"/>
                <!-- month (can't add an if,then,else '' statement here without getting an extra space at the end of the result-->
                <xsl:if test="regex-group(2)">
                    <xsl:value-of select="subsequence($months, number(regex-group(2)), 1)"/>
                </xsl:if>
                <!-- day -->
                <xsl:if test="regex-group(3)">
                    <xsl:number value="regex-group(3)" format="1"/>
                </xsl:if>
                <!-- still need to handle time... but if that's there, then I can just use xs:dateTime !!!! -->
            </xsl:matching-substring>
        </xsl:analyze-string>
    </xsl:function>
    
    <xsl:template match="@*|node()">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="ead:unitdate[@type ne 'bulk'] | ead:unitdate[not(@type)]" priority="2">
        <xsl:copy>
            <xsl:copy-of select="@*"/>
            <xsl:choose>
                <xsl:when test="not(@normal) or matches(replace(., '/|-', ''), '[\D]')">
                    <xsl:apply-templates/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:variable name="first-date" select="if (contains(@normal, '/')) then replace(substring-before(@normal, '/'), '\D', '') else replace(@normal, '\D', '')"/>
                    <xsl:variable name="second-date" select="replace(substring-after(@normal, '/'), '\D', '')"/>
                    <!-- just adding the next line until i write a date conversion function-->
                    <xsl:value-of select="mdc:iso-date-2-display-form($first-date)"/>
                    <xsl:if test="$second-date ne '' and ($first-date ne $second-date)">
                        <xsl:text>&#8211;</xsl:text>
                        <xsl:value-of select="mdc:iso-date-2-display-form($second-date)"/>
                    </xsl:if>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="ead:unitdate[@type = 'bulk']" priority="2">
        <xsl:copy>
            <xsl:copy-of select="@*"/>
            <xsl:choose>
                <!-- need to convert these to human readable form if more granular than just a 4-digit year-->
                <xsl:when test="not(@normal) or matches(replace(., '/|-|bulk', ''), '[\D]')">
                    <xsl:apply-templates/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:text>Bulk, </xsl:text>
                    <xsl:variable name="first-date" select="if (contains(@normal, '/')) then replace(substring-before(@normal, '/'), '\D', '') else replace(@normal, '\D', '')"/>
                    <xsl:variable name="second-date" select="replace(substring-after(@normal, '/'), '\D', '')"/>
                    <xsl:value-of select="mdc:iso-date-2-display-form($first-date)"/>
                    <xsl:if test="$second-date ne '' and ($first-date ne $second-date)">
                        <xsl:text>&#8211;</xsl:text>
                        <xsl:value-of select="mdc:iso-date-2-display-form($second-date)"/>
                    </xsl:if>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="ead:container/@label">
        <xsl:attribute name="label">
            <xsl:value-of select="translate(., '[]', '()')"/>
        </xsl:attribute>
    </xsl:template>
    
    <!-- adjustment for ASpace imports -->
    <xsl:template match="ead:extent">
        <xsl:copy>
            <xsl:value-of select="translate(lower-case(.), 'linear feet', 'linear_feet')"/>
        </xsl:copy>
    </xsl:template>
</xsl:stylesheet>