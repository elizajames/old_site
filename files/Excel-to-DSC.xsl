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
   
    <xsl:output method="xml" indent="yes" encoding="UTF-8"/>
    <xsl:strip-space elements="*"/>
    <xsl:preserve-space elements="*:Data *:Font"/>
    
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
                if (not(matches(ss:Cell[2]/ss:Data, '^(series|subseries|file|item|accession|box)$'))) then
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
                        if ($level = 'box') then
                            'otherlevel'
                        else
                            $level"
                />
            </xsl:attribute>
            <xsl:if test="$level = 'box'">
                <xsl:attribute name="otherlevel">
                    <xsl:text>Box</xsl:text>
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
                    <xsl:apply-templates/>
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
                                    preceding-sibling::ss:Cell[ss:NamedCell/@ss:Name = 'container_1_type']/ss:Data
                                else
                                    'Box'"
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
                                    preceding-sibling::ss:Cell[1]/ss:Data
                                else
                                    'Folder'"
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
                                    preceding-sibling::ss:Cell[1]/ss:Data
                                else
                                    'Item'"
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
                 <xsl:apply-templates/>
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

            <!-- hack way to deal with adding <head> elements for scope and content and other types of notes.-->
            <!-- also gotta check style ids, since if you re-save an Excel file, it'll strip the font element out and replace it with an ID :( -->
            
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
    
    
    
     <!-- this template provides the framework for the main worksheet, including all of the column headers-->
    <xsl:template match="ead:dsc">
        <Worksheet ss:Name="ContainerList" xmlns="urn:schemas-microsoft-com:office:spreadsheet">
            <Names>
                <NamedRange ss:Name="_FilterDatabase" ss:RefersTo="=ContainerList!R1C1:R16C38"
                    ss:Hidden="1"/>
            </Names>
            <Table ss:ExpandedColumnCount="56" x:FullColumns="1"
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
                <Column ss:AutoFitWidth="0" ss:Width="180"/>
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
                    <Cell>
                        <Data ss:Type="String">system @id (leave blank, unless value already present)</Data>
                        <NamedCell ss:Name="system_id"/>
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
                <NamedRange ss:Name="system_id" ss:RefersTo="=ContainerList!C56"/>
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
       <Cell ss:Index="3"><Data ss:Type="String">Page</Data><NamedCell
      ss:Name="ContainerValues"/></Cell>
    <Cell ss:Index="4"><Data ss:Type="String">film cartridges</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">ain - Ainu</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
        <Cell ss:Index="3"><Data ss:Type="String">Map Case</Data><NamedCell
      ss:Name="ContainerValues"/></Cell>
    <Cell ss:Index="4"><Data ss:Type="String">film cassettes</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">aka - Akan</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="14.4375">
       <Cell ss:Index="3"><Data ss:Type="String">Drawer</Data><NamedCell
      ss:Name="ContainerValues"/></Cell>
    <Cell ss:Index="4"><Data ss:Type="String">film loops</Data><NamedCell
      ss:Name="ExtentValues"/></Cell>
    <Cell><Data ss:Type="String">akk - Akkadian</Data><NamedCell
      ss:Name="LanguageCodes"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
       <Cell ss:Index="3"><Data ss:Type="String">item_barcode</Data><NamedCell
      ss:Name="ContainerValues"/></Cell>
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
