<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
    exclude-result-prefixes="office meta">

    <xsl:template name="docprops-core">
        <cp:coreProperties
            xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
            xmlns:dcmitype="http://purl.org/dc/dcmitype/">
            <xsl:apply-templates select="document('meta.xml')/office:document-meta/office:meta"
                mode="core"/>
        </cp:coreProperties>
    </xsl:template>

    <xsl:template name="docprops-custom">
        <xsl:apply-templates select="document('meta.xml')/office:document-meta/office:meta"
            mode="custom"/>
    </xsl:template>

    <xsl:template match="/office:document-meta/office:meta" mode="core">
        <!-- report lost properties -->
        <xsl:apply-templates select="meta:auto-reload" mode="core"/>
        <xsl:apply-templates select="meta:hyperlink-behaviour" mode="core"/>
        <!-- creation date -->
        <xsl:apply-templates select="meta:creation-date" mode="core"/>
        <!-- creator -->
        <xsl:apply-templates select="meta:initial-creator" mode="core"/>
        <!-- description -->
        <xsl:apply-templates select="dc:description" mode="core"/>
        <!-- identifier -->
        <xsl:apply-templates select="dc:identifier" mode="core"/>
        <!-- keywords -->
        <xsl:call-template name="MetaKeywords"/>
        <!-- language -->
        <xsl:apply-templates select="dc:language" mode="core"/>
        <!-- last modification author -->
        <xsl:apply-templates select="dc:creator" mode="core"/>
        <!-- last printing -->
        <xsl:apply-templates select="meta:printed-date" mode="core"/>
        <!-- last modification date -->
        <xsl:apply-templates select="dc:date" mode="core"/>
        <!-- number of times it was saved -->
        <xsl:apply-templates select="meta:editing-cycles" mode="core"/>
        <!-- topic -->
        <xsl:apply-templates select="dc:subject" mode="core"/>
        <!-- title -->
        <xsl:apply-templates select="dc:title" mode="core"/>
    </xsl:template>

    <xsl:template match="/office:document-meta/office:meta" mode="custom">
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
            <xsl:apply-templates select="meta:user-defined"/>
        </Properties>
    </xsl:template>

    <!-- creation date -->
    <xsl:template match="meta:creation-date" mode="core">
        <xsl:variable name="dateIsValid">
            <xsl:call-template name="validateDate">
                <xsl:with-param name="date">
                    <xsl:value-of select="."/>
                </xsl:with-param>
            </xsl:call-template>
        </xsl:variable>
        <xsl:choose>
            <xsl:when test="$dateIsValid != 'false' ">
                <dcterms:created xmlns:dcterms="http://purl.org/dc/terms/"
                    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="dcterms:W3CDTF">
                    <xsl:value-of select="$dateIsValid"/>
                </dcterms:created>
            </xsl:when>
            <xsl:otherwise>
                <xsl:message terminate="no">translation.odf2oox.documentCreationDate</xsl:message>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <!-- initial creator -->
    <xsl:template match="meta:initial-creator" mode="core">
        <dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/">
            <xsl:value-of select="."/>
        </dc:creator>
    </xsl:template>

    <!-- description -->
    <xsl:template match="dc:description" mode="core">
        <dc:description xmlns:dc="http://purl.org/dc/elements/1.1/">
            <xsl:value-of select="."/>
        </dc:description>
    </xsl:template>

    <!-- identifier -->
    <xsl:template match="dc:identifier" mode="core">
        <dc:identifier xmlns:dc="http://purl.org/dc/elements/1.1/">
            <xsl:value-of select="."/>
        </dc:identifier>
    </xsl:template>

    <!-- keywords -->
    <xsl:template name="MetaKeywords">
        <xsl:if test="/office:document-meta/office:meta/meta:keyword">
            <cp:keywords
                xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
                <xsl:for-each select="/office:document-meta/office:meta/meta:keyword">
                    <xsl:if test="not(position() = 1)">
                        <xsl:text> </xsl:text>
                    </xsl:if>
                    <xsl:value-of select="."/>
                </xsl:for-each>
            </cp:keywords>
        </xsl:if>
    </xsl:template>

    <!-- language -->
    <xsl:template match="dc:language" mode="core">
        <dc:language xmlns:dc="http://purl.org/dc/elements/1.1/">
            <xsl:value-of select="."/>
        </dc:language>
    </xsl:template>

    <!-- last modification author -->
    <xsl:template match="dc:creator" mode="core">
        <cp:lastModifiedBy
            xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
            <xsl:value-of select="."/>
        </cp:lastModifiedBy>
    </xsl:template>

    <!-- last printing -->
    <xsl:template match="meta:printed-date" mode="core">
        <cp:lastPrinted
            xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
            <xsl:value-of select="."/>
        </cp:lastPrinted>
    </xsl:template>

    <!-- last modification date -->
    <xsl:template match="dc:date" mode="core">
        <xsl:variable name="dateIsValid">
            <xsl:call-template name="validateDate">
                <xsl:with-param name="date">
                    <xsl:value-of select="."/>
                </xsl:with-param>
            </xsl:call-template>
        </xsl:variable>
        <xsl:choose>
            <xsl:when test="$dateIsValid != 'false' ">
                <dcterms:modified xmlns:dcterms="http://purl.org/dc/terms/"
                    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="dcterms:W3CDTF">
                    <xsl:value-of select="$dateIsValid"/>
                </dcterms:modified>
            </xsl:when>
            <xsl:otherwise>
                <xsl:message terminate="no"
                >translation.odf2oox.documentModificationDate</xsl:message>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <!-- number of times it was saved -->
    <xsl:template match="meta:editing-cycles" mode="core">
        <cp:revision
            xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
            <xsl:value-of select="."/>
        </cp:revision>
    </xsl:template>

    <!-- topic -->
    <xsl:template match="dc:subject" mode="core">
        <dc:subject xmlns:dc="http://purl.org/dc/elements/1.1/">
            <xsl:value-of select="."/>
        </dc:subject>
    </xsl:template>

    <!-- title -->
    <xsl:template match="dc:title" mode="core">
        <dc:title xmlns:dc="http://purl.org/dc/elements/1.1/">
            <xsl:value-of select="."/>
        </dc:title>
    </xsl:template>

    <!-- report lost properties -->
    <xsl:template match="meta:auto-reload" mode="core">
        <xsl:message terminate="no">translation.odf2oox.internetProperties</xsl:message>
    </xsl:template>

    <xsl:template match="meta:hyperlink-behaviour" mode="core">
        <xsl:message terminate="no">translation.odf2oox.internetProperties</xsl:message>
    </xsl:template>

    <!-- user defined properties -->
    <xsl:template match="/office:document-meta/office:meta/meta:user-defined">
        <property xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
            <xsl:attribute name="fmtid">{D5CDD505-2E9C-101B-9397-08002B2CF9AE}</xsl:attribute>
            <xsl:attribute name="pid">
                <xsl:value-of select="position() + 1"/>
            </xsl:attribute>
            <xsl:attribute name="name">
                <xsl:value-of select="@meta:name"/>
            </xsl:attribute>
            <vt:lpwstr>
                <xsl:value-of select="text()"/>
            </vt:lpwstr>
        </property>
    </xsl:template>

    <!-- page count statistics extended property -->
    <xsl:template match="@meta:page-count">
        <Pages xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
            <xsl:value-of select="."/>
        </Pages>
    </xsl:template>

    <!-- word count statistics extended property -->
    <xsl:template match="@meta:word-count">
        <Words xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
            <xsl:value-of select="."/>
        </Words>
    </xsl:template>

    <!-- application extended property -->
    <xsl:template name="GetApplicationExtendedProperty">
        <Application
            xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">ODF
            Converter</Application>
    </xsl:template>

    <!-- document security extended property -->
    <xsl:template name="GetDocSecurityExtendedProperty">
        <DocSecurity
            xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
        >0</DocSecurity>
    </xsl:template>

    <!-- paragraphs statistics extended property -->
    <xsl:template match="@meta:paragraph-count">
        <Paragraphs
            xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
            <xsl:value-of select="."/>
        </Paragraphs>
    </xsl:template>


    <!--  editing duration -->
    <xsl:template match="meta:editing-duration">
        <xsl:variable name="hours">
            <xsl:choose>
                <xsl:when test="contains(., 'H')">
                    <xsl:value-of select="substring-before(substring-after(., 'T'), 'H')"/>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:variable name="minutes">
            <xsl:choose>
                <xsl:when test="contains(., 'H')">
                    <xsl:value-of select="substring-before(substring-after(., 'H'), 'M')"/>
                </xsl:when>
                <xsl:when test="contains(., 'M')">
                    <xsl:value-of select="substring-before(substring-after(., 'T'), 'M')"/>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:variable name="seconds">
            <xsl:choose>
                <xsl:when test="contains(., 'M')">
                    <xsl:value-of select="substring-before(substring-after(., 'M'), 'S')"/>
                </xsl:when>
                <xsl:when test="contains(., 'S')">
                    <xsl:value-of select="substring-before(substring-after(., 'T'), 'S')"/>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:if test="number($hours) and number($minutes) and number($seconds)">
            <TotalTime
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
                <xsl:choose>
                    <xsl:when test="number($seconds) &gt; 30">
                        <xsl:value-of select="60 * number($hours) + number($minutes) + 1"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="60 * number($hours) + number($minutes)"/>
                    </xsl:otherwise>
                </xsl:choose>
            </TotalTime>
        </xsl:if>
    </xsl:template>


    <!-- characters with spaces statistics -->
    <xsl:template match="@meta:character-count">
        <CharactersWithSpaces
            xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
            <xsl:value-of select="."/>
        </CharactersWithSpaces>
    </xsl:template>

    <!-- non whitespace character count statistics -->
    <xsl:template match="@meta:non-whitespace-character-count">
        <Characters
            xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
            <xsl:value-of select="meta:document-statistic/@meta:non-whitespace-character-count"/>
        </Characters>
    </xsl:template>


    <!-- check if a date is valid regarding W3C format -->
    <xsl:template name="validateDate">
        <xsl:param name="date"/>
        <!-- year -->
        <xsl:variable name="Y">
            <xsl:choose>
                <xsl:when test="contains($date, '-')">
                    <xsl:value-of select="substring-before($date, '-')"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="$date"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <!-- month -->
        <xsl:variable name="M">
            <xsl:variable name="date_Y" select="substring-after($date, concat($Y, '-'))"/>
            <xsl:choose>
                <xsl:when test="contains($date_Y, '-')">
                    <xsl:value-of select="substring-before($date_Y, '-')"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="$date_Y"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <!-- day -->
        <xsl:variable name="D">
            <xsl:variable name="date_Y_m" select="substring-after($date, concat($Y, '-', $M, '-'))"/>
            <xsl:choose>
                <xsl:when test="contains($date_Y_m, 'T')">
                    <xsl:value-of select="substring-before($date_Y_m, 'T')"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="$date_Y_m"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <!-- hour -->
        <xsl:variable name="h">
            <xsl:if test="contains($date, 'T')">
                <xsl:value-of select="substring-before(substring-after($date, 'T'), ':')"/>
            </xsl:if>
        </xsl:variable>
        <!-- minutes -->
        <xsl:variable name="m">
            <xsl:if test="contains($date, 'T')">
                <xsl:variable name="time_h"
                    select="substring-after(substring-after($date, 'T'), ':')"/>
                <xsl:choose>
                    <xsl:when test="contains($time_h, ':')">
                        <xsl:value-of select="substring-before($time_h, ':')"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="$time_h"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:if>
        </xsl:variable>
        <!-- seconds (with our without decimal) -->
        <xsl:variable name="s">
            <xsl:if test="contains($date, 'T')">
                <xsl:variable name="time_h_m"
                    select="substring-after(substring-after($date, 'T'), concat($h, ':', $m, ':'))"/>
                <xsl:choose>
                    <xsl:when test="contains($time_h_m, 'Z')">
                        <xsl:value-of select="substring-before($time_h_m, 'Z')"/>
                    </xsl:when>
                    <xsl:when test="contains($time_h_m, '+')">
                        <xsl:value-of select="substring-before($time_h_m, '+')"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="$time_h_m"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:if>
        </xsl:variable>

        <!-- validate format -->
        <xsl:choose>
            <xsl:when test="string-length($Y) != 4">false</xsl:when>
            <xsl:when test="$M != '' and string-length($M) != 2">false</xsl:when>
            <xsl:when test="$D != '' and string-length($D) != 2">false</xsl:when>
            <xsl:when test="$h != '' and string-length($h) != 2">false</xsl:when>
            <xsl:when test="$m != '' and string-length($m) != 2">false</xsl:when>
            <xsl:when test="$s != '' and string-length($s) &gt; 4">false</xsl:when>
            <xsl:when test="string-length(substring-after($date, 'Z')) &gt; 0">false</xsl:when>
            <xsl:when test="$M != '' and number($M) &gt; 12">false</xsl:when>
            <xsl:when test="$M != '' and number($M) &lt; 1">false</xsl:when>
            <xsl:when test="$D != '' and number($D) &gt; 31">false</xsl:when>
            <xsl:when test="$D != '' and number($D) &lt; 1">false</xsl:when>
            <xsl:when test="$h != '' and number($h) &gt; 24">false</xsl:when>
            <xsl:when test="$h != '' and number($h) &lt; 1">false</xsl:when>
            <xsl:when test="$m != '' and number($m) &gt; 60">false</xsl:when>
            <xsl:when test="$m != '' and number($m) &lt; 1">false</xsl:when>
            <xsl:when test="$s != '' and number($s) &gt; 61">false</xsl:when>
            <xsl:when test="$s != '' and number($s) &lt; 1">false</xsl:when>
            <xsl:when
                test="contains($date, 'T') and not(contains($date, 'Z') or contains($date, '+'))">
                <xsl:value-of select="concat($date, 'Z')"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="$date"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

</xsl:stylesheet>
